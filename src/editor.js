// Editing behaviour on top of the rendered document.
//
// The page is a single contenteditable host so selections can span paragraphs. The browser is
// left to handle edits inside one paragraph (typing, Ctrl+B, ...). Anything that would change the
// paragraph structure (Enter, deleting at a paragraph edge, any edit of a selection spanning
// paragraphs) is intercepted and done by hand, so every HTML paragraph keeps exactly one <w:p>.
//
// Undo is ours, not the browser's: the native stack can't follow paragraphs we split, merge or
// re-render, so every edit is recorded in a History and Ctrl+Z / Ctrl+Y are routed to it.
import { loadDocx, saveDocx, disposeDocx } from './docx/package.js';
import { createModel, renderDocument, renderParagraph, renumber, lastLeaf } from './render.js';
import { serializeParagraph } from './serialize.js';
import { History } from './history.js';
import { caretIn, placeCaret } from './caret.js';
import { caretRects, onEdgeLine, lineMoveStaysIn, verticalTarget, horizontalTarget, placeAtColumn } from './navigate.js';
import { wChild, wEl, putInOrder, PPR_ORDER } from './docx/xml.js';

// Consecutive typing in one paragraph is one undo step unless paused for this long.
const COALESCE_MS = 1000;
const NAVIGATION_KEYS = new Set(['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End', 'PageUp', 'PageDown']);
const DELETE_BACKWARD = new Set(['deleteContentBackward', 'deleteWordBackward', 'deleteSoftLineBackward', 'deleteHardLineBackward']);
const DELETE_FORWARD = new Set(['deleteContentForward', 'deleteWordForward', 'deleteSoftLineForward', 'deleteHardLineForward']);

const hasSectionBreak = (p) => !!wChild(wChild(p, 'pPr'), 'sectPr');
const isPara = (el) => el.classList.contains('para');

function editableParaOf(node) {
  const el = (node?.nodeType === Node.ELEMENT_NODE ? node : node?.parentElement)?.closest('.para');
  return el && !el.classList.contains('locked') && el.isContentEditable ? el : null;
}

/** True if nothing but the placeholder <br> lies between `el`'s edge and the given point. */
function blankBetween(el, edge, container, offset) {
  const probe = document.createRange();
  if (edge === 'start') {
    probe.setStart(el, 0);
    probe.setEnd(container, offset);
  } else {
    probe.setStart(container, offset);
    probe.setEnd(el, el.childNodes.length);
  }
  const frag = probe.cloneContents();
  if (frag.textContent !== '' || frag.querySelector('[data-atom]')) return false;
  const last = lastLeaf(el);
  const placeholder = last?.nodeName === 'BR' && probe.intersectsNode(last) ? 1 : 0;
  return frag.querySelectorAll('br').length <= placeholder;
}

/** True if the caret is collapsed at the very start/end of `el`. */
function caretAtEdge(el, edge) {
  const sel = getSelection();
  if (!sel.rangeCount || !sel.isCollapsed) return false;
  const { startContainer, startOffset } = sel.getRangeAt(0);
  return blankBetween(el, edge, startContainer, startOffset);
}

/**
 * The selection as a Range, adjusted so its ends sit in paragraphs where possible (select-all can
 * report container offsets) and never inside an atom (atoms are deleted whole or not at all).
 */
function selectionRange() {
  const sel = getSelection();
  if (!sel.rangeCount) return null;
  const range = sel.getRangeAt(0).cloneRange();
  const { startContainer: sc, startOffset: so, endContainer: ec, endOffset: eo } = range;
  if (sc.nodeType === Node.ELEMENT_NODE && !sc.closest('.para') && sc.childNodes[so]?.classList?.contains('para')) {
    range.setStart(sc.childNodes[so], 0);
  }
  if (ec.nodeType === Node.ELEMENT_NODE && !ec.closest('.para') && ec.childNodes[eo - 1]?.classList?.contains('para')) {
    const para = ec.childNodes[eo - 1];
    range.setEnd(para, para.childNodes.length);
  }
  const atomOf = (n) => (n.nodeType === Node.ELEMENT_NODE ? n : n.parentElement)?.closest('[data-atom]');
  const startAtom = atomOf(range.startContainer);
  if (startAtom) range.setStartBefore(startAtom);
  const endAtom = atomOf(range.endContainer);
  if (endAtom) range.setEndAfter(endAtom);
  return range;
}

/** Plain text of a multi-paragraph selection, one line per block (paste splits on lines). */
function selectedText({ range, blocks }) {
  return blocks.map((b, i) => {
    const part = document.createRange();
    part.selectNodeContents(b);
    if (i === 0) part.setStart(range.startContainer, range.startOffset);
    if (i === blocks.length - 1) part.setEnd(range.endContainer, range.endOffset);
    return part.toString();
  }).join('\n');
}

// Invariant: a <br> that is the last leaf of a paragraph is a placeholder (as browsers use it to
// give an empty line height) and is not saved. These helpers restore that after DOM surgery.

/** Give an empty paragraph its placeholder so it stays one line tall and can hold the caret. */
function fillIfEmpty(el) {
  if (!lastLeaf(el)) el.append(document.createElement('br'));
}

/** After cutting a paragraph's tail off, a trailing <br> is a real line break: protect it. */
function protectTrailingBreak(el) {
  const last = lastLeaf(el);
  if (!last || last.nodeName === 'BR') el.append(document.createElement('br'));
}

function editKind(inputType) {
  if (inputType.startsWith('delete')) return 'delete';
  if (inputType.startsWith('format')) return 'format';
  return 'insert';
}

export class Editor {
  constructor(host, onChange) {
    this.host = host;
    this.onChange = onChange;
    this.model = null;
    this.page = null;
    this.current = null; // paragraph the selection is (or was last) in
    this.unsaved = false;
    this.notice = null; // message for the user about an edit we refused
    this.history = new History((from, to) => this.#swap(from, to));
    this.recording = true; // false while a compound edit is already being recorded
    this.pendingDirty = null; // paragraphs a native edit is about to change
    this.goalX = null; // column kept across consecutive Up/Down presses (like the browser's own)
    this.goalAt = null; // where the last Up/Down press left the caret

    host.addEventListener('input', () => this.#onInput());
    host.addEventListener('keydown', (e) => this.#onKeyDown(e));
    host.addEventListener('beforeinput', (e) => this.#onBeforeInput(e));
    host.addEventListener('paste', (e) => this.#onPaste(e));
    host.addEventListener('cut', (e) => this.#onCut(e));
    host.addEventListener('compositionstart', () => this.#onCompositionStart());
    host.addEventListener('drop', (e) => e.preventDefault());
    host.addEventListener('pointerdown', () => {
      this.history.seal();
      this.goalX = null;
    });
    host.addEventListener('keyup', (e) => {
      if (e.key === 'ArrowUp' || e.key === 'ArrowDown') this.goalAt = this.#caretAt();
    });
    document.addEventListener('keydown', (e) => this.#onShortcut(e));
    document.addEventListener('selectionchange', () => this.#onSelectionChange());
  }

  async open(data) {
    const pkg = await loadDocx(data);
    if (this.model) disposeDocx(this.model.pkg);
    this.model = createModel(pkg);
    this.page = renderDocument(this.model);
    this.host.replaceChildren(this.page);
    this.current = null;
    this.unsaved = false;
    this.history.clear();
    this.#emit();
  }

  /** Write all pending paragraph edits into the XML and return the .docx as a Blob. */
  async save() {
    for (const el of this.model.dirty) this.#commit(el);
    const blob = await saveDocx(this.model.pkg, { modified: this.model.modified });
    this.unsaved = false;
    this.#emit();
    return blob;
  }

  undo() {
    if (this.model && this.history.undo()) this.#changed();
  }

  redo() {
    if (this.model && this.history.redo()) this.#changed();
  }

  /** Apply a browser formatting command (bold/italic/underline) to the selection. */
  format(command) {
    if (this.model && this.#noteFormat()) this.#unrecorded(() => document.execCommand(command));
  }

  /** Set the paragraph style of the selected paragraphs (or the last one the caret was in). */
  setParagraphStyle(styleId) {
    if (!this.model) return;
    let els = null;
    const sel = getSelection();
    if (sel.rangeCount && this.page.contains(sel.anchorNode)) {
      const plan = this.#plan({ deleting: false, parasOnly: true });
      if ('error' in plan) return this.#notify(plan.error);
      els = plan.blocks;
    } else if (this.current?.isConnected && editableParaOf(this.current)) {
      els = [this.current];
    }
    if (!els || els.every((el) => el.dataset.style === styleId)) return;

    const { model } = this;
    const entry = this.#beginEdit(els);
    for (const el of els) {
      const p = model.paraOf.get(el);
      let pPr = wChild(p, 'pPr');
      if (!pPr) p.prepend((pPr = wEl(p.ownerDocument, 'pPr')));
      if (styleId === model.styles.defaultParaStyle) wChild(pPr, 'pStyle')?.remove();
      else putInOrder(pPr, wEl(p.ownerDocument, 'pStyle', { val: styleId }), PPR_ORDER);
      if (!pPr.children.length) pPr.remove();
      renderParagraph(model, p, el);
    }
    model.modified = true;
    renumber(model, this.page);
    this.current = placeCaret(els, entry.before.caret);
    this.#record(entry, els);
  }

  get state() {
    const { model } = this;
    if (!model) return { loaded: false };
    const paras = [...this.page.querySelectorAll('.para')];
    const used = new Set(paras.map((el) => el.dataset.style));
    return {
      loaded: true,
      unsaved: this.unsaved,
      canUndo: this.history.canUndo,
      canRedo: this.history.canRedo,
      paragraphs: paras.length,
      locked: paras.filter((el) => el.classList.contains('locked')).length,
      styles: model.styles.paragraphStyles(used),
      currentStyle: editableParaOf(this.current) ? this.current.dataset.style : null,
      notice: this.notice,
    };
  }

  #emit() {
    this.onChange(this.state);
  }

  #changed() {
    this.unsaved = true;
    this.#emit();
  }

  #notify(message) {
    this.notice = message;
    this.#emit();
    this.notice = null;
  }

  #xmlOf(el) {
    return isPara(el) ? this.model.paraOf.get(el) : this.model.blockOf.get(el);
  }

  /** Write a paragraph's pending edits into its <w:p>. */
  #commit(el) {
    if (!this.model.dirty.delete(el)) return;
    serializeParagraph(this.model, el);
    this.model.modified = true;
  }

  /** Commit a block: the paragraph itself, or every paragraph inside a table etc. */
  #commitBlock(el) {
    for (const p of isPara(el) ? [el] : el.querySelectorAll('.para')) this.#commit(p);
  }

  /**
   * Snapshot blocks before an edit changes them. Paragraphs are copied, since edits change them in
   * place; other blocks are only ever removed whole, so the originals are kept. Finish with #record.
   */
  #beginEdit(blocks, caret = caretIn(blocks)) {
    for (const el of blocks) this.#commitBlock(el);
    return {
      before: {
        els: blocks,
        ps: blocks.map((el) => (isPara(el) ? this.#xmlOf(el).cloneNode(true) : this.#xmlOf(el))),
        caret,
      },
      after: { els: null, ps: null, caret: null },
      open: false,
    };
  }

  /** Push an entry from #beginEdit, now that the edit has produced blocks `els`. */
  #record(entry, els) {
    entry.after.els = els;
    this.history.push(entry);
    this.#changed();
  }

  /** Take side `from` off screen (keeping its XML) and put side `to` in its place. */
  #swap(from, to) {
    const { model } = this;
    from.caret = caretIn(from.els);
    for (const el of from.els) this.#commitBlock(el);
    from.ps = from.els.map((el) => this.#xmlOf(el));

    from.ps[0].before(...to.ps);
    for (const p of from.ps) p.remove();

    const slot = document.createComment('');
    from.els[0].before(slot);
    for (const el of from.els) el.remove();
    to.els.forEach((el, i) => {
      if (isPara(el)) renderParagraph(model, to.ps[i], el);
    });
    slot.replaceWith(...to.els);
    to.ps = null;

    model.modified = true;
    renumber(model, this.page);
    this.current = placeCaret(to.els, to.caret);
  }

  /**
   * What an edit of the selection would touch: {start, end, blocks, cross, range}, where blocks
   * runs from the start paragraph to the end paragraph, or {error} if it can't be edited as a whole.
   */
  #plan({ deleting = true, parasOnly = false } = {}) {
    const range = selectionRange();
    const start = range && editableParaOf(range.startContainer);
    const end = range && editableParaOf(range.endContainer);
    if (start && start === end) return { range, start, end, blocks: [start], cross: false };
    if (!start || !end || !this.page.contains(start) || !this.page.contains(end)) {
      return { error: 'The selection must start and end in editable text.' };
    }
    if (start.parentElement !== end.parentElement) {
      return { error: 'Editing across table cells or content controls isn’t supported.' };
    }
    const blocks = [start];
    for (let b = start.nextElementSibling; b !== end; b = b.nextElementSibling) {
      if (!b) return { error: 'This selection can’t be edited as a whole.' };
      blocks.push(b);
    }
    blocks.push(end);
    if (blocks.some((b) => b.classList.contains('locked'))) {
      return { error: 'The selection includes read-only paragraphs.' };
    }
    if (parasOnly && !blocks.every(isPara)) {
      return { error: 'The selection includes a table or other non-text content.' };
    }
    if (deleting && blocks.some((b) => isPara(b) && hasSectionBreak(this.#xmlOf(b)))) {
      return { error: 'Deleting across a section break isn’t supported.' };
    }
    // Blocks flattened out of a wrapper (e.g. block-level custom XML) can't be merged with others.
    const parents = new Set(blocks.map((b) => this.#xmlOf(b)?.parentNode));
    if (parents.size !== 1 || parents.has(undefined)) {
      return { error: 'This selection can’t be edited as a whole.' };
    }
    return { range, start, end, blocks, cross: true };
  }

  /**
   * Run `fn` as one undo step on the selection. A selection spanning paragraphs is deleted first.
   * `fn(el)` gets the paragraph holding the caret and returns the paragraphs the edit leaves.
   * Returns false (after telling the user) if the selection can't be edited.
   */
  #edit(fn = (el) => [el], { caret } = {}) {
    const plan = this.#plan();
    if ('error' in plan) {
      this.#notify(plan.error);
      return false;
    }
    const entry = this.#beginEdit(plan.blocks, caret);
    this.#unrecorded(() => {
      const el = plan.cross ? this.#deleteAcross(plan) : plan.start;
      this.#record(entry, fn(el));
    });
    return true;
  }

  #unrecorded(fn) {
    this.recording = false;
    try {
      return fn();
    } finally {
      this.recording = true;
    }
  }

  /** Record a typing/formatting step in `el`, merging it into the previous one where natural. */
  #noteEdit(el, kind) {
    if (!this.recording) return;
    const top = this.history.top;
    const now = performance.now();
    if (top?.open && top.kind === kind && top.after.els.length === 1 && top.after.els[0] === el &&
        now - top.time < COALESCE_MS) {
      top.time = now;
      return;
    }
    const entry = this.#beginEdit([el]);
    Object.assign(entry, { kind, time: now, open: kind !== 'format' });
    this.#record(entry, [el]);
  }

  /** Record a formatting change to the selection, which the browser then applies. */
  #noteFormat() {
    const plan = this.#plan({ deleting: false, parasOnly: true });
    if ('error' in plan) {
      this.#notify(plan.error);
      return false;
    }
    if (plan.cross) this.#record(this.#beginEdit(plan.blocks), plan.blocks);
    else this.#noteEdit(plan.start, 'format');
    this.pendingDirty = plan.blocks;
    return true;
  }

  /** Insert text (or a line break) at the selection, replacing it. */
  #insert(command, value) {
    const plan = this.#plan();
    if ('error' in plan) return this.#notify(plan.error);
    if (plan.cross) {
      this.#edit((el) => {
        document.execCommand(command, false, value);
        return [el];
      });
    } else {
      this.#noteEdit(plan.start, 'insert');
      this.#unrecorded(() => document.execCommand(command, false, value));
    }
  }

  #onInput() {
    const els = this.pendingDirty ?? [editableParaOf(getSelection().anchorNode)];
    this.pendingDirty = null;
    this.goalX = null;
    for (const el of els) if (el?.isConnected) this.model.dirty.add(el);
    this.#changed();
  }

  #onSelectionChange() {
    const el = editableParaOf(getSelection().anchorNode);
    if (el && el !== this.current && this.page?.contains(el)) {
      this.current = el;
      this.#emit();
    }
  }

  #onShortcut(e) {
    if (!this.model || !(e.ctrlKey || e.metaKey) || e.altKey) return;
    if (e.target.closest?.('input, textarea, select')) return;
    const key = e.key.toLowerCase();
    if (key === 'z' && !e.shiftKey) {
      e.preventDefault();
      this.undo();
    } else if ((key === 'z' && e.shiftKey) || (key === 'y' && !e.shiftKey)) {
      e.preventDefault();
      this.redo();
    }
  }

  #onKeyDown(e) {
    if (!this.model || e.isComposing) return;
    const vertical = (e.key === 'ArrowUp' || e.key === 'ArrowDown') && !e.ctrlKey && !e.metaKey && !e.altKey && !e.shiftKey;
    // The column is forgotten once the caret moves any other way (a key, a click, a script).
    if (!vertical || (!e.repeat && !this.#caretAt(this.goalAt))) this.goalX = null;
    if (NAVIGATION_KEYS.has(e.key)) {
      this.history.seal();
      if (e.key.startsWith('Arrow') && !e.metaKey && !e.altKey && !e.shiftKey) this.#arrow(e, vertical);
      return;
    }
    if (e.ctrlKey || e.metaKey || e.altKey) return;
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      this.#edit((el) => [el, this.#splitAtCaret(el)]);
    } else if (e.key === 'Tab') {
      e.preventDefault();
      this.#insert('insertText', '\t');
    }
  }

  /** With no argument, the caret position; with one, whether the caret is (still) there. */
  #caretAt(pos) {
    const sel = getSelection();
    const now = sel.rangeCount ? [sel.focusNode, sel.focusOffset] : null;
    if (pos === undefined) return now;
    return !!pos && !!now && pos[0] === now[0] && pos[1] === now[1];
  }

  /**
   * Arrow keys at a paragraph's edge, where the browser would skip or mangle a table cell or a
   * read-only paragraph. Moves within a paragraph, or to the next paragraph in the same editing
   * host, stay native. Ctrl+Left/Right are included: at an edge they'd escape the same way.
   */
  #arrow(e, vertical) {
    if (!vertical && e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') return; // Ctrl+Up/Down
    const sel = getSelection();
    if (!sel.rangeCount || !sel.isCollapsed) return;
    const para = editableParaOf(sel.anchorNode);
    if (!para) return;
    const dir = e.key === 'ArrowUp' || e.key === 'ArrowLeft' ? -1 : 1;
    let target;
    let probed = false; // the selection was reset, so the native move would start from the wrong line
    if (vertical) {
      const rects = caretRects(sel.getRangeAt(0));
      const edges = rects.filter((r) => onEdgeLine(para, r, dir));
      if (edges.length < rects.length) {
        if (!edges.length) {
          this.goalX ??= rects.at(-1).left;
          return;
        }
        const inner = rects.find((r) => !edges.includes(r));
        // At a soft wrap, with one side on the edge line: let the browser, which knows, decide.
        probed = true;
        if (lineMoveStaysIn(para, dir)) {
          this.goalX ??= inner.left;
          e.preventDefault();
          return;
        }
      }
      this.goalX ??= edges[0].left;
      target = verticalTarget(para, dir, this.goalX);
    } else {
      if (!caretAtEdge(para, dir < 0 ? 'start' : 'end')) return;
      target = horizontalTarget(this.page, para, dir);
    }
    if (!target || (!probed && target === (dir < 0 ? para.previousElementSibling : para.nextElementSibling))) return;
    e.preventDefault();
    if (vertical) placeAtColumn(target, this.goalX, dir);
    else placeCaret([target], dir < 0 ? { index: 0, offset: Infinity } : undefined);
  }

  #onBeforeInput(e) {
    const type = e.inputType;
    if (type === 'historyUndo' || type === 'historyRedo') {
      // Edit menu / context menu undo: never let the native stack run.
      e.preventDefault();
      if (type === 'historyUndo') this.undo();
      else this.redo();
      return;
    }
    if (type === 'insertFromDrop' || type === 'deleteByDrag') {
      e.preventDefault();
      return;
    }
    if (type.startsWith('format')) {
      if (!this.#noteFormat()) e.preventDefault();
      return;
    }

    const plan = this.#plan();
    if ('error' in plan) {
      e.preventDefault();
      this.#notify(plan.error);
      return;
    }
    if (plan.cross || type === 'insertParagraph') {
      // The browser would restructure the paragraphs its own way; do it ourselves.
      e.preventDefault();
      if (type === 'insertParagraph') this.#edit((el) => [el, this.#splitAtCaret(el)]);
      else if (type === 'insertLineBreak') this.#insert('insertLineBreak');
      else if (type.startsWith('insert') && e.data) this.#insert('insertText', e.data);
      else this.#edit();
      return;
    }
    const el = plan.start;
    const backward = DELETE_BACKWARD.has(type);
    if ((backward && caretAtEdge(el, 'start')) || (DELETE_FORWARD.has(type) && caretAtEdge(el, 'end'))) {
      e.preventDefault();
      this.#joinAtEdge(el, backward);
      return;
    }
    this.#noteEdit(el, editKind(type));
    this.pendingDirty = [el];
  }

  #onPaste(e) {
    if (!this.model) return;
    e.preventDefault();
    // Plain text only: rich HTML would bring markup we can't map back to WordprocessingML.
    const lines = e.clipboardData.getData('text/plain').replace(/\r\n?/g, '\n').split('\n');
    this.#edit((el) => {
      const els = [el];
      lines.forEach((line, i) => {
        if (i > 0) els.push(this.#splitAtCaret(els.at(-1)));
        if (line) document.execCommand('insertText', false, line);
      });
      return els;
    });
  }

  #onCut(e) {
    if (!this.model) return;
    const plan = this.#plan();
    if (!plan.cross && !('error' in plan)) return; // within a paragraph: the browser's cut is fine
    e.preventDefault();
    if ('error' in plan) return this.#notify(plan.error);
    e.clipboardData.setData('text/plain', selectedText(plan));
    this.#edit();
  }

  #onCompositionStart() {
    // IME input can't be cancelled, so clear a multi-paragraph selection before it starts.
    if (this.model && this.#plan().cross) this.#edit();
  }

  /** Backspace at the start / Delete at the end of a paragraph: join it with its neighbour. */
  #joinAtEdge(el, backward) {
    const sibling = backward ? el.previousElementSibling : el.nextElementSibling;
    const other = sibling && editableParaOf(sibling) === sibling ? sibling : null;
    if (!other) return; // a table, read-only paragraph or container edge: nothing to join
    const [first, second] = backward ? [other, el] : [el, other];
    const caret = caretIn([first, second]);
    // Joining is deleting the (invisible) selection between the two paragraphs.
    getSelection().setBaseAndExtent(first, first.childNodes.length, second, 0);
    if (!this.#edit(undefined, { caret })) placeCaret([first, second], caret);
  }

  /**
   * Delete a selection spanning paragraphs; returns the surviving paragraph. Usually that's the
   * first one, with the rest of the last appended. A selection from the very start of the first
   * paragraph removes it whole instead, so the last paragraph keeps its own formatting (as when
   * deleting whole paragraphs in Word).
   */
  #deleteAcross({ range, start, end, blocks }) {
    const { model } = this;
    const head = document.createRange();
    head.setStart(range.startContainer, range.startOffset);
    head.setEnd(start, start.childNodes.length);
    const tail = document.createRange();
    tail.setStart(end, 0);
    tail.setEnd(range.endContainer, range.endOffset);
    const wholeFirst = blankBetween(start, 'start', range.startContainer, range.startOffset);
    const survivor = wholeFirst ? end : start;

    tail.deleteContents();
    let caret = null;
    if (wholeFirst) {
      fillIfEmpty(end);
    } else {
      const placeholder = lastLeaf(start)?.nodeName === 'BR' ? lastLeaf(start) : null;
      head.deleteContents();
      if (placeholder?.isConnected) placeholder.remove();
      // Prefer the caret inside the last text run, so typing keeps that run's formatting.
      const join = lastLeaf(start);
      const marker = document.createTextNode('');
      const rest = lastLeaf(end) ? [...end.childNodes] : [];
      start.append(marker, ...rest);
      if (!rest.length) protectTrailingBreak(start);
      caret = join?.nodeType === Node.TEXT_NODE ? [join, join.length] : [marker, 0];
    }

    const markers = model.markersOf.get(survivor);
    const moved = [];
    for (const b of blocks) {
      if (b === survivor) continue;
      if (isPara(b)) moved.push(...model.markersOf.get(b).start, ...model.markersOf.get(b).end);
      this.#xmlOf(b).remove();
      b.remove();
      for (const el of isPara(b) ? [b] : b.querySelectorAll('.para')) model.dirty.delete(el);
    }
    if (wholeFirst) markers.start.unshift(...moved);
    else markers.end.push(...moved);
    model.dirty.add(survivor);
    model.modified = true;

    if (caret) getSelection().collapse(...caret);
    else placeCaret([survivor]);
    renumber(model, this.page);
    this.current = survivor;
    return survivor;
  }

  /** Split `el` at the caret (deleting a selection within it); returns the new second paragraph. */
  #splitAtCaret(el) {
    const range = getSelection().getRangeAt(0);
    if (!range.collapsed) range.deleteContents();
    const tail = document.createRange();
    tail.setStart(range.startContainer, range.startOffset);
    tail.setEnd(el, el.childNodes.length);
    const fresh = el.cloneNode(false);
    fresh.append(tail.extractContents());
    protectTrailingBreak(el);
    fillIfEmpty(fresh);
    el.after(fresh);

    // The new <w:p> inherits paragraph properties; a section break moves to the second half,
    // since the section still ends after it.
    const { model } = this;
    const p = model.paraOf.get(el);
    const np = wEl(p.ownerDocument, 'p');
    const pPr = wChild(p, 'pPr');
    if (pPr) {
      np.append(pPr.cloneNode(true));
      wChild(pPr, 'sectPr')?.remove();
    }
    p.after(np);
    model.paraOf.set(fresh, np);
    model.markersOf.set(fresh, { start: [], end: [] });
    model.dirty.add(el).add(fresh);
    model.modified = true;

    placeCaret([fresh]);
    renumber(model, this.page);
    this.current = fresh;
    return fresh;
  }
}

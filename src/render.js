// Rendering document.xml into editable HTML.
//
// Each <w:p> becomes a <p class="para">. Plain runs become <span data-r> elements whose original
// <w:rPr> is kept in model.rprs, so edited text can be written back with its formatting. Anything
// we don't understand inline (hyperlinks, fields, images, footnote refs, ...) becomes an "atom": a
// non-editable span referencing the original XML nodes, which are re-inserted verbatim on save.
// Paragraphs containing tracked changes are rendered read-only.
//
// The whole page is one contenteditable host, so selections can span paragraphs. Read-only
// paragraphs and placeholders are non-editable islands; tables are islands whose cells are
// nested hosts, so a selection can't reach into (and break up) a table's structure.
import { NS, isW, wAttr, wChild, wDescendants, wEl } from './docx/xml.js';
import { StyleSheet, applyParaCss, applyRunCss, readParaProps, readRunProps } from './docx/styles.js';
import { Numbering } from './docx/numbering.js';

// Zero-width range markers. Kept (at the paragraph's start or end) when a paragraph is rewritten.
const MARKERS = new Set([
  'bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd', 'permStart', 'permEnd',
  'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd',
]);
// Inline wrappers that make a paragraph read-only (tracked changes and legacy smart tags).
const LOCKING = ['ins', 'del', 'moveFrom', 'moveTo', 'customXml', 'smartTag'];
// Run children that can be dropped: they're layout caches or spelling state.
const DROPPABLE = new Set(['lastRenderedPageBreak', 'proofErr']);

export function createModel(pkg) {
  return {
    pkg,
    styles: new StyleSheet(pkg.stylesDoc, pkg.themeDoc),
    numbering: new Numbering(pkg.numberingDoc),
    atoms: [], // index -> XML nodes to re-insert
    rprs: [], // index -> cloned <w:rPr> (or null)
    paraOf: new WeakMap(), // <p> element -> <w:p>
    blockOf: new WeakMap(), // other block element (table, content control, placeholder) -> XML node
    markersOf: new WeakMap(), // <p> element -> {start: Node[], end: Node[]}
    dirty: new Set(), // <p> elements whose text must be written back
    modified: false, // document.xml changed in some way
  };
}

export function renderDocument(model) {
  const body = wChild(model.pkg.doc.documentElement, 'body');
  const page = document.createElement('article');
  page.className = 'page';
  page.contentEditable = 'true';
  page.spellcheck = true;
  applyPageSize(page, wChild(body, 'sectPr'));
  for (const child of body.children) renderBlock(model, child, page);
  renumber(model, page);
  return page;
}

function applyPageSize(page, sectPr) {
  const size = wChild(sectPr, 'pgSz');
  const mar = wChild(sectPr, 'pgMar');
  const pt = (el, name, fallback) => (Number(wAttr(el, name) ?? fallback) / 20) + 'pt';
  page.style.width = pt(size, 'w', 11906);
  page.style.padding = `${pt(mar, 'top', 1440)} ${pt(mar, 'right', 1440)} ${pt(mar, 'bottom', 1440)} ${pt(mar, 'left', 1440)}`;
}

function renderBlock(model, node, container) {
  if (node.namespaceURI !== NS.w) {
    container.append(placeholder(model, node));
    return;
  }
  switch (node.localName) {
    case 'p':
      container.append(renderParagraph(model, node));
      break;
    case 'tbl':
      container.append(renderTable(model, node));
      break;
    case 'sdt': {
      const box = document.createElement('div');
      box.className = 'sdt';
      model.blockOf.set(box, node);
      for (const c of wChild(node, 'sdtContent')?.children ?? []) renderBlock(model, c, box);
      container.append(box);
      break;
    }
    case 'customXml':
      for (const c of node.children) renderBlock(model, c, container);
      break;
    case 'sectPr':
    case 'tblPr':
    case 'tcPr':
    case 'proofErr':
      break;
    default:
      if (!MARKERS.has(node.localName)) container.append(placeholder(model, node));
  }
}

function placeholder(model, node) {
  const div = document.createElement('div');
  div.className = 'unsupported';
  div.contentEditable = 'false';
  model.blockOf.set(div, node);
  div.textContent = `[${node.localName}]`;
  div.title = 'Not shown in this editor; preserved on save.';
  return div;
}

function renderTable(model, tbl) {
  const table = document.createElement('table');
  table.className = 'tbl';
  table.contentEditable = 'false';
  model.blockOf.set(table, tbl);
  const restartCells = new Map(); // grid column -> <td> that started a vertical merge
  for (const tr of rowsOf(tbl)) {
    const row = table.insertRow();
    let col = Number(wAttr(wChild(wChild(tr, 'trPr'), 'gridBefore'), 'val') ?? 0);
    for (const tc of cellsOf(tr)) {
      const tcPr = wChild(tc, 'tcPr');
      const span = Number(wAttr(wChild(tcPr, 'gridSpan'), 'val') ?? 1);
      const vMerge = wChild(tcPr, 'vMerge');
      const merge = vMerge ? wAttr(vMerge, 'val') ?? 'continue' : null;
      if (merge === 'continue' && restartCells.has(col)) {
        // Continuation cells are left untouched in the XML; they usually hold an empty paragraph.
        restartCells.get(col).rowSpan += 1;
        col += span;
        continue;
      }
      const td = row.insertCell();
      td.contentEditable = 'true';
      if (span > 1) td.colSpan = span;
      if (merge === 'restart') restartCells.set(col, td);
      else restartCells.delete(col);
      const width = wChild(tcPr, 'tcW');
      if (wAttr(width, 'type') === 'dxa') td.style.width = Number(wAttr(width, 'w')) / 20 + 'pt';
      const fill = wAttr(wChild(tcPr, 'shd'), 'fill');
      if (fill && fill !== 'auto') td.style.backgroundColor = '#' + fill;
      for (const c of tc.children) renderBlock(model, c, td);
      col += span;
    }
  }
  return table;
}

// Rows and cells can be wrapped in content controls or custom XML.
function* unwrap(el, name) {
  for (const c of el.children) {
    if (isW(c, name)) yield c;
    else if (isW(c, 'sdt')) yield* unwrap(wChild(c, 'sdtContent') ?? c, name);
    else if (isW(c, 'customXml')) yield* unwrap(c, name);
  }
}
const rowsOf = (tbl) => unwrap(tbl, 'tr');
const cellsOf = (tr) => unwrap(tr, 'tc');

/** Effective paragraph properties: style < numbering level < direct formatting. */
function paragraphProps(model, p) {
  const pPr = wChild(p, 'pPr');
  const styleId = wAttr(wChild(pPr, 'pStyle'), 'val');
  const base = model.styles.paragraphStyle(styleId);
  const direct = readParaProps(pPr, {});
  const numPr = direct.numPr ?? base.para.numPr;
  const level = numPr?.numId && numPr.numId !== '0'
    ? model.numbering.level(numPr.numId, numPr.ilvl) : null;
  const para = readParaProps(pPr, readParaProps(level?.pPr, { ...base.para }));
  para.numPr = level ? numPr : null;
  return { para, run: base.run, styleId: styleId ?? model.styles.defaultParaStyle };
}

/** Render `p` into a new element, or into `el` (reset first) to keep its identity. */
export function renderParagraph(model, p, el = document.createElement('p')) {
  el.replaceChildren();
  for (const { name } of [...el.attributes]) el.removeAttribute(name);
  el.className = 'para';
  model.paraOf.set(el, p);

  const { para, run, styleId } = paragraphProps(model, p);
  applyParaCss(el.style, para);
  applyRunCss(el.style, run);
  setRenderedState(el, run);
  el.dataset.style = styleId ?? '';

  const lock = LOCKING.find((name) => wDescendants(p, name).length > 0);
  const markers = { start: [], end: [] };
  model.markersOf.set(el, markers);
  renderInline(model, p, el, run, markers);

  // A trailing <br> doesn't render a line in HTML, and an empty paragraph has no height; both get
  // a placeholder <br> (dropped again on save).
  const last = lastLeaf(el);
  if (!last || last.nodeName === 'BR') el.append(document.createElement('br'));

  if (lock) {
    el.classList.add('locked');
    el.contentEditable = 'false';
    el.title = 'Read-only: this paragraph contains tracked changes or smart tags.';
  }
  return el;
}

/** Remember the formatting we rendered so serialization can detect user changes. */
function setRenderedState(el, props) {
  el.dataset.b = props.bold ? '1' : '0';
  el.dataset.i = props.italic ? '1' : '0';
  el.dataset.u = props.underline ? '1' : '0';
}

function renderInline(model, parent, el, paraRun, markers) {
  const kids = [...parent.children].filter((c) => !isW(c, 'pPr') && !isW(c, 'rPr'));
  let seenContent = false;
  for (let i = 0; i < kids.length; i++) {
    const k = kids[i];
    if (k.namespaceURI === NS.w && MARKERS.has(k.localName)) {
      (seenContent ? markers.end : markers.start).push(k);
      continue;
    }
    if (k.namespaceURI === NS.w && DROPPABLE.has(k.localName)) continue;
    seenContent = true;

    if (isW(k, 'r') && fieldCharType(k) === 'begin') {
      // A complex field spans several runs (begin ... separate ... end); keep it as one atom.
      const group = [k];
      let depth = 1;
      while (depth > 0 && ++i < kids.length) {
        group.push(kids[i]);
        const t = fieldCharType(kids[i]);
        if (t === 'begin') depth++;
        else if (t === 'end') depth--;
      }
      el.append(renderAtom(model, group, 'field'));
    } else if (isW(k, 'r')) {
      renderRun(model, k, el, paraRun);
    } else if (['ins', 'del', 'moveFrom', 'moveTo'].some((n) => isW(k, n))) {
      const wrap = document.createElement('span');
      wrap.className = isW(k, 'ins') || isW(k, 'moveTo') ? 'rev-ins' : 'rev-del';
      renderInline(model, k, wrap, paraRun, markers);
      el.append(wrap);
    } else if (isW(k, 'smartTag') || isW(k, 'customXml')) {
      renderInline(model, k, el, paraRun, markers);
    } else {
      el.append(renderAtom(model, [k], k.localName));
    }
  }
}

function fieldCharType(r) {
  return wAttr(wChild(r, 'fldChar'), 'fldCharType');
}

function renderRun(model, r, el, paraRun) {
  const rPr = wChild(r, 'rPr');
  const charStyle = wAttr(wChild(rPr, 'rStyle'), 'val');
  const props = readRunProps(rPr, {
    ...paraRun,
    ...(charStyle ? model.styles.characterStyle(charStyle) : {}),
  }, model.styles.theme);

  const span = document.createElement('span');
  span.dataset.r = model.rprs.push(rPr ? rPr.cloneNode(true) : null) - 1;
  setRenderedState(span, props);
  applyRunCss(span.style, props);

  for (const c of r.children) {
    if (c.namespaceURI !== NS.w) {
      span.append(renderAtom(model, [wrapInRun(r, c)], c.localName));
      continue;
    }
    switch (c.localName) {
      case 'rPr':
        break;
      case 't':
      case 'delText':
        span.append(c.textContent);
        break;
      case 'tab':
        span.append('\t');
        break;
      case 'noBreakHyphen':
        span.append('‑');
        break;
      case 'softHyphen':
        span.append('­');
        break;
      case 'cr':
        span.append(document.createElement('br'));
        break;
      case 'br': {
        const type = wAttr(c, 'type');
        if (type === 'page' || type === 'column') span.append(renderAtom(model, [wrapInRun(r, c)], 'break'));
        else span.append(document.createElement('br'));
        break;
      }
      default:
        if (!DROPPABLE.has(c.localName)) span.append(renderAtom(model, [wrapInRun(r, c)], c.localName));
    }
  }
  if (span.childNodes.length) el.append(span);
}

/** A standalone <w:r> holding a copy of `child` with the run's formatting. */
function wrapInRun(r, child) {
  const run = wEl(r.ownerDocument, 'r');
  const rPr = wChild(r, 'rPr');
  if (rPr) run.append(rPr.cloneNode(true));
  run.append(child.cloneNode(true));
  return run;
}

function renderAtom(model, nodes, kind) {
  const el = document.createElement('span');
  el.className = 'atom atom-' + kind;
  el.contentEditable = 'false';
  el.dataset.atom = model.atoms.push(nodes) - 1;
  describeAtom(model, nodes, kind, el);
  return el;
}

function drawingFrame(el) {
  while (el && !(el.namespaceURI === NS.wp && (el.localName === 'inline' || el.localName === 'anchor'))) {
    el = el.parentElement;
  }
  return el;
}

const textOf = (nodes, name = 't') =>
  nodes.flatMap((n) => wDescendants(n, name)).map((t) => t.textContent).join('');

function describeAtom(model, nodes, kind, el) {
  const node = nodes[0];
  const blips = nodes.flatMap((n) => [...n.getElementsByTagNameNS(NS.a, 'blip')]);
  if (blips.length) {
    for (const blip of blips) {
      const url = model.pkg.imageUrl(blip.getAttributeNS(NS.r, 'embed'));
      const extent = drawingFrame(blip)?.getElementsByTagNameNS(NS.wp, 'extent')[0];
      const img = url ? document.createElement('img') : document.createElement('span');
      if (url) img.src = url;
      else img.textContent = '[image]';
      if (extent) {
        // EMU -> CSS px (914400 EMU per inch, 96 px per inch).
        img.style.width = Number(extent.getAttribute('cx')) / 9525 + 'px';
        img.style.height = Number(extent.getAttribute('cy')) / 9525 + 'px';
      }
      el.append(img);
    }
    el.classList.add('atom-image');
    return;
  }
  switch (kind) {
    case 'hyperlink': {
      const rel = model.pkg.rels.get(node.getAttributeNS(NS.r, 'id'));
      const anchor = wAttr(node, 'anchor');
      el.textContent = textOf(nodes);
      el.title = rel?.target ?? (anchor ? '#' + anchor : 'link');
      return;
    }
    case 'field': {
      // Show the cached result (the runs after "separate"), like Word does.
      const sep = nodes.findIndex((r) => fieldCharType(r) === 'separate');
      el.textContent = sep >= 0 ? textOf(nodes.slice(sep + 1)) : '';
      el.title = 'Field: ' + textOf(nodes, 'instrText').trim();
      if (!el.textContent) el.textContent = '{field}';
      return;
    }
    case 'fldSimple':
      el.textContent = textOf(nodes) || '{field}';
      el.title = 'Field: ' + (wAttr(node, 'instr') ?? '').trim();
      return;
    case 'break':
      el.textContent = 'page break';
      return;
    case 'footnoteReference':
    case 'endnoteReference':
      el.textContent = '[' + wAttr(wChild(node, kind), 'id') + ']';
      el.title = kind === 'footnoteReference' ? 'Footnote' : 'Endnote';
      return;
    case 'sym': {
      const char = wAttr(wChild(node, 'sym'), 'char');
      const code = parseInt(char ?? '', 16);
      // Codes in the F000 range index into a symbol font, which we can't map reliably.
      el.textContent = code > 0 && code < 0xf000 ? String.fromCharCode(code) : '□';
      return;
    }
    case 'sdt':
      el.textContent = textOf(nodes) || '[content control]';
      el.title = 'Content control';
      return;
    case 'oMath':
    case 'oMathPara':
      el.textContent = [...node.getElementsByTagNameNS(NS.m, 't')].map((t) => t.textContent).join('') || '[equation]';
      el.title = 'Equation';
      return;
    case 'fldChar':
      // Boundary of a field that spans several paragraphs (e.g. a table of contents).
      el.textContent = '¤';
      el.title = 'Field boundary';
      return;
    case 'commentReference':
      el.textContent = '💬';
      el.title = 'Comment';
      return;
    default:
      el.textContent = textOf(nodes) || `[${kind}]`;
      el.title = `Preserved element: ${kind}`;
  }
}

/** Last non-empty text node, <br> or atom inside `node`, in document order. */
export function lastLeaf(node) {
  for (let c = node.lastChild; c; c = c.previousSibling) {
    if (c.nodeType === Node.TEXT_NODE) {
      if (c.data) return c;
    } else if (c.nodeType === Node.ELEMENT_NODE) {
      if (c.dataset.atom != null || c.nodeName === 'BR') return c;
      const inner = lastLeaf(c);
      if (inner) return inner;
    }
  }
  return null;
}

/** Recompute list markers for every paragraph, in document order. */
export function renumber(model, root) {
  const counter = model.numbering.counter();
  for (const el of root.querySelectorAll('.para')) {
    const { para } = paragraphProps(model, model.paraOf.get(el));
    const marker = para.numPr ? counter.next(para.numPr.numId, para.numPr.ilvl) : null;
    if (marker != null) el.dataset.marker = marker;
    else delete el.dataset.marker;
  }
}

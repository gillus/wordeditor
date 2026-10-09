// Caret positions as character offsets within a paragraph, so they survive re-rendering.
// Text counts one per character; a <br> or an atom counts as one.
import { lastLeaf } from './render.js';

const isUnit = (node) => node.dataset.atom != null || node.nodeName === 'BR';

function measure(node) {
  let n = 0;
  for (const c of node.childNodes) {
    if (c.nodeType === Node.TEXT_NODE) n += c.data.length;
    else if (c.nodeType === Node.ELEMENT_NODE) n += isUnit(c) ? 1 : measure(c);
  }
  return n;
}

/** {index, offset} of the selection start within `els`, or null if it isn't in any of them. */
export function caretIn(els) {
  const sel = getSelection();
  if (!sel.rangeCount) return null;
  const { startContainer, startOffset } = sel.getRangeAt(0);
  const index = els.findIndex((el) => el.contains(startContainer));
  if (index < 0) return null;
  const range = document.createRange();
  range.setStart(els[index], 0);
  range.setEnd(startContainer, startOffset);
  return { index, offset: measure(range.cloneContents()) };
}

function locate(node, state) {
  for (let i = 0; i < node.childNodes.length; i++) {
    const c = node.childNodes[i];
    if (c.nodeType === Node.TEXT_NODE) {
      if (state.remaining <= c.data.length) return [c, state.remaining];
      state.remaining -= c.data.length;
    } else if (c.nodeType === Node.ELEMENT_NODE) {
      if (isUnit(c)) {
        if (state.remaining === 0) return [node, i];
        state.remaining -= 1;
      } else {
        const hit = locate(c, state);
        if (hit) return hit;
      }
    }
  }
  return null;
}

/** Focus the paragraph and put the caret at `caret` (from caretIn), or at the start of `els[0]`. */
export function placeCaret(els, caret) {
  const el = els[caret?.index ?? 0];
  // Paragraphs aren't editing hosts themselves: focus the page or table cell holding them.
  el.closest('[contenteditable="true"]')?.focus({ preventScroll: true });
  let hit = locate(el, { remaining: caret?.offset ?? 0 });
  if (!hit) {
    // Past the end: stop before the placeholder <br>, if there is one.
    const last = lastLeaf(el);
    hit = last?.nodeName === 'BR'
      ? [last.parentNode, [...last.parentNode.childNodes].indexOf(last)]
      : [el, el.childNodes.length];
  }
  getSelection().collapse(...hit);
  el.scrollIntoView({ block: 'nearest' });
  return el;
}

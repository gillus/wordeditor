// Arrow-key moves the browser gets wrong. Each table cell is its own editing host inside a
// non-editable table, so native caret movement skips cells, walks them in source order for Up and
// Down, or leaves the caret between blocks. The same goes for read-only paragraphs. These helpers
// find where the caret should go instead: Left/Right follow document order (cells row by row),
// Up/Down follow the layout (the cell above or below, at the same column).
import { placeCaret } from './caret.js';

const SLACK = 3; // px of overlap tolerated between cells (collapsed borders)

const navigable = (el) => el.classList.contains('para') && !el.classList.contains('locked') && el.isContentEditable;
const sibling = (el, dir) => (dir < 0 ? el.previousElementSibling : el.nextElementSibling);

function lineRects(el) {
  const range = document.createRange();
  range.selectNodeContents(el);
  return [...range.getClientRects()].filter((r) => r.height > 0);
}

/**
 * The collapsed range's caret box(es). At a soft line wrap there are two, the end of one line and
 * the start of the next: which one the caret shows at depends on how it got there, which the DOM
 * doesn't tell. Some positions (next to a <br>, in an empty run) have none, so a neighbour is measured.
 */
export function caretRects(range) {
  const rects = [...range.getClientRects()].filter((r) => r.height > 0);
  if (rects.length) return rects.length > 1 ? [rects[0], rects.at(-1)] : rects;
  const { startContainer: c, startOffset: o } = range;
  const probe = document.createRange();
  if (c.nodeType === Node.TEXT_NODE && c.length) {
    const before = o === c.length;
    probe.setStart(c, before ? o - 1 : o);
    probe.setEnd(c, before ? o : o + 1);
    const r = probe.getBoundingClientRect();
    const x = before ? r.right : r.left;
    return [{ left: x, right: x, top: r.top, bottom: r.bottom }];
  }
  const next = c.nodeType === Node.ELEMENT_NODE ? c.childNodes[o] ?? c.childNodes[o - 1] : null;
  if (next) probe.selectNode(next);
  else probe.selectNodeContents(c.nodeType === Node.ELEMENT_NODE ? c : c.parentNode);
  return [probe.getBoundingClientRect()];
}

/**
 * For a caret at a soft wrap: do the native line move if it stays inside `para` (the caret was on
 * an inner line), and return whether it did. Otherwise leave the caret where it was.
 */
export function lineMoveStaysIn(para, dir) {
  const sel = getSelection();
  const { startContainer, startOffset } = sel.getRangeAt(0);
  sel.modify('move', dir < 0 ? 'backward' : 'forward', 'line');
  const now = sel.getRangeAt(0);
  if (para.contains(now.startContainer) && (now.startContainer !== startContainer || now.startOffset !== startOffset)) return true;
  sel.collapse(startContainer, startOffset);
  return false;
}

/** True if `rect` is on the first (dir -1) or last (dir +1) line of `para`. */
export function onEdgeLine(para, rect, dir) {
  return !lineRects(para).some((r) => (dir < 0 ? r.bottom <= rect.top + 0.5 : r.top >= rect.bottom - 0.5));
}

/**
 * The cell of `table` in the next row up (dir -1) / down (+1) from `rect`, or in the bottom/top
 * row when entering the table from outside (no rect), whichever is nearest to column `x`.
 * Works from the layout, so row and column spans need no bookkeeping.
 */
function cellToward(table, dir, x, rect) {
  let cells = [...table.rows].flatMap((row) => [...row.cells])
    .map((cell) => [cell, cell.getBoundingClientRect()])
    .filter(([, r]) => r.height > 0);
  if (rect) cells = cells.filter(([, r]) => (dir < 0 ? r.bottom <= rect.top + SLACK : r.top >= rect.bottom - SLACK));
  if (!cells.length) return null;
  const edge = dir < 0 ? Math.max(...cells.map(([, r]) => r.bottom)) : Math.min(...cells.map(([, r]) => r.top));
  const row = cells.filter(([, r]) => (dir < 0 ? r.bottom >= edge - SLACK : r.top <= edge + SLACK));
  const distance = (r) => Math.max(r.left - x, x - r.right, 0);
  return row.reduce((best, c) => (distance(c[1]) < distance(best[1]) ? c : best))[0];
}

/** The paragraph a caret entering `block` going up (dir -1) / down (+1) at column x lands in. */
function enter(block, dir, x) {
  if (block.classList.contains('para')) return navigable(block) ? block : null;
  if (block.matches('table.tbl')) {
    for (let cell = cellToward(block, dir, x, null); cell; cell = cellToward(block, dir, x, cell.getBoundingClientRect())) {
      const hit = enterChildren(cell, dir, x);
      if (hit) return hit;
    }
    return null;
  }
  if (block.classList.contains('sdt')) return enterChildren(block, dir, x);
  return null; // unsupported content: skipped
}

function enterChildren(box, dir, x) {
  const blocks = dir < 0 ? [...box.children].reverse() : [...box.children];
  for (const b of blocks) {
    const hit = enter(b, dir, x);
    if (hit) return hit;
  }
  return null;
}

/** Where Up (dir -1) / Down (+1) from the first/last line of `para` goes, at column x; null at the document's edge. */
export function verticalTarget(para, dir, x) {
  let from = para;
  while (!from.classList.contains('page')) {
    for (let b = sibling(from, dir); b; b = sibling(b, dir)) {
      const hit = enter(b, dir, x);
      if (hit) return hit;
    }
    const box = from.parentElement;
    if (box.matches('td')) {
      // Out of the cell: the cell above/below, else out of the table.
      const table = box.closest('table');
      for (let cell = cellToward(table, dir, x, box.getBoundingClientRect()); cell; cell = cellToward(table, dir, x, cell.getBoundingClientRect())) {
        const hit = enterChildren(cell, dir, x);
        if (hit) return hit;
      }
      from = table;
    } else {
      from = box;
    }
  }
  return null;
}

/** Where Left (dir -1) / Right (+1) from the start/end of `para` goes; null at the document's edge. */
export function horizontalTarget(page, para, dir) {
  const walker = document.createTreeWalker(page, NodeFilter.SHOW_ELEMENT, {
    acceptNode: (n) => (n.classList.contains('para') ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP),
  });
  walker.currentNode = para;
  for (let n; (n = dir < 0 ? walker.previousNode() : walker.nextNode());) if (navigable(n)) return n;
  return null;
}

/** A collapsed Range at a viewport point: the standard API where there is one, else WebKit/Blink's. */
function rangeFromPoint(x, y) {
  if (document.caretPositionFromPoint) {
    const pos = document.caretPositionFromPoint(x, y);
    if (!pos) return null;
    const range = document.createRange();
    range.setStart(pos.offsetNode, pos.offset);
    return range;
  }
  return document.caretRangeFromPoint?.(x, y) ?? null;
}

/** Put the caret in `para` at column x, on its last line when moving up (dir -1), first line when moving down. */
export function placeAtColumn(para, x, dir) {
  para.scrollIntoView({ block: 'nearest' });
  const rects = lineRects(para);
  if (rects.length) {
    const ref = rects.reduce((a, b) => (dir < 0 ? (b.bottom > a.bottom ? b : a) : (b.top < a.top ? b : a)));
    const y = (ref.top + ref.bottom) / 2;
    const line = rects.filter((r) => r.top <= y && r.bottom >= y);
    const left = Math.min(...line.map((r) => r.left));
    const right = Math.max(...line.map((r) => r.right));
    const px = Math.max(left + 0.5, Math.min(x, right - 0.5));
    const hit = rangeFromPoint(px, y);
    if (hit && para.contains(hit.startContainer)) {
      // Atoms are whole units: stop before or after one, whichever side is nearer.
      const atom = (hit.startContainer.nodeType === Node.ELEMENT_NODE ? hit.startContainer : hit.startContainer.parentElement).closest('[data-atom]');
      if (atom && para.contains(atom)) {
        const r = atom.getBoundingClientRect();
        if (px < (r.left + r.right) / 2) hit.setStartBefore(atom);
        else hit.setStartAfter(atom);
      }
      para.closest('[contenteditable="true"]')?.focus({ preventScroll: true });
      getSelection().collapse(hit.startContainer, hit.startOffset);
      return;
    }
  }
  // Not measurable (or off screen): the start or end of the paragraph.
  placeCaret([para], dir < 0 ? { index: 0, offset: Infinity } : undefined);
}

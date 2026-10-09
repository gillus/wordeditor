// Writing an edited paragraph's HTML back into its <w:p>.
//
// The paragraph's <w:pPr> is kept as-is; everything else is regenerated from the DOM: text becomes
// runs carrying the original <w:rPr> of the span it sits in (adjusted if the user toggled bold,
// italic or underline), atoms are re-inserted verbatim, and range markers go back at the edges.
import { NS, wChild, wEl, putInOrder, RPR_ORDER } from './docx/xml.js';
import { lastLeaf } from './render.js';

const SPECIAL = { '\t': 'tab', '\n': 'br', '‑': 'noBreakHyphen', '­': 'softHyphen' };

export function serializeParagraph(model, el) {
  const p = model.paraOf.get(el);
  const pPr = wChild(p, 'pPr');
  for (const c of [...p.childNodes]) if (c !== pPr) c.remove();

  // The final <br> is the browser's placeholder for an otherwise empty last line.
  const trailing = lastLeaf(el);
  const ctx = { model, root: el, skip: trailing?.nodeName === 'BR' ? trailing : null, segments: [], usedAtoms: new Set() };
  collect(ctx, el);

  const markers = model.markersOf.get(el);
  p.append(...markers.start, ...buildRuns(p.ownerDocument, ctx.segments), ...markers.end);
}

function collect(ctx, node) {
  for (const child of node.childNodes) {
    if (child.nodeType === Node.TEXT_NODE) {
      if (child.data) ctx.segments.push({ text: child.data, rPr: effectiveRPr(ctx, child.parentElement) });
    } else if (child.nodeType === Node.ELEMENT_NODE) {
      if (child.dataset.atom != null) {
        const nodes = ctx.model.atoms[Number(child.dataset.atom)];
        // Re-inserting the same node twice would move it; duplicates get copies.
        const id = child.dataset.atom;
        ctx.segments.push({ atom: ctx.usedAtoms.has(id) ? nodes.map((n) => n.cloneNode(true)) : nodes });
        ctx.usedAtoms.add(id);
      } else if (child.nodeName === 'BR') {
        if (child !== ctx.skip) ctx.segments.push({ text: '\n', rPr: effectiveRPr(ctx, child.parentElement) });
      } else {
        collect(ctx, child);
      }
    }
  }
}

/**
 * The <w:rPr> for text rendered inside `el`: the rPr of the run span it came from, with b/i/u
 * overridden where the computed style no longer matches what we originally rendered.
 */
function effectiveRPr({ model, root }, el) {
  const span = el.closest('[data-r]');
  const base = span && root.contains(span) ? span : root;
  const baseRPr = base === root ? null : model.rprs[Number(base.dataset.r)];

  const cs = getComputedStyle(el);
  let underline = false;
  for (let n = el; n && root.contains(n); n = n.parentElement) {
    if (getComputedStyle(n).textDecorationLine.includes('underline')) underline = true;
  }
  const now = { b: Number(cs.fontWeight) >= 600, i: cs.fontStyle !== 'normal', u: underline };

  const doc = model.pkg.doc;
  let rPr = baseRPr;
  for (const key of ['b', 'i', 'u']) {
    if (now[key] === (base.dataset[key] === '1')) continue;
    if (rPr === baseRPr) rPr = baseRPr ? baseRPr.cloneNode(true) : wEl(doc, 'rPr');
    const prop = key === 'u'
      ? wEl(doc, 'u', { val: now.u ? 'single' : 'none' })
      : wEl(doc, key, now[key] ? {} : { val: '0' });
    putInOrder(rPr, prop, RPR_ORDER);
  }
  return rPr;
}

function buildRuns(doc, segments) {
  const out = [];
  const signatures = new Map();
  const signature = (rPr) => {
    if (!rPr) return '';
    if (!signatures.has(rPr)) signatures.set(rPr, new XMLSerializer().serializeToString(rPr));
    return signatures.get(rPr);
  };
  let run = null;
  let runSig = null;
  for (const seg of segments) {
    if (seg.atom) {
      out.push(...seg.atom);
      run = null;
      continue;
    }
    const sig = signature(seg.rPr);
    if (!run || sig !== runSig) {
      run = wEl(doc, 'r');
      if (seg.rPr) run.append(seg.rPr.cloneNode(true));
      out.push(run);
      runSig = sig;
    }
    appendText(doc, run, seg.text);
  }
  return out;
}

function appendText(doc, run, text) {
  for (const piece of text.split(/([\t\n‑­])/)) {
    if (!piece) continue;
    if (SPECIAL[piece]) {
      run.append(wEl(doc, SPECIAL[piece]));
    } else if (run.lastChild?.localName === 't') {
      run.lastChild.textContent += piece;
    } else {
      const t = wEl(doc, 't');
      t.setAttributeNS(NS.xml, 'xml:space', 'preserve');
      t.textContent = piece;
      run.append(t);
    }
  }
}

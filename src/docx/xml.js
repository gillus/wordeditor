// Namespace constants and small helpers for working with WordprocessingML DOM nodes.

export const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  xml: 'http://www.w3.org/XML/1998/namespace',
};

export const isW = (node, name) =>
  node?.nodeType === 1 && node.namespaceURI === NS.w && node.localName === name;

export function wChild(el, name) {
  if (!el) return null;
  for (const c of el.children) if (isW(c, name)) return c;
  return null;
}

export function wChildren(el, name) {
  return el ? [...el.children].filter((c) => isW(c, name)) : [];
}

export function wDescendants(el, name) {
  return el ? [...el.getElementsByTagNameNS(NS.w, name)] : [];
}

export function wAttr(el, name) {
  if (!el) return null;
  return el.hasAttributeNS(NS.w, name) ? el.getAttributeNS(NS.w, name) : null;
}

export function wSetAttr(el, name, value) {
  el.setAttributeNS(NS.w, 'w:' + name, value);
}

export function wEl(doc, name, attrs = {}) {
  const el = doc.createElementNS(NS.w, 'w:' + name);
  for (const [k, v] of Object.entries(attrs)) wSetAttr(el, k, v);
  return el;
}

/** Interpret an ST_OnOff toggle element (<w:b/>, <w:b w:val="0"/>, ...). */
export function onOff(el) {
  if (!el) return undefined;
  const v = wAttr(el, 'val');
  return !(v === '0' || v === 'false' || v === 'off');
}

export function parseXml(text) {
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  const err = doc.getElementsByTagName('parsererror')[0];
  if (err) throw new Error('Invalid XML: ' + err.textContent.slice(0, 200));
  return doc;
}

export function serializeXml(doc) {
  const xml = new XMLSerializer().serializeToString(doc);
  return xml.startsWith('<?xml')
    ? xml
    : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml;
}

/** Insert `child` into `parent`, replacing any same-named element and respecting schema order. */
export function putInOrder(parent, child, order) {
  const existing = wChild(parent, child.localName);
  if (existing) {
    existing.replaceWith(child);
    return;
  }
  const rank = order.indexOf(child.localName);
  for (const c of parent.children) {
    const r = order.indexOf(c.localName);
    if (r > rank) {
      parent.insertBefore(child, c);
      return;
    }
  }
  parent.appendChild(child);
}

// CT_RPr and CT_PPr child sequences; Word rejects files whose children are out of order.
export const RPR_ORDER = [
  'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike',
  'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden',
  'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr',
  'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish',
  'oMath', 'rPrChange',
];

export const PPR_ORDER = [
  'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
  'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
  'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
  'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc',
  'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
  'sectPr', 'pPrChange',
];

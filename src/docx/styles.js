// Resolving paragraph/run formatting from styles.xml and converting it to CSS.
import { NS, wAttr, wChild, wChildren, onOff } from './xml.js';

const HIGHLIGHT = {
  yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff',
  red: '#ff0000', darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000',
  darkMagenta: '#800080', darkRed: '#800000', darkYellow: '#808000', darkGray: '#808080',
  lightGray: '#c0c0c0', black: '#000000', white: '#ffffff',
};

const twipsToPt = (v) => Number(v) / 20;

export function readRunProps(rPr, out, theme) {
  if (!rPr) return out;
  for (const c of rPr.children) {
    if (c.namespaceURI !== NS.w) continue;
    const val = wAttr(c, 'val');
    switch (c.localName) {
      case 'b': out.bold = onOff(c); break;
      case 'i': out.italic = onOff(c); break;
      case 'strike': case 'dstrike': out.strike = onOff(c); break;
      case 'caps': out.caps = onOff(c); break;
      case 'smallCaps': out.smallCaps = onOff(c); break;
      case 'vanish': out.hidden = onOff(c); break;
      case 'u': out.underline = val !== 'none'; break;
      case 'sz': out.size = Number(val) / 2; break;
      case 'color': if (val && val !== 'auto') out.color = '#' + val; break;
      case 'highlight': out.highlight = HIGHLIGHT[val]; break;
      case 'shd': {
        const fill = wAttr(c, 'fill');
        if (fill && fill !== 'auto') out.shading = '#' + fill;
        break;
      }
      case 'vertAlign': out.vertAlign = val; break;
      case 'rFonts': {
        const themeFont = wAttr(c, 'asciiTheme') ?? wAttr(c, 'hAnsiTheme');
        const font = themeFont ? theme[themeFont.startsWith('major') ? 'major' : 'minor'] : null;
        out.font = font ?? wAttr(c, 'ascii') ?? wAttr(c, 'hAnsi') ?? out.font;
        break;
      }
    }
  }
  return out;
}

export function readParaProps(pPr, out) {
  if (!pPr) return out;
  for (const c of pPr.children) {
    if (c.namespaceURI !== NS.w) continue;
    switch (c.localName) {
      case 'jc': {
        const v = wAttr(c, 'val');
        out.align = { left: 'left', start: 'left', center: 'center', right: 'right', end: 'right',
          both: 'justify', distribute: 'justify' }[v] ?? out.align;
        break;
      }
      case 'ind': {
        const left = wAttr(c, 'left') ?? wAttr(c, 'start');
        const right = wAttr(c, 'right') ?? wAttr(c, 'end');
        const first = wAttr(c, 'firstLine');
        const hanging = wAttr(c, 'hanging');
        if (left != null) out.indentLeft = twipsToPt(left);
        if (right != null) out.indentRight = twipsToPt(right);
        if (first != null) { out.firstLine = twipsToPt(first); out.hanging = 0; }
        if (hanging != null) { out.hanging = twipsToPt(hanging); out.firstLine = 0; }
        break;
      }
      case 'spacing': {
        const before = wAttr(c, 'before');
        const after = wAttr(c, 'after');
        const line = wAttr(c, 'line');
        if (before != null) out.spaceBefore = twipsToPt(before);
        if (after != null) out.spaceAfter = twipsToPt(after);
        if (line != null) {
          const rule = wAttr(c, 'lineRule') ?? 'auto';
          // Word's "single" spacing is roughly 1.15x the font size in CSS terms.
          out.lineHeight = rule === 'auto' ? String((Number(line) / 240) * 1.15) : twipsToPt(line) + 'pt';
        }
        break;
      }
      case 'numPr': {
        const numId = wAttr(wChild(c, 'numId'), 'val');
        const ilvl = wAttr(wChild(c, 'ilvl'), 'val');
        out.numPr = { numId: numId ?? out.numPr?.numId, ilvl: Number(ilvl ?? out.numPr?.ilvl ?? 0) };
        break;
      }
    }
  }
  return out;
}

export class StyleSheet {
  constructor(stylesDoc, themeDoc) {
    this.theme = readThemeFonts(themeDoc);
    this.styles = new Map();
    this.defaultParaStyle = null;
    this.defaultRun = {};
    this.defaultPara = {};
    this.cache = new Map();
    if (!stylesDoc) return;

    const root = stylesDoc.documentElement;
    const docDefaults = wChild(root, 'docDefaults');
    readRunProps(wChild(wChild(docDefaults, 'rPrDefault'), 'rPr'), this.defaultRun, this.theme);
    readParaProps(wChild(wChild(docDefaults, 'pPrDefault'), 'pPr'), this.defaultPara);

    for (const s of wChildren(root, 'style')) {
      const id = wAttr(s, 'styleId');
      const style = {
        id,
        type: wAttr(s, 'type'),
        name: wAttr(wChild(s, 'name'), 'val') ?? id,
        basedOn: wAttr(wChild(s, 'basedOn'), 'val'),
        pPr: wChild(s, 'pPr'),
        rPr: wChild(s, 'rPr'),
        hidden: !!wChild(s, 'semiHidden') || !!wChild(s, 'hidden'),
      };
      this.styles.set(id, style);
      const isDefault = wAttr(s, 'default');
      if (style.type === 'paragraph' && isDefault != null && !['0', 'false', 'off'].includes(isDefault)) {
        this.defaultParaStyle = id;
      }
    }
  }

  /** Style chain from the root ancestor down to `id`. */
  chain(id) {
    const out = [];
    const seen = new Set();
    for (let s = this.styles.get(id); s && !seen.has(s.id); s = this.styles.get(s.basedOn)) {
      seen.add(s.id);
      out.unshift(s);
    }
    return out;
  }

  /** Resolved {para, run} properties for a paragraph style (falls back to the default style). */
  paragraphStyle(id) {
    const key = 'p:' + (id ?? '');
    if (!this.cache.has(key)) {
      const para = { ...this.defaultPara };
      const run = { ...this.defaultRun };
      const styleId = id && this.styles.get(id)?.type === 'paragraph' ? id : this.defaultParaStyle;
      for (const s of this.chain(styleId)) {
        readParaProps(s.pPr, para);
        readRunProps(s.rPr, run, this.theme);
      }
      this.cache.set(key, { para, run });
    }
    return this.cache.get(key);
  }

  /** Run properties contributed by a character style. */
  characterStyle(id) {
    const key = 'c:' + id;
    if (!this.cache.has(key)) {
      const run = {};
      for (const s of this.chain(id)) readRunProps(s.rPr, run, this.theme);
      this.cache.set(key, run);
    }
    return this.cache.get(key);
  }

  /** Paragraph styles to offer in the UI: visible ones plus any already used. */
  paragraphStyles(usedIds) {
    return [...this.styles.values()]
      .filter((s) => s.type === 'paragraph' && (!s.hidden || usedIds.has(s.id)))
      .map((s) => ({ id: s.id, name: s.name }))
      .sort((a, b) => a.name.localeCompare(b.name));
  }
}

function readThemeFonts(themeDoc) {
  const fonts = { major: null, minor: null };
  if (!themeDoc) return fonts;
  for (const kind of ['major', 'minor']) {
    const el = themeDoc.getElementsByTagNameNS(NS.a, kind + 'Font')[0];
    const latin = el && [...el.children].find((c) => c.localName === 'latin');
    fonts[kind] = latin?.getAttribute('typeface') || null;
  }
  return fonts;
}

export function applyRunCss(style, p) {
  style.fontWeight = p.bold ? '700' : '400';
  style.fontStyle = p.italic ? 'italic' : 'normal';
  const deco = [p.underline && 'underline', p.strike && 'line-through'].filter(Boolean);
  style.textDecorationLine = deco.length ? deco.join(' ') : 'none';
  if (p.size) style.fontSize = p.size + 'pt';
  if (p.color) style.color = p.color;
  if (p.font) style.fontFamily = `"${p.font}", sans-serif`;
  if (p.highlight ?? p.shading) style.backgroundColor = p.highlight ?? p.shading;
  if (p.caps) style.textTransform = 'uppercase';
  if (p.smallCaps) style.fontVariantCaps = 'small-caps';
  if (p.hidden) style.opacity = '0.4';
  if (p.vertAlign === 'superscript' || p.vertAlign === 'subscript') {
    style.verticalAlign = p.vertAlign === 'superscript' ? 'super' : 'sub';
    style.fontSize = (p.size ?? 11) * 0.65 + 'pt';
  }
}

export function applyParaCss(style, p) {
  if (p.align) style.textAlign = p.align;
  style.marginTop = (p.spaceBefore ?? 0) + 'pt';
  style.marginBottom = (p.spaceAfter ?? 0) + 'pt';
  style.paddingLeft = (p.indentLeft ?? 0) + 'pt';
  style.paddingRight = (p.indentRight ?? 0) + 'pt';
  style.textIndent = (p.firstLine || -(p.hanging ?? 0)) + 'pt';
  style.setProperty('--hang', (p.hanging ?? 0) + 'pt');
  if (p.lineHeight) style.lineHeight = p.lineHeight;
}

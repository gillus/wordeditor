// List numbering: level definitions from numbering.xml and running counters for markers.
import { wAttr, wChild, wChildren } from './xml.js';

// Symbol/Wingdings private-use bullet glyphs and their closest Unicode equivalents.
const BULLETS = { '': '•', '': '▪', '': '❖', '': '➢', '': '✓', o: '◦' };

function readLevel(lvl) {
  return {
    start: Number(wAttr(wChild(lvl, 'start'), 'val') ?? 1),
    format: wAttr(wChild(lvl, 'numFmt'), 'val') ?? 'decimal',
    text: wAttr(wChild(lvl, 'lvlText'), 'val') ?? '',
    pPr: wChild(lvl, 'pPr'),
  };
}

function levelsOf(el) {
  const levels = new Map();
  for (const lvl of wChildren(el, 'lvl')) levels.set(Number(wAttr(lvl, 'ilvl')), readLevel(lvl));
  return levels;
}

function toRoman(n) {
  const table = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'],
    [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
  let out = '';
  for (const [v, s] of table) while (n >= v) { out += s; n -= v; }
  return out;
}

function toLetters(n) {
  // Word repeats the letter: a..z, aa..zz, aaa...
  const letter = String.fromCharCode(97 + ((n - 1) % 26));
  return letter.repeat(Math.floor((n - 1) / 26) + 1);
}

function formatNumber(n, format) {
  switch (format) {
    case 'lowerLetter': return toLetters(n);
    case 'upperLetter': return toLetters(n).toUpperCase();
    case 'lowerRoman': return toRoman(n);
    case 'upperRoman': return toRoman(n).toUpperCase();
    case 'decimalZero': return String(n).padStart(2, '0');
    case 'none': return '';
    default: return String(n);
  }
}

export class Numbering {
  constructor(numberingDoc) {
    this.abstracts = new Map();
    this.nums = new Map();
    if (!numberingDoc) return;
    const root = numberingDoc.documentElement;
    for (const a of wChildren(root, 'abstractNum')) {
      this.abstracts.set(wAttr(a, 'abstractNumId'), levelsOf(a));
    }
    for (const n of wChildren(root, 'num')) {
      const overrides = new Map();
      for (const o of wChildren(n, 'lvlOverride')) {
        const lvl = wChild(o, 'lvl');
        overrides.set(Number(wAttr(o, 'ilvl')), {
          start: wAttr(wChild(o, 'startOverride'), 'val'),
          level: lvl ? readLevel(lvl) : null,
        });
      }
      this.nums.set(wAttr(n, 'numId'), {
        abstractId: wAttr(wChild(n, 'abstractNumId'), 'val'),
        overrides,
      });
    }
  }

  level(numId, ilvl) {
    const num = this.nums.get(numId);
    if (!num) return null;
    return num.overrides.get(ilvl)?.level ?? this.abstracts.get(num.abstractId)?.get(ilvl) ?? null;
  }

  /**
   * Counter state for one pass over the document. Lists sharing an abstract definition continue
   * each other's numbering unless a num restarts it with a startOverride.
   */
  counter() {
    const counts = new Map(); // abstractId -> number[] per level
    const restarted = new Set();
    return {
      next: (numId, ilvl) => {
        const num = this.nums.get(numId);
        const def = this.level(numId, ilvl);
        if (!num || !def) return null;
        let levels = counts.get(num.abstractId);
        if (!levels) counts.set(num.abstractId, (levels = []));
        if (!restarted.has(numId)) {
          restarted.add(numId);
          for (const [lvl, o] of num.overrides) if (o.start != null) levels[lvl] = Number(o.start) - 1;
        }
        levels[ilvl] = (levels[ilvl] ?? this.level(numId, ilvl).start - 1) + 1;
        levels.length = ilvl + 1; // deeper levels restart
        if (def.format === 'bullet') return BULLETS[def.text] ?? (def.text || '•');
        return def.text.replace(/%(\d)/g, (_, d) => {
          const lvl = Number(d) - 1;
          const value = levels[lvl] ?? this.level(numId, lvl)?.start ?? 1;
          return formatNumber(value, this.level(numId, lvl)?.format ?? 'decimal');
        });
      },
    };
  }
}

import { describe, expect, it } from 'vitest';
import {
  isSymbolFontFamily,
  symbolFontToUnicode,
  symbolTextToUnicodeSegments,
} from './symbol-font';
import { createSymbolFaceOracle, createSymbolMonochromeProbe } from './symbol-face';
import { SENTINEL_FAMILIES } from './route-sentinels';
import { ADOBE_SYMBOL_UNICODE } from './symbol-encoding';

const ch = (code: number): string => String.fromCodePoint(code);
const VS15 = '\uFE0E';
const missing = { native: () => false };
const mono = { native: () => false, monochrome: () => true };

describe('symbol text conversion for a face that lacks the codes', () => {
  const both = (family: string, code: number, expected: string) => {
    expect(symbolFontToUnicode(ch(code), family, missing)).toBe(expected);
    expect(symbolFontToUnicode(ch(0xf000 + code), family, missing)).toBe(expected);
  };

  it('maps the whole Symbol repertoire, bare and F0-shifted (#2177)', () => {
    both('Symbol', 0x61, ch(0x3b1)); both('Symbol', 0x62, ch(0x3b2)); both('Symbol', 0x63, ch(0x3c7));
    both('Symbol', 0x70, ch(0x3c0)); both('Symbol', 0x53, ch(0x3a3)); both('Symbol', 0xe5, ch(0x2211));
    both('Symbol', 0xf2, ch(0x222b)); both('Symbol', 0xb6, ch(0x2202)); both('Symbol', 0xb7, ch(0x2022));
    both('Symbol', 0xa0, ch(0x20ac)); both('Symbol', 0xe6, ch(0x239b)); both('Symbol', 0x60, ch(0x203e));
    // Phi: the documented exchange (looped 0x6A, stroked 0x66).
    both('Symbol', 0x6a, ch(0x3c6)); both('Symbol', 0x66, ch(0x3d5));
    expect(symbolFontToUnicode('abcp', 'Symbol', missing)).toBe('\u03b1\u03b2\u03c7\u03c0');
    expect(symbolFontToUnicode(ch(0xf0f0), 'Symbol', missing)).toBe(ch(0xf0f0));
  });

  it('decodes the stored Adobe table exactly (160 entries, pinned hash)', () => {
    let h = 0;
    for (let code = 0x20; code <= 0xfe; code++) {
      const unicode = ADOBE_SYMBOL_UNICODE[code];
      if (unicode !== undefined) h = (Math.imul(h, 31) + unicode * (code - 0x1f)) >>> 0;
    }
    expect([Object.keys(ADOBE_SYMBOL_UNICODE).length, h]).toEqual([160, 110040407]);
  });

  it('agrees with Adobe symbol.txt except the recorded phi exchange', () => {
    for (const [key, unicode] of Object.entries(ADOBE_SYMBOL_UNICODE)) {
      const code = Number(key);
      const expected = ch(unicode);
      if (expected === ch(code) || /(?![#*0-9])\p{Emoji}/u.test(expected)) continue;
      expect(symbolFontToUnicode(ch(0xf000 + code), 'Symbol', missing)).toBe(expected);
    }
  });

  it('keeps each family on its own table', () => {
    // 0xA7 is a club in Symbol and a small square in Wingdings (both pictographs).
    expect(symbolFontToUnicode(ch(0xf0a7), 'Symbol', mono)).toBe(ch(0x2663) + VS15);
    expect(symbolFontToUnicode(ch(0xf0a7), 'Wingdings', mono)).toBe(ch(0x25aa) + VS15);
    both('Wingdings 2', 0x80, ch(0x2609)); // SUN
    both('Wingdings 3', 0x21, ch(0x2b60)); // LEFTWARDS TRIANGLE-HEADED ARROW
    expect(symbolFontToUnicode('a', 'Calibri', missing)).toBe('a');
    expect(isSymbolFontFamily(' WINGDINGS 2 ')).toBe(true);
    expect(isSymbolFontFamily('SymbolMT')).toBe(false);
  });

  // Pins the decoded N4384 tables (dingbat-encoding.ts stores deltas): mapped
  // count and an order-sensitive hash of every (code, code point) pair, taken
  // from the verified literal tables.
  it.each([
    ['Wingdings', 221, 2837286855], ['Wingdings 2', 216, 3816834078],
    ['Wingdings 3', 207, 3231793635], ['Webdings', 222, 1275301490],
  ])('decodes the whole %s table', (family, count, hash) => {
    let mapped = 0;
    let h = 0;
    for (let code = 0x20; code <= 0xff; code++) {
      const out = symbolFontToUnicode(ch(0xf000 + code), family, { ...missing, monochrome: () => true });
      if (out === ch(0xf000 + code)) continue;
      mapped++;
      h = (Math.imul(h, 31) + out.codePointAt(0)! * (code - 0x20 + 1)) >>> 0;
    }
    expect([mapped, h]).toEqual([count, hash]);
  });

  it('writes pictographs with VS15 and converts them only when proven monochrome', () => {
    // Symbol 0xA9 is BLACK HEART SUIT; Webdings 0x21 is SPIDER (U+1F577).
    expect(symbolFontToUnicode(ch(0xf0a9), 'Symbol', mono)).toBe(ch(0x2665) + VS15);
    expect(symbolFontToUnicode(ch(0xf021), 'Webdings', mono)).toBe(ch(0x1f577) + VS15);
    expect(symbolFontToUnicode(ch(0xf0a9), 'Symbol', missing)).toBe(ch(0xf0a9));
    expect(symbolFontToUnicode(ch(0xf021), 'Webdings', { ...mono, monochrome: () => false })).toBe(ch(0xf021));
  });

  it('splits at every mapped/unmapped boundary', () => {
    expect(symbolTextToUnicodeSegments(ch(0xf061) + ch(0xf0f0) + ch(0xf062), 'Symbol', missing)).toEqual([
      { text: ch(0x3b1), mapped: true }, { text: ch(0xf0f0), mapped: false }, { text: ch(0x3b2), mapped: true },
    ]);
  });
});


describe('symbol text conversion for a font that draws the codes', () => {
  it('keeps every code the painted font draws, per form', () => {
    const native = (code: string) => code === ch(0xf061) || code === 'b';
    expect(symbolTextToUnicodeSegments(ch(0xf061) + 'b' + ch(0xf063), 'Symbol', { native })).toEqual([
      { text: ch(0xf061) + 'b', mapped: false }, { text: ch(0x3c7), mapped: true },
    ]);
    // A font that draws only the PUA form still converts the bare form.
    expect(symbolFontToUnicode(ch(0xf061) + 'a', 'Symbol', { native: (code) => code === ch(0xf061) }))
      .toBe(ch(0xf061) + ch(0x3b1));
  });

  it('keeps a native bare code that bidi mirroring would replace as its F0-shifted form', () => {
    // Right-to-left shaping turns a bare "\u00AB" into "\u00BB": in Symbol,
    // the almost-equal sign instead of the left-right arrow. The shifted form
    // is the same glyph and is never mirrored.
    const native = () => true;
    expect(symbolTextToUnicodeSegments('a\u00AB\uF0BB', 'Symbol', { native }))
      .toEqual([{ text: 'a\uF0AB\uF0BB', mapped: false }]);
    expect(symbolFontToUnicode('(x)', 'Wingdings', { native })).toBe('\uF028x\uF029');
    // Symbol's own parentheses are parentheses: mirroring them is right.
    expect(symbolFontToUnicode('(', 'Symbol', { native })).toBe('(');
    // Where the caller would paint the shifted form in another font, the code converts.
    expect(symbolFontToUnicode('\u00AB', 'Symbol', { native, keepsShifted: () => false, monochrome: () => true }))
      .toBe('\u2194' + VS15);
    // Not native, or not judged: converted as any other code.
    expect(symbolFontToUnicode('\u00AB', 'Symbol', mono)).toBe('\u2194' + VS15);
    expect(symbolFontToUnicode('\u00AB', 'Symbol')).toBe('\u2194' + VS15);
  });

  it('asks only about codes that have a conversion', () => {
    const asked: string[] = [];
    // Digits are themselves in Symbol; 0xF0 has no target; "α" is no code.
    symbolFontToUnicode('1' + ch(0xf0f0) + ch(0x3b1) + 'a', 'Symbol', { native: (code) => { asked.push(code); return false; } });
    expect(asked).toEqual(['a']);
  });

  it('keeps the 0.89 subset when the font could not be judged', () => {
    expect(symbolFontToUnicode(ch(0xf0b7) + ch(0xf061), 'Symbol', { native: () => undefined }))
      .toBe(ch(0x2022) + ch(0xf061));
    expect(symbolFontToUnicode(ch(0xf0e0), 'Wingdings')).toBe(ch(0x2192));
    expect(symbolFontToUnicode(ch(0xf0a7), 'Webdings')).toBe(ch(0xf0a7));
  });
});

/** A Node-style canvas class whose text ink is `pixel(fillStyle, text)`. */
const inkCanvas = (pixel: (fill: string, text: string) => number[], onCreate = () => {}) => class {
  constructor(readonly width: number, readonly height: number) { onCreate(); }
  getContext() {
    let last: string | null = null;
    return {
      font: '', fillStyle: '', textBaseline: '', globalAlpha: 1, globalCompositeOperation: '',
      clearRect() { last = null; },
      fillText(text: string) { last = text; },
      getImageData() {
        return { data: new Uint8ClampedArray(last === null ? [0, 0, 0, 0] : pixel(this.fillStyle, last)) };
      },
    };
  }
};
const filled = (fill: string) => (fill === '#f00' ? [255, 0, 0, 255] : fill === '#00f' ? [0, 0, 255, 255] : [0, 0, 0, 255]);

describe('createSymbolMonochromeProbe', () => {
  it('draws on another canvas of the painting canvas class (Node canvases)', () => {
    let created = 0;
    const NodeCanvas = inkCanvas(filled, () => { created++; });
    const owner = { canvas: new NodeCanvas(1, 1) };
    expect(createSymbolMonochromeProbe(owner, '32px sans-serif')('♥︎')).toBe(true);
    expect(created).toBe(2);
  });

  it('allocates through the canvas factory of the owner, not the global helpers of the realm', () => {
    // The owner canvas's class and document paint in colour; the factory's
    // backend, the one that paints, follows the fill.
    const Other = inkCanvas(() => [200, 30, 30, 255]);
    const Painting = inkCanvas(filled);
    const sizes: number[][] = [];
    const factory = (width: number, height: number) => { sizes.push([width, height]); return new Painting(width, height); };
    const owner = { canvas: Object.assign(new Other(1, 1), { ownerDocument: { createElement: () => new Other(1, 1) } }) };
    expect(createSymbolMonochromeProbe(owner, '32px sans-serif', factory)('\u2665\uFE0E')).toBe(true);
    expect(createSymbolMonochromeProbe(owner, '32px sans-serif')('\u2665\uFE0E')).toBe(false);
    expect(sizes).toEqual([[48, 48]]);
    // Through the oracle too; a factory that yields nothing leaves it inconclusive.
    const measureText = (text: string) => ({ width: text === 'i' ? 3 : 7 });
    const judged = (createCanvas: Parameters<typeof createSymbolFaceOracle>[1]) =>
      createSymbolFaceOracle(Object.assign({ font: '', measureText }, owner), createCanvas)
        .segments('\uF0A9', 'Symbol', '400 32px "Absent"').map((part) => part.mapped);
    expect(judged({ createCanvas: factory })).toEqual([true]);
    expect(judged({})).toEqual([false]);
    expect(judged({ createCanvas: () => null })).toEqual([false]);
  });

  it.each([
    ['a fixed black palette', () => [0, 0, 0, 255]],
    ['its own colours', () => [200, 30, 30, 255]],
    ['one of the probe colours', () => [255, 0, 0, 255]],
  ])('rejects a glyph painted in %s whatever the fill', (_case, pixel) => {
    const Canvas = inkCanvas(pixel);
    expect(createSymbolMonochromeProbe({ canvas: new Canvas(1, 1) })('♥︎')).toBe(false);
  });

  it('judges solid pixels only, and nothing without ink', () => {
    // A faint antialiasing pixel carries too little colour to judge.
    const Faint = inkCanvas((fill) => [...filled(fill), ...[90, 90, 90, 20]]);
    expect(createSymbolMonochromeProbe({ canvas: new Faint(1, 1) })('x')).toBe(true);
    // A wholly translucent glyph has no pixel to judge: inconclusive, not proven.
    const Translucent = inkCanvas(() => [200, 30, 30, 50]);
    expect(createSymbolMonochromeProbe({ canvas: new Translucent(1, 1) })('\u{1F44D}\uFE0E')).toBeUndefined();
    const Empty = inkCanvas(() => [0, 0, 0, 0]);
    expect(createSymbolMonochromeProbe({ canvas: new Empty(1, 1) })(' ')).toBeUndefined();
  });

  it('leaves a spoofed readback inconclusive, never "not monochrome"', () => {
    // An anti-fingerprinting canvas: every readback carries noise, even of
    // the canvas just cleared, so the colour pixels prove nothing.
    const Spoofed = class {
      getContext() {
        return {
          font: '', fillStyle: '', textBaseline: '', globalAlpha: 1, globalCompositeOperation: '',
          clearRect() {}, fillText() {},
          getImageData() { return { data: new Uint8ClampedArray([3, 250, 7, 255]) }; },
        };
      }
    };
    expect(createSymbolMonochromeProbe({ canvas: new Spoofed() })('\u2665\uFE0E')).toBeUndefined();
  });
});

describe('oracle colour proof', () => {
  const measureText = (text: string) => ({ width: text === 'i' ? 3 : 7 });
  const withCanvas = (canvas: object) => ({ font: '', canvas, measureText });

  it('requires the base without VS15 to be monochrome too (a shaper may split them)', () => {
    // Colour unless the text carries VS15.
    const Canvas = inkCanvas((fill, text) => (text.includes('︎') ? filled(fill) : [200, 30, 30, 255]));
    const oracle = createSymbolFaceOracle(withCanvas(new Canvas(1, 1)));
    expect(oracle.segments('', 'Symbol', '400 32px "Absent"')).toEqual([{ text: '', mapped: false }]);
  });

  it('keeps the authored code of any target a colour glyph would paint, emoji-capable or not', () => {
    // Symbol 0x61 is alpha: no Emoji property, yet a colour face may hold it.
    const Colour = inkCanvas(() => [0, 0, 0, 255]);
    const coloured = createSymbolFaceOracle(withCanvas(new Colour(1, 1)));
    expect(coloured.segments('', 'Symbol', '400 32px "Absent"')).toEqual([{ text: '', mapped: false }]);
    const Plain = inkCanvas(filled);
    const plain = createSymbolFaceOracle(withCanvas(new Plain(1, 1)));
    expect(plain.segments('', 'Symbol', '400 32px "Absent"')).toEqual([{ text: 'α♥︎', mapped: true }]);
    // No canvas to read back: a text-default target converts, an emoji-capable one is kept.
    const blind = createSymbolFaceOracle({ font: '', measureText });
    expect(blind.segments('', 'Symbol', '400 32px "Absent"'))
      .toEqual([{ text: 'α', mapped: true }, { text: '', mapped: false }]);
  });
});

// A context that draws the "Sym" face's codes 0x41..0x5A at width 2+code
// (both forms), the "Pua" face's codes only as U+F000+code, and every other
// text in a fallback at one width per character class. In a right-to-left
// direction a bare code is mirrored (another width), as "(" is, so an
// identity holds only when measured left to right.
const letters = (code: number) => code >= 0x41 && code <= 0x5a;
const fallbackWidth = (text: string) => {
  const code = text.codePointAt(0)!;
  return code >= 0xf000 ? 1 : text === 'i' ? 3 : 5 + (code % 7);
};
const namedWidth = function (this: { font: string; direction?: CanvasDirection }, text: string): number | undefined {
  const code = text.codePointAt(0)!;
  const pua = code >= 0xf000;
  const base = pua ? code - 0xf000 : code;
  if (this.font.includes('"Sym"') && letters(base)) return 2 + base + (this.direction === 'rtl' && !pua ? 1 : 0);
  if (this.font.includes('"Pua"') && pua && letters(base)) return 2 + base;
  // A named Latin face: bare codes only.
  if (this.font.includes('"Latin"') && !pua) return 40 + (code % 5);
  return undefined;
};

describe('symbol face oracle without a FontFaceSet (plain mode: a Node canvas)', () => {
  const context = () => ({
    font: '10px sans-serif',
    direction: 'rtl' as CanvasDirection,
    measureText(text: string) { return { width: namedWidth.call(this, text) ?? fallbackWidth(text) }; },
  });

  it('keeps exactly the codes whose bare and PUA forms draw identically in the painted font', () => {
    const ctx = context();
    const oracle = createSymbolFaceOracle(ctx);
    const font = 'italic 700 20px "Sym", "Latin", serif';
    expect([oracle.draws('A', font), oracle.draws(ch(0xf041), font)]).toEqual([true, true]);
    expect([oracle.draws('a', font), oracle.draws(ch(0xf061), font)]).toEqual([false, false]);
    expect(oracle.draws('A', '20px "Other", serif')).toBe(false);
    // A font that maps only the PUA form is not proven here: it converts.
    expect([oracle.draws('A', '20px "Pua"'), oracle.draws(ch(0xf041), '20px "Pua"')]).toEqual([false, false]);
    expect(oracle.segments(ch(0xf041) + 'A' + ch(0xf061) + 'a', 'Symbol', font)).toEqual([
      { text: ch(0xf041) + 'A', mapped: false }, { text: ch(0x3b1) + ch(0x3b1), mapped: true },
    ]);
    expect([ctx.font, ctx.direction]).toEqual(['10px sans-serif', 'rtl']);
  });

  it('keeps the PUA form of a default-ignorable code in a font proven symbol-encoded', () => {
    const ctx = {
      font: '',
      measureText(text: string) {
        const code = text.codePointAt(0)!;
        if (code === 0xad) return { width: 0 };
        if (this.font.includes('"Sym"') && code === 0xf0ad) return { width: 77 };
        return { width: namedWidth.call(this, text) ?? fallbackWidth(text) };
      },
    };
    const oracle = createSymbolFaceOracle(ctx);
    expect([oracle.draws(ch(0xf0ad), '20px "Sym"'), oracle.draws(ch(0xad), '20px "Sym"')]).toEqual([true, false]);
    expect(oracle.draws(ch(0xf0ad), '20px "Other"')).toBe(false);
  });

  it('judges nothing without a context, with one that measures nothing, or for a font it cannot split', () => {
    expect(createSymbolFaceOracle(null).draws('A', '20px "Sym"')).toBeUndefined();
    expect(createSymbolFaceOracle({ font: '', measureText: () => ({ width: 0 }) }).draws('A', '20px "Sym"')).toBeUndefined();
    expect(createSymbolFaceOracle(context()).draws('A', '"Sym"')).toBeUndefined();
    // Not a symbol code at all.
    expect(createSymbolFaceOracle(context()).draws(ch(0x3b1), '20px "Sym"')).toBe(false);
  });

  it('keeps a verdict for the oracle lifetime only', () => {
    let calls = 0;
    const ctx = { font: '', measureText: (text: string) => { calls++; return { width: fallbackWidth(text) }; } };
    const oracle = createSymbolFaceOracle(ctx);
    oracle.draws('A', '400 32px "A"');
    const measured = calls;
    expect(measured).toBeGreaterThan(0);
    // A kept verdict is answered without touching the context.
    oracle.draws('A', '400 32px "A"');
    oracle.segments('A', 'Symbol', '400 32px "A"');
    expect(calls).toBe(measured);
    createSymbolFaceOracle(ctx).draws('A', '400 32px "A"');
    expect(calls).toBe(measured * 2);
  });
});

describe('symbol face oracle with route sentinels (a font realm)', () => {
  // The sentinels as route-sentinels.ts registers them: each draws
  // every character, at its own advance, wherever no named family before it
  // does. Generic families and system fallback draw everything else.
  const SENTINEL = [
    { width: 512, actualBoundingBoxAscent: 8192, fontBoundingBoxAscent: 716, fontBoundingBoxDescent: 204 },
    { width: 768, actualBoundingBoxAscent: 9216, fontBoundingBoxAscent: 614, fontBoundingBoxDescent: 307 },
  ];
  const context = (loaded = true) => ({
    font: '10px sans-serif',
    canvas: { ownerDocument: { fonts: {} } },
    measureText(text: string) {
      const named = namedWidth.call(this, text);
      if (named !== undefined) return { width: named };
      const families = this.font.slice(this.font.indexOf('px ') + 3).split(', ');
      for (const family of families) {
        const sentinel = loaded ? SENTINEL_FAMILIES.findIndex((name: string) => family === `"${name}"`) : -1;
        if (sentinel >= 0) {
          const scale = Number.parseFloat(/([\d.]+)px/.exec(this.font)![1]!) / 1024;
          const metric = SENTINEL[sentinel]!;
          // A default-ignorable character is drawn by no face.
          if (text === '­') return { width: 0 };
          return Object.fromEntries(Object.entries(metric).map(([key, value]) => [key, value * scale])) as { width: number };
        }
        if (family === 'serif') break;
      }
      return { width: fallbackWidth(text) };
    },
  });

  it('keeps a code the named families draw: the PUA form by itself, the bare form through its identity', () => {
    const oracle = createSymbolFaceOracle(context());
    const font = '700 20px "Sym", "Latin", serif';
    expect([oracle.draws('A', font), oracle.draws(ch(0xf041), font)]).toEqual([true, true]);
    // "a": the Latin face draws it, as a Latin letter; U+F061: no named family.
    expect([oracle.draws('a', font), oracle.draws(ch(0xf061), font)]).toEqual([false, false]);
    // A face that maps only the PUA form keeps it; its bare form converts.
    expect([oracle.draws('A', '20px "Pua", "Latin", serif'), oracle.draws(ch(0xf041), '20px "Pua", "Latin", serif')])
      .toEqual([false, true]);
    // A later named family draws what the first lacks.
    expect(oracle.draws(ch(0xf041), '20px "Other", "Pua", serif')).toBe(true);
    // Nothing after the first generic family is a named face of the route.
    expect(oracle.draws(ch(0xf041), '20px "Other", serif, "Sym"')).toBe(false);
    expect(oracle.draws(ch(0xf041), '20px serif')).toBe(false);
    expect(oracle.draws(ch(0xad), font)).toBe(false);
  });

  it('judges nothing in a font realm whose sentinels are not loaded', () => {
    const oracle = createSymbolFaceOracle(context(false));
    expect(oracle.draws('A', '20px "Sym"')).toBeUndefined();
    expect(oracle.segments(ch(0xf0b7) + ch(0xf061), 'Symbol', '20px "Other"'))
      .toEqual([{ text: ch(0x2022), mapped: true }, { text: ch(0xf061), mapped: false }]);
  });

  it('keeps no verdict while the sentinels do not work, and judges once they do', () => {
    let loaded = false;
    const ctx = context(true);
    const measure = ctx.measureText;
    const off = context(false).measureText;
    ctx.measureText = function (text: string) { return (loaded ? measure : off).call(this, text); };
    const oracle = createSymbolFaceOracle(ctx);
    expect(oracle.draws(ch(0xf041), '20px "Sym"')).toBeUndefined();
    loaded = true;
    expect(oracle.draws(ch(0xf041), '20px "Sym"')).toBe(true);
  });
});

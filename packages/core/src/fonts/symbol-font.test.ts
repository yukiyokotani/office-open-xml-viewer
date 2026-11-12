import { describe, expect, it } from 'vitest';
import {
  SYMBOL_FACE_UNICODE,
  SYMBOL_FACE_UNKNOWN,
  isSymbolFontFamily,
  symbolCodeSet,
  symbolFontToUnicode,
  symbolTextToUnicodeSegments,
  type SymbolFaceEncoding,
} from './symbol-font';
import {
  createSymbolFaceOracle,
  createSymbolMonochromeProbe,
  measureSymbolFaceEncoding,
  recordHeldSymbolEncoding,
  symbolCmapEncoding,
  symbolEncodingOfFaces,
} from './symbol-face';
import { ADOBE_SYMBOL_UNICODE } from './symbol-encoding';
import { symbolFixtureFontBytes } from '../testing/symbol-font-fixture';

const ch = (code: number): string => String.fromCodePoint(code);
const VS15 = '\uFE0E';
const missing = { encoding: SYMBOL_FACE_UNICODE };
const mono = { encoding: SYMBOL_FACE_UNICODE, monochrome: () => true };
const owning = (codes: number[]): SymbolFaceEncoding => ({
  kind: 'known', pua: symbolCodeSet(codes), bare: symbolCodeSet(codes),
});

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

describe('symbol text conversion for a face that draws the codes', () => {
  it('keeps every code the selected face owns, per form', () => {
    const encoding = owning([0x61, 0x62]);
    expect(symbolTextToUnicodeSegments(ch(0xf061) + 'b' + ch(0xf063), 'Symbol', { encoding })).toEqual([
      { text: ch(0xf061) + 'b', mapped: false }, { text: ch(0x3c7), mapped: true },
    ]);
    // A face that maps only the PUA form still converts the bare form.
    const puaOnly: SymbolFaceEncoding = { kind: 'known', pua: symbolCodeSet([0x61]), bare: symbolCodeSet([]) };
    expect(symbolFontToUnicode(ch(0xf061) + 'a', 'Symbol', { encoding: puaOnly })).toBe(ch(0xf061) + ch(0x3b1));
  });

  it('keeps the 0.89 subset when the face could not be judged', () => {
    expect(symbolFontToUnicode(ch(0xf0b7) + ch(0xf061), 'Symbol', { encoding: SYMBOL_FACE_UNKNOWN }))
      .toBe(ch(0x2022) + ch(0xf061));
    expect(symbolFontToUnicode(ch(0xf0e0), 'Wingdings')).toBe(ch(0x2192));
    expect(symbolFontToUnicode(ch(0xf0a7), 'Webdings')).toBe(ch(0xf0a7));
  });
});

describe('symbolCmapEncoding', () => {
  it('reads a symbol cmap exactly: PUA entries, and bare codes through the remap', () => {
    const encoding = symbolCmapEncoding(symbolFixtureFontBytes({ cmap: 'symbol', first: 0x40, last: 0x7a }));
    expect(encoding).toEqual(owning(Array.from({ length: 0x7a - 0x40 + 1 }, (_, i) => 0x40 + i)));
    expect(symbolFontToUnicode(ch(0xf061) + ch(0xf07b), 'Symbol', { encoding: encoding! })).toBe(ch(0xf061) + '{');
  });

  it('treats a Unicode cmap as drawing no symbol code', () => {
    const latin = Array.from({ length: 0x7e - 0x20 + 1 }, (_, i) => 0x20 + i);
    expect(symbolCmapEncoding(symbolFixtureFontBytes({ cmap: 'unicode', first: 0x20, last: 0x7e })))
      .toEqual({ ...SYMBOL_FACE_UNICODE, latin: symbolCodeSet(latin) });
    expect(symbolCmapEncoding(new Uint8Array(8))).toBeNull();
  });

  it('never owns a default-ignorable bare code, though its PUA form is mapped', () => {
    // Shapers draw nothing for a bare SOFT HYPHEN, even through a symbol cmap.
    const encoding = symbolCmapEncoding(symbolFixtureFontBytes({ cmap: 'symbol', first: 0xa0, last: 0xb0 }))!;
    expect(symbolFontToUnicode('\u00AD\uF0AD', 'Symbol', { encoding })).toBe('\u2191\uF0AD');
  });

  it('leaves a face with both a symbol and a Unicode subtable to measurement', () => {
    // Engines disagree on which subtable they read, so the bytes cannot decide.
    expect(symbolCmapEncoding(symbolFixtureFontBytes({ cmap: 'both', first: 0x40, last: 0x7a }))).toBeNull();
  });
});

describe('symbolEncodingOfFaces', () => {
  const symbolFace = (unicodeRange: string, weight = 'normal', style = 'normal') => ({
    weight, style, unicodeRange, symbolEncoding: owning([0x61, 0x62]),
  });

  it('owns a scalar only through the last matching face whose unicode-range covers it', () => {
    const encoding = symbolEncodingOfFaces([symbolFace('U+F000-F0FF')], { weight: 400, style: 'normal' })!;
    // PUA-only range: the PUA form is drawn, the bare form falls back. A later
    // Unicode face covering U+F062 lacks it, so the earlier face still draws it.
    expect(symbolFontToUnicode(ch(0xf061) + 'a', 'Symbol', { encoding })).toBe(ch(0xf061) + ch(0x3b1));
    const later = symbolEncodingOfFaces([
      symbolFace('U+0-10FFFF'), { weight: 'normal', style: 'normal', unicodeRange: 'U+F062', symbolEncoding: SYMBOL_FACE_UNICODE },
    ], { weight: 400, style: 'normal' })!;
    expect(symbolFontToUnicode(ch(0xf061) + ch(0xf062), 'Symbol', { encoding: later })).toBe(ch(0xf061) + ch(0xf062));
  });

  it('selects faces by CSS font matching, synthesis included (CSS Fonts 4 §5.2)', () => {
    // A regular-only face paints bold and italic text (synthesized), as does
    // a condensed-only face normal-width text.
    for (const desired of [{ weight: 700, style: 'normal' as const }, { weight: 400, style: 'italic' as const }]) {
      expect(symbolEncodingOfFaces([symbolFace('U+0-10FFFF')], desired)?.bare).toBe(symbolCodeSet([0x61, 0x62]));
    }
    expect(symbolEncodingOfFaces([{ ...symbolFace('U+0-10FFFF'), stretch: 'condensed' }], { weight: 400, style: 'normal' }))
      .toBeDefined();
    // The nearest face wins: bold text takes the bold Unicode face, not the
    // regular symbol face; weight 400 looks up to 500 before going lighter.
    const unicode = (weight: string) => ({ weight, style: 'normal', unicodeRange: 'U+0-10FFFF',
      symbolEncoding: { kind: 'known', pua: symbolCodeSet([]), bare: symbolCodeSet([]), latin: symbolCodeSet([0x61]) } as const });
    expect(symbolEncodingOfFaces([symbolFace('U+0-10FFFF'), unicode('700')], { weight: 700, style: 'normal' })?.bare)
      .toBe(symbolCodeSet([]));
    expect(symbolEncodingOfFaces([symbolFace('U+0-10FFFF', '300'), unicode('500')], { weight: 400, style: 'normal' })?.bare)
      .toBe(symbolCodeSet([]));
    // Italic text prefers an italic face over a regular one.
    expect(symbolEncodingOfFaces([symbolFace('U+0-10FFFF'), { ...unicode('400'), style: 'italic' }],
      { weight: 400, style: 'italic' })?.bare).toBe(symbolCodeSet([]));
  });

  it('matches oblique angles as slopes and declines ties between different descriptors', () => {
    const unicodeA = { kind: 'known', pua: symbolCodeSet([]), bare: symbolCodeSet([]), latin: symbolCodeSet([0x61]) } as const;
    // Italic text takes the steeper oblique face (slope 20 before 5).
    const steep = { ...symbolFace('U+0-10FFFF'), style: 'oblique 20deg' };
    const shallow = { weight: 'normal', style: 'oblique 5deg', unicodeRange: 'U+0-10FFFF', symbolEncoding: unicodeA };
    expect(symbolEncodingOfFaces([steep, shallow], { weight: 400, style: 'italic' })?.bare).toBe(symbolCodeSet([0x61, 0x62]));
    // Two weight ranges that both contain 700: engines break the tie their own way.
    expect(symbolEncodingOfFaces([symbolFace('U+0-10FFFF', '400 900'), { ...shallow, style: 'normal', weight: '400 700' }],
      { weight: 700, style: 'normal' })).toBeUndefined();
  });

  it('declines a selected face without read bytes, or one not loaded', () => {
    expect(symbolEncodingOfFaces([{ weight: 'normal', style: 'normal', unicodeRange: 'U+0-10FFFF' }], { weight: 400, style: 'normal' }))
      .toBeUndefined();
    expect(symbolEncodingOfFaces([{ ...symbolFace('U+0-10FFFF'), status: 'loading' }], { weight: 400, style: 'normal' }))
      .toBeUndefined();
  });

  it('looks past a later face that has no glyph for a bare code', () => {
    // The later Unicode face maps only "B": "a" falls through to the symbol face.
    const later = { weight: 'normal', style: 'normal', unicodeRange: 'U+0-10FFFF',
      symbolEncoding: { ...SYMBOL_FACE_UNICODE, latin: symbolCodeSet([0x42]) } };
    const encoding = symbolEncodingOfFaces([symbolFace('U+0-10FFFF'), later], { weight: 400, style: 'normal' })!;
    expect(symbolFontToUnicode('ab', 'Symbol', { encoding })).toBe('ab');
    const shadowing = { ...later, symbolEncoding: { ...SYMBOL_FACE_UNICODE, latin: symbolCodeSet([0x61]) } };
    expect(symbolFontToUnicode('a', 'Symbol', {
      encoding: symbolEncodingOfFaces([symbolFace('U+0-10FFFF'), shadowing], { weight: 400, style: 'normal' })!,
    })).toBe(ch(0x3b1));
  });
});

describe('createSymbolMonochromeProbe', () => {
  it('draws on another canvas of the painting canvas class (Node canvases)', () => {
    let created = 0;
    class NodeCanvas {
      constructor(readonly width: number, readonly height: number) { created++; }
      getContext() {
        return {
          clearRect() {}, fillText() {}, font: '', fillStyle: '', textBaseline: '',
          globalAlpha: 1, globalCompositeOperation: '',
          getImageData: () => ({ data: new Uint8ClampedArray([0, 0, 0, 255]) }),
        };
      }
    }
    const owner = { canvas: new NodeCanvas(1, 1) };
    expect(createSymbolMonochromeProbe(owner, '32px sans-serif')('\u2665\uFE0E')).toBe(true);
    expect(created).toBe(2);
  });
});

describe('oracle colour proof', () => {
  it('requires the base without VS15 to be monochrome too (a shaper may split them)', () => {
    let last = '';
    class Canvas {
      constructor(readonly width: number, readonly height: number) {}
      getContext() {
        return {
          clearRect() {}, font: '', fillStyle: '', textBaseline: '', globalAlpha: 1, globalCompositeOperation: '',
          fillText(text: string) { last = text; },
          // Colour unless the text carries VS15.
          getImageData: () => ({ data: new Uint8ClampedArray(last.includes('\uFE0E') ? [0, 0, 0, 255] : [200, 30, 30, 255]) }),
        };
      }
    }
    const ctx = { font: '', canvas: new Canvas(1, 1), measureText: (text: string) => ({ width: text === 'i' ? 3 : 7 }) };
    const oracle = createSymbolFaceOracle(ctx);
    const parts = oracle.segments('\uF0A9', 'Symbol', { font: '400 32px "Absent"' });
    expect(parts).toEqual([{ text: '\uF0A9', mapped: false }]);
  });
});

/** A registered FontFace stand-in (CSS Font Loading 3 descriptors). */
const registered = (family: string, unicodeRange: string, held?: SymbolFaceEncoding | null) => {
  const face = { family, status: 'loaded', weight: 'normal', style: 'normal', stretch: 'normal', unicodeRange };
  if (held !== undefined) recordHeldSymbolEncoding(face as unknown as FontFace, held);
  return face;
};

describe('which face wins a symbol code (held cmaps and measurement, one oracle)', () => {
  // A context in which no face draws anything of its own: every measured
  // family owns no code (a missing face, or a Unicode one).
  const realm = (faces: object[]) => {
    const fonts = { *[Symbol.iterator]() { yield* faces; } };
    return {
      font: '', canvas: { ownerDocument: { fonts } },
      measureText: (text: string) => ({ width: text.length + (text === 'i' ? 1 : 2) }),
    };
  };
  const joined = (parts: { text: string }[]) => parts.map((part) => part.text).join('');

  it('reads a held substitute through its unicode-range, deciding bare and PUA codes apart', () => {
    // The selected face is missing (measured); the loaded substitute's bytes
    // are symbol-encoded but declared for U+F000-F0FF only: it draws U+F061,
    // while a bare "a" never reaches it and converts.
    const ctx = realm([registered('Held PUA Substitute', 'U+F000-F0FF', owning([0x61]))]);
    const oracle = createSymbolFaceOracle(ctx);
    const selection = { font: '400 32px "Absent Primary", "Held PUA Substitute"' };
    expect(joined(oracle.segments(ch(0xf061) + 'a', 'Symbol', selection))).toBe(ch(0xf061) + ch(0x3b1));
  });

  it('measures the rest of the stack as it paints after a family the bytes cannot answer for', () => {
    // The missing primary is measured together with the substitute it falls
    // through to: a substitute painting both forms of 0x61 keeps a bare "a".
    const ctx = {
      font: '', canvas: { ownerDocument: { fonts: { *[Symbol.iterator]() {} } } },
      measureText(text: string) {
        const code = text.codePointAt(0)!;
        const owned = this.font.includes('"Full Range Substitute"') && (code === 0x61 || code === 0xf061);
        return { width: owned ? 100 : code >= 0xf000 ? 1 : text === 'i' ? 3 : 5 + (code % 7) };
      },
    };
    const oracle = createSymbolFaceOracle(ctx);
    const selection = { font: '400 32px "Absent Primary", "Full Range Substitute"' };
    expect(joined(oracle.segments(ch(0xf061) + 'ab', 'Symbol', selection))).toBe(ch(0xf061) + 'a' + ch(0x3b2));
  });

  it('re-derives when a registered face moves to another selected family', () => {
    const later = registered('Moved B', 'U+0-10FFFF', owning([0x61]));
    const ctx = realm([registered('Moved A', 'U+0-10FFFF', { kind: 'known', pua: symbolCodeSet([]),
      bare: symbolCodeSet([]), latin: symbolCodeSet([0x61]) }), later]);
    const oracle = createSymbolFaceOracle(ctx);
    const selection = { font: '400 32px "Moved A", "Moved B"' };
    // A's Unicode face draws "a" as the letter: it converts.
    expect(joined(oracle.segments('a', 'Symbol', selection))).toBe(ch(0x3b1));
    // The symbol face, now declared later under A, wins "a" there.
    later.family = 'Moved A';
    expect(joined(oracle.segments('a', 'Symbol', selection))).toBe('a');
  });

  it('walks the whole painted family list, as font matching does', () => {
    // The selected face is missing; a named fallback later in the list (an
    // embedded alias) holds symbol bytes and paints the code.
    const ctx = realm([registered('Embedded Symbol Alias', 'U+0-10FFFF', owning([0x61]))]);
    const oracle = createSymbolFaceOracle(ctx);
    const font = { font: 'normal 400 32px "Absent Primary", "Embedded Symbol Alias", sans-serif' };
    expect(joined(oracle.segments(ch(0xf061), 'Symbol', font))).toBe(ch(0xf061));
    expect(joined(oracle.segments(ch(0xf061), 'Symbol', { font: '400 32px "Absent Primary", sans-serif' }))).toBe(ch(0x3b1));
  });

  it('reads a regular-only held face for bold and italic text (synthesis)', () => {
    const ctx = realm([registered('Regular Only Symbol', 'U+0-10FFFF', owning([0x61]))]);
    const oracle = createSymbolFaceOracle(ctx);
    for (const font of ['bold 32px "Regular Only Symbol"', 'italic 700 32px "Regular Only Symbol"']) {
      expect(joined(oracle.segments('a' + ch(0xf061), 'Symbol', { font }))).toBe('a' + ch(0xf061));
    }
  });

  it('names a registered face by its literal family, matched caselessly', () => {
    const faces: object[] = [];
    const ctx = realm(faces);
    const oracle = createSymbolFaceOracle(ctx);
    const caseless = { font: '32px "s\u00FDmbol caseless"' };
    expect(joined(oracle.segments('a', 'Symbol', caseless))).toBe(ch(0x3b1));
    faces.push(registered('S\u00DDMBOL CASELESS', 'U+0-10FFFF', owning([0x61])));
    expect(joined(oracle.segments('a', 'Symbol', caseless))).toBe('a');
    // A programmatic face named with literal quotes is not the family "Quoted Only".
    faces.push(registered('"Quoted Only"', 'U+0-10FFFF', owning([0x61])));
    expect(joined(oracle.segments('a', 'Symbol', { font: '32px "Quoted Only"' }))).toBe(ch(0x3b1));
    // An unquoted name decodes its escapes.
    faces.push(registered('Real Escaped', 'U+0-10FFFF', owning([0x61])));
    expect(joined(oracle.segments('a', 'Symbol', { font: '32px Real\\20 Escaped' }))).toBe('a');
  });

  it('measures at the width the font asks for', () => {
    const fonts: string[] = [];
    const ctx = { get font() { return fonts.at(-1) ?? ''; }, set font(value: string) { fonts.push(value); },
      measureText: (text: string) => ({ width: text.length + (text === 'i' ? 1 : 2) }) };
    createSymbolFaceOracle(ctx).encoding({ font: 'condensed 32px "Width Probe"' });
    expect(fonts.some((font) => font.includes('condensed') && font.includes('"Width Probe"'))).toBe(true);
  });

  it('re-derives a held answer when a face registered later under the family wins the codes', () => {
    const faces: object[] = [registered('Held Symbol Face', 'U+0-10FFFF', owning([0x61]))];
    const ctx = realm(faces);
    const selection = { font: '400 32px "Held Symbol Face"' };
    const text = () => joined(createSymbolFaceOracle(ctx, { shared: true }).segments('a', 'Symbol', selection));
    expect(text()).toBe('a');
    // A later application face whose bytes the library did not read: measured.
    faces.push(registered('Held Symbol Face', 'U+0-10FFFF'));
    expect(text()).toBe(ch(0x3b1));
    // A later held Unicode face that maps "a" draws it as the letter.
    faces[1] = registered('Held Symbol Face', 'U+0-10FFFF', { kind: 'known', pua: symbolCodeSet([]), bare: symbolCodeSet([]), latin: symbolCodeSet([0x61]) });
    expect(text()).toBe(ch(0x3b1));
    // Without it the held symbol face wins again.
    faces.pop();
    expect(text()).toBe('a');
  });
});

describe('measureSymbolFaceEncoding', () => {
  // A context that draws the "Sym" face's codes 0x41..0x5A at width 2+code
  // (both forms) and every other text at one width per character class. In
  // a right-to-left direction a bare code is mirrored (another width), as
  // "(" is, so the identity holds only when measured left to right.
  const context = (owned: (code: number) => boolean) => ({
    font: '10px sans-serif',
    direction: 'rtl' as CanvasDirection,
    measureText(text: string) {
      const code = text.codePointAt(0)!;
      const symbolFace = this.font.includes('"Sym"');
      const base = code >= 0xf000 ? code - 0xf000 : code;
      if (symbolFace && owned(base)) return { width: 2 + base + (this.direction === 'rtl' && code < 0xf000 ? 1 : 0) };
      return { width: code >= 0xf000 ? 1 : text === 'i' ? 3 : 5 + (code % 7) };
    },
  });

  it('owns exactly the codes whose bare and PUA forms draw identically in the face', () => {
    const ctx = context((code) => code >= 0x41 && code <= 0x5a);
    const letters = Array.from({ length: 26 }, (_, i) => 0x41 + i);
    expect(measureSymbolFaceEncoding(ctx, '"Sym"')).toEqual({
      kind: 'known', pua: symbolCodeSet([...letters, 0x7f, 0xad]), bare: symbolCodeSet(letters),
    });
    expect(measureSymbolFaceEncoding(ctx, '"Other"')).toEqual(SYMBOL_FACE_UNICODE);
    expect([ctx.font, ctx.direction]).toEqual(['10px sans-serif', 'rtl']);
  });

  it('is inconclusive without a context or with one that measures nothing', () => {
    expect(measureSymbolFaceEncoding(null, '"Sym"')).toBe(SYMBOL_FACE_UNKNOWN);
    expect(measureSymbolFaceEncoding({ font: '', measureText: () => ({ width: 0 }) }, '"Sym"')).toBe(SYMBOL_FACE_UNKNOWN);
  });

  it('shares measurements across oracles of one realm until the measured families change', () => {
    let calls = 0;
    const faces: object[] = [];
    const fonts = { *[Symbol.iterator]() { yield* faces; } };
    const ctx = {
      font: '', canvas: { ownerDocument: { fonts } },
      measureText: (text: string) => { calls++; return { width: text.length + (text === 'i' ? 1 : 2) }; },
    };
    const selection = { font: '400 32px "Shared Probe Face"' };
    createSymbolFaceOracle(ctx, { shared: true }).encoding(selection);
    const first = calls;
    createSymbolFaceOracle(ctx, { shared: true }).encoding(selection);
    expect(calls).toBe(first);
    // A face registered (or replaced) for the measured family re-measures.
    faces.push(registered('Shared Probe Face', 'U+0-10FFFF'));
    createSymbolFaceOracle(ctx, { shared: true }).encoding(selection);
    expect(calls).toBe(first * 2);
    faces[0] = registered('Shared Probe Face', 'U+0-10FFFF');
    createSymbolFaceOracle(ctx, { shared: true }).encoding(selection);
    expect(calls).toBe(first * 3);
  });

  it('caches measurements per selection', () => {
    let calls = 0;
    const ctx = { font: '', measureText: (text: string) => { calls++; return { width: text.length + (text === 'i' ? 1 : 2) }; } };
    const oracle = createSymbolFaceOracle(ctx);
    oracle.encoding({ font: '400 32px "A"' });
    const measured = calls;
    expect(measured).toBeGreaterThan(0);
    oracle.encoding({ font: '400 32px "A"' });
    expect(calls).toBe(measured);
  });

});

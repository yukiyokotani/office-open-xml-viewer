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
import { createSymbolFaceOracle, measureSymbolFaceEncoding, symbolCmapEncoding, symbolEncodingOfFaces } from './symbol-face';
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
    expect(symbolCmapEncoding(symbolFixtureFontBytes({ cmap: 'unicode', first: 0x20, last: 0x7e })))
      .toEqual(SYMBOL_FACE_UNICODE);
    expect(symbolCmapEncoding(new Uint8Array(8))).toBeNull();
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
    const encoding = symbolEncodingOfFaces([symbolFace('U+F000-F0FF')], 400, 'normal')!;
    // PUA-only range: the PUA form is drawn, the bare form falls back. A later
    // Unicode face covering U+F062 lacks it, so the earlier face still draws it.
    expect(symbolFontToUnicode(ch(0xf061) + 'a', 'Symbol', { encoding })).toBe(ch(0xf061) + ch(0x3b1));
    const later = symbolEncodingOfFaces([
      symbolFace('U+0-10FFFF'), { weight: 'normal', style: 'normal', unicodeRange: 'U+F062', symbolEncoding: SYMBOL_FACE_UNICODE },
    ], 400, 'normal')!;
    expect(symbolFontToUnicode(ch(0xf061) + ch(0xf062), 'Symbol', { encoding: later })).toBe(ch(0xf061) + ch(0xf062));
  });

  it('declines a tuple no face matches exactly, or a face without read bytes', () => {
    expect(symbolEncodingOfFaces([symbolFace('U+0-10FFFF')], 700, 'normal')).toBeUndefined();
    expect(symbolEncodingOfFaces([{ weight: 'normal', style: 'normal', unicodeRange: 'U+0-10FFFF' }], 400, 'normal'))
      .toBeUndefined();
  });
});

describe('measureSymbolFaceEncoding', () => {
  // A context that draws the "Sym" face's codes 0x41..0x5A at width 2+code
  // (both forms) and every other text at one width per character class.
  const context = (owned: (code: number) => boolean) => ({
    font: '10px sans-serif',
    measureText(text: string) {
      const code = text.codePointAt(0)!;
      const symbolFace = this.font.includes('"Sym"');
      const base = code >= 0xf000 ? code - 0xf000 : code;
      if (symbolFace && owned(base)) return { width: 2 + base };
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
    expect(ctx.font).toBe('10px sans-serif');
  });

  it('is inconclusive without a context or with one that measures nothing', () => {
    expect(measureSymbolFaceEncoding(null, '"Sym"')).toBe(SYMBOL_FACE_UNKNOWN);
    expect(measureSymbolFaceEncoding({ font: '', measureText: () => ({ width: 0 }) }, '"Sym"')).toBe(SYMBOL_FACE_UNKNOWN);
  });

  it('prefers held bytes over measurement and caches per selection', () => {
    let calls = 0;
    const ctx = { font: '', measureText: (text: string) => { calls++; return { width: text.length + (text === 'i' ? 1 : 2) }; } };
    const oracle = createSymbolFaceOracle(ctx);
    const exact = owning([0x61]);
    expect(oracle.encoding({ familyText: '"A"', weight: 400, style: 'normal', encoding: exact })).toBe(exact);
    expect(calls).toBe(0);
    oracle.encoding({ familyText: '"A"', weight: 400, style: 'normal' });
    const measured = calls;
    oracle.encoding({ familyText: '"A"', weight: 400, style: 'normal' });
    expect(calls).toBe(measured);
  });
});

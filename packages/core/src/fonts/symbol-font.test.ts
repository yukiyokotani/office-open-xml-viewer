import { describe, expect, it } from 'vitest';
import {
  SYMBOL_MAP,
  symbolFontToUnicode,
  isSymbolFontFamily,
  symbolTextToUnicodeSegments,
} from './symbol-font';

const ch = (code: number): string => String.fromCharCode(code);

describe('symbolFontToUnicode', () => {
  // The load-bearing case: list bullets. Word serializes a Symbol bullet as the
  // PUA code point U+F0B7 and a Wingdings square as U+F0A7; both the bare and
  // PUA-shifted forms must normalize so the marker draws in any fallback face.
  it('maps Symbol / Wingdings bullet code points (bare + PUA) to Unicode', () => {
    expect(symbolFontToUnicode(ch(0xf0b7), 'Symbol')).toBe('•');
    expect(symbolFontToUnicode(ch(0xb7), 'Symbol')).toBe('•');
    expect(symbolFontToUnicode(ch(0xf0a7), 'Wingdings')).toBe('▪');
    expect(symbolFontToUnicode(ch(0xa7), 'Wingdings')).toBe('▪');
  });

  // Each family has its own encoding, so the same code point is a different
  // character in each table (WG2 N4384 section 7 for the Wingdings/Webdings
  // families, Adobe symbol.txt for Symbol).
  it('resolves the same code point through each family\'s own table', () => {
    expect(symbolFontToUnicode(ch(0xf0a7), 'Symbol')).toBe('♣');
    expect(symbolFontToUnicode(ch(0xf0a7), 'Wingdings')).toBe('▪');
    expect(symbolFontToUnicode(ch(0xf0a7), 'Wingdings 2')).toBe('\u{1F794}');
    expect(symbolFontToUnicode(ch(0xf0a7), 'Wingdings 3')).toBe('\u{1F824}');
    expect(symbolFontToUnicode(ch(0xf0a7), 'Webdings')).toBe('\u{1F56E}');
  });

  it('maps representative Wingdings codes, bare and F0-shifted', () => {
    const both = (code: number, expected: string) => {
      expect(symbolFontToUnicode(ch(code), 'Wingdings')).toBe(expected);
      expect(symbolFontToUnicode(ch(0xf000 + code), 'Wingdings')).toBe(expected);
    };
    both(0x21, '\u{1F589}'); // pencil
    both(0x4a, '☺'); // the "J" smiley
    both(0x71, '❑'); // list box bullet
    both(0x76, '❖');
    both(0xb7, '\u{1F550}'); // clock face one o'clock
    both(0xd8, '⮚'); // list arrowhead bullet
    both(0xe0, '\u{1F86A}'); // wide-headed rightwards barb arrow
    both(0xfc, '✓');
    both(0xfe, '\u{1F5F9}'); // ballot box with bold check
  });

  it('maps Wingdings 2, Wingdings 3 and Webdings codes', () => {
    expect(symbolFontToUnicode(ch(0xf050), 'Wingdings 2')).toBe('\u{1F5F8}');
    expect(symbolFontToUnicode(ch(0x26), 'Wingdings 2')).toBe('\u2700'); // safety scissors
    expect(symbolFontToUnicode(ch(0x21), 'Wingdings 3')).toBe('\u2B60');
    expect(symbolFontToUnicode(ch(0xf021), 'Webdings')).toBe('\u{1F577}'); // spider
    expect(symbolFontToUnicode(ch(0xff), 'Webdings')).toBe('\u{1F54A}'); // dove
  });

  // Codes the source gives no Unicode character for keep the authored face.
  it('passes through codes with no Unicode character', () => {
    expect(symbolFontToUnicode(ch(0xf0ff), 'Wingdings')).toBe(ch(0xf0ff)); // Windows logo
    expect(symbolFontToUnicode(ch(0xf0fa), 'Wingdings 2')).toBe(ch(0xf0fa));
    expect(symbolFontToUnicode(ch(0xf0f1), 'Wingdings 3')).toBe(ch(0xf0f1));
    expect(symbolFontToUnicode(' ', 'Webdings')).toBe(' ');
    expect(symbolFontToUnicode(ch(0xf0f0), 'Symbol')).toBe(ch(0xf0f0));
  });

  it('maps Symbol arrows from the Adobe encoding', () => {
    // symbol.txt: 0xAE → U+2192 RIGHTWARDS ARROW, 0xAC → U+2190 LEFTWARDS ARROW
    expect(symbolFontToUnicode(ch(0xae), 'Symbol')).toBe('→');
    expect(symbolFontToUnicode(ch(0xac), 'Symbol')).toBe('←');
  });

  // The gate is keyed on the requested font family, not the character — a normal
  // body font must never be remapped (a real "·" / "§" / "a" stays itself).
  it('passes characters through for non-symbol fonts', () => {
    expect(symbolFontToUnicode(ch(0xb7), 'Calibri')).toBe(ch(0xb7));
    expect(symbolFontToUnicode(ch(0xa7), 'Times New Roman')).toBe(ch(0xa7));
    expect(symbolFontToUnicode('a', 'Calibri')).toBe('a');
    expect(symbolFontToUnicode(ch(0xf0b7), null)).toBe(ch(0xf0b7));
    expect(symbolFontToUnicode(ch(0xf0b7), undefined)).toBe(ch(0xf0b7));
    // Table selection is exact (trimmed, case-insensitive).
    expect(symbolFontToUnicode(ch(0xf0a7), '  WINGDINGS  ')).toBe('▪');
    expect(symbolFontToUnicode(ch(0xf0a7), 'Wingdings 4')).toBe(ch(0xf0a7));
  });

  // In Symbol, 0x21 is "!", 0x22 is "∀" and 0x24 is "∃" — Wingdings dingbats
  // must NOT leak across the font gate.
  it('keeps Wingdings dingbats out of the Symbol table', () => {
    expect(symbolFontToUnicode(ch(0x21), 'Symbol')).toBe('!');
    expect(symbolFontToUnicode(ch(0x22), 'Symbol')).toBe('∀');
    expect(symbolFontToUnicode(ch(0x24), 'Symbol')).toBe('∃');
  });
});

// Issue #2177: the whole Adobe Symbol repertoire (unicode.org VENDORS/ADOBE/
// symbol.txt), not only bullets and arrows.
describe('symbolFontToUnicode — full Symbol encoding', () => {
  const both = (code: number, expected: string) => {
    expect(symbolFontToUnicode(ch(code), 'Symbol')).toBe(expected);
    expect(symbolFontToUnicode(ch(0xf000 + code), 'Symbol')).toBe(expected);
  };

  it('maps Greek capitals and lowercase, bare (ASCII) and F0-shifted', () => {
    both(0x41, 'Α'); both(0x43, 'Χ'); both(0x44, 'Δ'); both(0x47, 'Γ');
    both(0x51, 'Θ'); both(0x53, 'Σ'); both(0x57, 'Ω'); both(0x5a, 'Ζ');
    both(0x61, 'α'); both(0x62, 'β'); both(0x63, 'χ'); both(0x6c, 'λ');
    both(0x6d, 'μ'); both(0x70, 'π'); both(0x77, 'ω'); both(0x7a, 'ζ');
    // Symbol's two phis: 0x66 is the stroked phi, 0x6A the open, looped one.
    both(0x66, 'ϕ'); both(0x6a, 'φ');
    // Every letter slot resolves, none passes through as Latin.
    for (let code = 0x41; code <= 0x5a; code++) {
      expect(SYMBOL_MAP[code]).toMatch(/\p{Script=Greek}/u);
      expect(SYMBOL_MAP[code + 0x20]).toMatch(/\p{Script=Greek}/u);
    }
  });

  it('maps operators, relations and set symbols', () => {
    both(0x22, '∀'); both(0x24, '∃'); both(0x2d, '−'); both(0x40, '≅');
    both(0x5c, '∴'); both(0x7e, '∼'); both(0xa5, '∞'); both(0xb6, '∂');
    both(0xbb, '≈'); both(0xc6, '∅'); both(0xce, '∈'); both(0xd1, '∇');
    both(0xd5, '∏'); both(0xd6, '√'); both(0xe5, '∑'); both(0xf2, '∫');
  });

  it('maps arrows, the euro and identity punctuation', () => {
    both(0xab, '↔'); both(0xae, '→'); both(0xbf, '↵'); both(0xde, '⇒');
    both(0xdb, '⇔'); both(0xa0, '€');
    both(0x28, '('); both(0x31, '1');
  });

  it('maps the Corporate-Use pieces to their standard Unicode characters', () => {
    both(0xe6, '⎛'); both(0xef, '⎪'); both(0xf4, '⎮'); both(0xfe, '⎭');
    both(0xbd, '⏐'); both(0xbe, '⎯');
    both(0xd2, '®'); both(0xe3, '©'); both(0xd4, '™');
    both(0x60, '‾'); // radicalex: the top bar Office draws
  });

  it('converts every code point of a multi-character string', () => {
    expect(symbolFontToUnicode('abcp', 'Symbol')).toBe('αβχπ');
    expect(symbolFontToUnicode('a.', 'Symbol')).toBe('α.');
  });
});

describe('isSymbolFontFamily', () => {
  it('matches exactly "symbol" / "webdings" and any "wingdings" family', () => {
    expect(isSymbolFontFamily('Symbol')).toBe(true);
    expect(isSymbolFontFamily('SYMBOL')).toBe(true);
    expect(isSymbolFontFamily('Wingdings')).toBe(true);
    expect(isSymbolFontFamily('Wingdings 2')).toBe(true);
    expect(isSymbolFontFamily('Webdings')).toBe(true);
  });

  it('does NOT match ordinary faces', () => {
    expect(isSymbolFontFamily('Calibri')).toBe(false);
    // "Symbol" must be exact, not a substring (e.g. "SymbolMT" is a real Latin
    // PostScript face), matching symbolFontToUnicode's own gate.
    expect(isSymbolFontFamily('SymbolMT')).toBe(false);
    expect(isSymbolFontFamily(null)).toBe(false);
    expect(isSymbolFontFamily(undefined)).toBe(false);
  });
});

describe('symbolTextToUnicodeSegments', () => {
  it('returns the whole string unmapped for a non-symbol font', () => {
    expect(symbolTextToUnicodeSegments('abc', 'Calibri')).toEqual([
      { text: 'abc', mapped: false },
    ]);
    expect(symbolTextToUnicodeSegments(ch(0xf0a7), null)).toEqual([
      { text: ch(0xf0a7), mapped: false },
    ]);
  });

  it('maps every Wingdings PUA glyph and flags the run as mapped', () => {
    // U+F0A7 (square) U+F0FC (check mark) → ▪ ✓
    const segs = symbolTextToUnicodeSegments(ch(0xf0a7) + ch(0xf0fc), 'Wingdings');
    expect(segs).toEqual([{ text: '▪✓', mapped: true }]);
  });

  it('splits at mapped/unmapped boundaries so each run is single-font', () => {
    // The space and the Windows logo (0xFF) have no Unicode character
    // (passthrough); 0xF0A7 maps to ▪.
    const segs = symbolTextToUnicodeSegments(' ' + ch(0xf0a7) + ch(0xf0ff), 'Wingdings');
    expect(segs).toEqual([
      { text: ' ', mapped: false },
      { text: '▪', mapped: true },
      { text: ch(0xf0ff), mapped: false },
    ]);
  });

  it('preserves astral targets by iterating code points', () => {
    // Wingdings 0x24 → U+1F453 EYEGLASSES (astral). Must survive as one glyph.
    const segs = symbolTextToUnicodeSegments(ch(0x24), 'Wingdings');
    expect(segs).toEqual([{ text: '\u{1F453}', mapped: true }]);
  });

  it('maps ASCII Symbol text and keeps identity codes in the Symbol face', () => {
    expect(symbolTextToUnicodeSegments('a+b 1', 'Symbol')).toEqual([
      { text: 'α', mapped: true },
      { text: '+', mapped: false },
      { text: 'β', mapped: true },
      { text: ' 1', mapped: false },
    ]);
  });

  it('handles the empty string', () => {
    expect(symbolTextToUnicodeSegments('', 'Wingdings')).toEqual([
      { text: '', mapped: false },
    ]);
  });
});

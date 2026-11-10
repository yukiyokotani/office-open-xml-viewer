import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  SYMBOL_MAP,
  symbolFontToUnicode,
  isSymbolFontFamily,
  symbolTextToUnicodeSegments,
  setSymbolFaceProbeForTesting,
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

  // The core reason for splitting into two tables: the SAME code point resolves
  // to a DIFFERENT glyph depending on the font, because Symbol and Wingdings use
  // different encodings. A single shared table cannot be correct for both.
  it('resolves the same code point differently for Symbol vs Wingdings', () => {
    // 0xA7: Adobe Symbol → U+2663 club; Wingdings cmap → U+25AA small square.
    expect(symbolFontToUnicode(ch(0xa7), 'Symbol')).toBe('♣');
    expect(symbolFontToUnicode(ch(0xa7), 'Wingdings')).toBe('▪');
    // 0xB7: Adobe Symbol → U+2022 bullet; Wingdings cmap → a clock face
    // (astral U+1F550), which we leave as passthrough rather than mismap.
    expect(symbolFontToUnicode(ch(0xb7), 'Symbol')).toBe('•');
    expect(symbolFontToUnicode(ch(0xb7), 'Wingdings')).toBe(ch(0xb7));
    // 0xB8: Adobe Symbol → U+00F7 division; Wingdings → a clock face (passthrough).
    expect(symbolFontToUnicode(ch(0xb8), 'Symbol')).toBe('÷');
    expect(symbolFontToUnicode(ch(0xb8), 'Wingdings')).toBe(ch(0xb8));
  });

  // Restored ambiguous entries (removed in PR #519) now carry the correct
  // font-specific value rather than a single guessed glyph.
  it('restores the previously-removed code points with font-specific values', () => {
    // Symbol side
    expect(symbolFontToUnicode(ch(0xb8), 'Symbol')).toBe('÷'); // U+00F7
    expect(symbolFontToUnicode(ch(0xb9), 'Symbol')).toBe('≠'); // U+2260
    // Wingdings side
    expect(symbolFontToUnicode(ch(0x24), 'Wingdings')).toBe('\u{1F453}'); // eyeglasses
    expect(symbolFontToUnicode(ch(0x4b), 'Wingdings')).toBe('\u{1F610}'); // neutral face
    expect(symbolFontToUnicode(ch(0x76), 'Wingdings')).toBe('❖'); // U+2756
    expect(symbolFontToUnicode(ch(0xfe), 'Wingdings')).toBe('☑'); // U+2611 ballot box w/ check
  });

  it('maps the Wingdings barb arrow block', () => {
    expect(symbolFontToUnicode(ch(0xf0e0), 'Wingdings')).toBe('→');
    expect(symbolFontToUnicode(ch(0xdf), 'Wingdings')).toBe('←');
  });

  // Geometric markers. 0x74 lozenge6 / 0x77 rhombus4 are both SOLID glyphs —
  // both must be the BLACK diamond, never the white ◇ (regression guard for the
  // 0x77 fill error).
  it('maps the Wingdings solid geometric markers to black glyphs', () => {
    expect(symbolFontToUnicode(ch(0x6c), 'Wingdings')).toBe('●');
    expect(symbolFontToUnicode(ch(0x6e), 'Wingdings')).toBe('■');
    expect(symbolFontToUnicode(ch(0x74), 'Wingdings')).toBe('◆');
    expect(symbolFontToUnicode(ch(0x77), 'Wingdings')).toBe('◆');
  });

  it('maps Symbol arrows from the Adobe encoding', () => {
    // symbol.txt: 0xAE → U+2192 RIGHTWARDS ARROW, 0xAC → U+2190 LEFTWARDS ARROW
    expect(symbolFontToUnicode(ch(0xae), 'Symbol')).toBe('→');
    expect(symbolFontToUnicode(ch(0xac), 'Symbol')).toBe('←');
  });

  // The gate is keyed on the requested font family, not the character — a normal
  // body font must never be remapped (a real "·" / "§" / "A" stays itself).
  it('passes characters through for non-symbol fonts', () => {
    expect(symbolFontToUnicode(ch(0xb7), 'Calibri')).toBe(ch(0xb7));
    expect(symbolFontToUnicode(ch(0xa7), 'Times New Roman')).toBe(ch(0xa7));
    expect(symbolFontToUnicode('a', 'Calibri')).toBe('a');
  });

  it('passes through when the family is missing or unmapped', () => {
    expect(symbolFontToUnicode(ch(0xf0b7), null)).toBe(ch(0xf0b7));
    expect(symbolFontToUnicode(ch(0xf0b7), undefined)).toBe(ch(0xf0b7));
    // A code point with no entry in the selected table falls through unchanged
    // rather than guessing: 0xF0 is unassigned in Microsoft's Symbol.
    expect(symbolFontToUnicode(ch(0xf0f0), 'Symbol')).toBe(ch(0xf0f0));
  });

  it('matches exactly "wingdings" (case-insensitive, trimmed) for the Wingdings table', () => {
    expect(symbolFontToUnicode(ch(0xf0a7), 'WINGDINGS')).toBe('▪');
    expect(symbolFontToUnicode(ch(0xf0a7), '  Wingdings  ')).toBe('▪');
  });

  // Wingdings 2 / 3 and Webdings use DIFFERENT private encodings than Wingdings 1,
  // so they must NOT be looked up in the Wingdings 1 table (WINGDINGS_MAP) — doing
  // so produces a *wrong* glyph. They are unmapped (passthrough) so the real font
  // draws the intended glyph if installed; better the bare PUA point than tofu of a
  // wrong glyph.
  it('does NOT map Wingdings 2 / Wingdings 3 / Webdings through the Wingdings 1 table', () => {
    expect(symbolFontToUnicode(ch(0xa7), 'Wingdings 2')).toBe(ch(0xa7));
    expect(symbolFontToUnicode(ch(0xf0a7), 'Wingdings 2')).toBe(ch(0xf0a7));
    expect(symbolFontToUnicode(ch(0xa7), 'Wingdings 3')).toBe(ch(0xa7));
    expect(symbolFontToUnicode(ch(0xa7), 'Webdings')).toBe(ch(0xa7));
    // Wingdings 1 itself still resolves the same code point to the square bullet.
    expect(symbolFontToUnicode(ch(0xa7), 'Wingdings')).toBe('▪');
  });

  // Pencil/scissors/glasses are Wingdings glyphs only. In Symbol, 0x21 is "!",
  // 0x22 is "∀" and 0x24 is "∃" — so these dingbats must NOT leak across the
  // font gate.
  it('keeps Wingdings dingbats out of the Symbol table', () => {
    expect(symbolFontToUnicode(ch(0x21), 'Symbol')).toBe('!');
    expect(symbolFontToUnicode(ch(0x22), 'Symbol')).toBe('∀');
    expect(symbolFontToUnicode(ch(0x24), 'Symbol')).toBe('∃');
  });

  // Wingdings 0x21 "pencil" → U+270F PENCIL, 0x22 "scissors" → U+2702 BLACK
  // SCISSORS — both faithful BMP dingbats that render in ordinary fallback fonts.
  // (bare + PUA-shifted form.)
  it('maps the Wingdings pencil (0x21) and scissors (0x22) to BMP dingbats', () => {
    expect(symbolFontToUnicode(ch(0x21), 'Wingdings')).toBe('✏');
    expect(symbolFontToUnicode(ch(0xf021), 'Wingdings')).toBe('✏');
    expect(symbolFontToUnicode(ch(0x22), 'Wingdings')).toBe('✂');
    expect(symbolFontToUnicode(ch(0xf022), 'Wingdings')).toBe('✂');
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
  it('matches exactly "symbol" and any "wingdings" family', () => {
    expect(isSymbolFontFamily('Symbol')).toBe(true);
    expect(isSymbolFontFamily('SYMBOL')).toBe(true);
    expect(isSymbolFontFamily('Wingdings')).toBe(true);
    expect(isSymbolFontFamily('Wingdings 2')).toBe(true);
  });

  it('does NOT match Webdings (no core table) or ordinary faces', () => {
    // Webdings is intentionally excluded — core has no Webdings table, so the
    // gate must not claim it can normalize a Webdings run.
    expect(isSymbolFontFamily('Webdings')).toBe(false);
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
    // U+F0A7 (square) U+F0E0 (right arrow) → ▪ →
    const segs = symbolTextToUnicodeSegments(ch(0xf0a7) + ch(0xf0e0), 'Wingdings');
    expect(segs).toEqual([{ text: '▪→', mapped: true }]);
  });

  it('splits at mapped/unmapped boundaries so each run is single-font', () => {
    // 0x41 'A' / 0x42 'B' have no Wingdings entry (passthrough); 0xF0A7 maps to ▪.
    const segs = symbolTextToUnicodeSegments(ch(0x41) + ch(0xf0a7) + ch(0x42), 'Wingdings');
    expect(segs).toEqual([
      { text: ch(0x41), mapped: false },
      { text: '▪', mapped: true },
      { text: ch(0x42), mapped: false },
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

// Approach A: convert only where the host's Symbol face cannot draw the codes.
describe('a symbol-encoded host Symbol face', () => {
  afterEach(() => setSymbolFaceProbeForTesting(undefined));

  it('keeps Symbol codes (bare and F0-shifted) for the native face', () => {
    setSymbolFaceProbeForTesting(() => true);
    expect(symbolFontToUnicode('\uF061\uF0B7', 'Symbol')).toBe('\uF061\uF0B7');
    expect(symbolTextToUnicodeSegments('abcp', 'Symbol')).toEqual([{ text: 'abcp', mapped: false }]);
    // Wingdings keeps its own table; the probe concerns Symbol only.
    expect(symbolFontToUnicode(ch(0xf0a7), 'Wingdings')).toBe('▪');
  });

  it('converts when the host Symbol face is missing or Unicode-encoded', () => {
    setSymbolFaceProbeForTesting(() => false);
    expect(symbolFontToUnicode('\uF061b', 'Symbol')).toBe('αβ');
  });

  // The Canvas probe: a face that owns the codes changes their advances.
  it('detects a symbol-encoded face by its advances in a Canvas', () => {
    const fakeCanvas = (ownsCodes: boolean) => class {
      getContext() {
        let font = '';
        return {
          get font() { return font; }, set font(v: string) { font = v; },
          measureText: (t: string) => ({
            width: [...t].length * (ownsCodes && font.includes('"Symbol"') ? 55 : 60),
          }),
        };
      }
    };
    try {
      vi.stubGlobal('OffscreenCanvas', fakeCanvas(true));
      setSymbolFaceProbeForTesting(undefined);
      expect(symbolFontToUnicode('a', 'Symbol')).toBe('a');
      vi.stubGlobal('OffscreenCanvas', fakeCanvas(false));
      setSymbolFaceProbeForTesting(undefined);
      expect(symbolFontToUnicode('a', 'Symbol')).toBe('α');
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it('probes the host once per realm', () => {
    let calls = 0;
    setSymbolFaceProbeForTesting(() => { calls++; return true; });
    symbolFontToUnicode('a', 'Symbol');
    symbolFontToUnicode('b', 'Symbol');
    expect(calls).toBe(1);
  });
});

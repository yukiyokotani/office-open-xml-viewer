/**
 * Symbol / Wingdings / Webdings font-encoding → Unicode normalization.
 *
 * "Symbol", "Wingdings", "Wingdings 2", "Wingdings 3" and "Webdings" are
 * symbol-encoded fonts: each carries an OpenType symbol cmap (platform 3,
 * encoding 0) whose codes 0x20..0xFF are the font's own glyphs, and the
 * encodings disagree (0xB7 is BULLET in Symbol but a clock face in
 * Wingdings). OOXML stores run and marker text as those codes, either
 * F0-shifted into the Private Use Area (U+F020..U+F0FF, what Insert > Symbol
 * writes) or bare (an ASCII "a" in a Symbol run is alpha).
 * https://learn.microsoft.com/en-us/typography/opentype/spec/cmap
 *
 * ECMA-376 does not define these repertoires: §17.3.2.26 (`w:rFonts`),
 * §17.3.3.30 (`w:sym`), §17.9.x (`w:lvlText`) and §21.1.2.3.10 (`a:sym`) name
 * a font and a code; the glyph belongs to the font.
 *
 * Whether to convert is decided per CODE, by asking the context that paints
 * whether the font the kept text paints with draws it (symbol-face.ts):
 *   - a code that font draws itself (a symbol-encoded face that maps it) is
 *     kept, so an installed Windows/macOS Wingdings or Windows Symbol keeps
 *     its own glyph, advance and colour;
 *   - a code it does not draw (the font is missing, or the face is
 *     Unicode-encoded, as Apple's Symbol is) converts to the Unicode
 *     character of the font's glyph, drawn in a Unicode fallback face;
 *   - when the font could not be judged, the conversion is today's: only the
 *     long-standing bullet/arrow subset ({@link LEGACY_LOOKUPS}).
 * The PUA and bare forms are judged separately (a Unicode face never draws a
 * bare code as the symbol glyph; a face may draw only one form).
 *
 * A converted pictograph is written with U+FE0E VARIATION SELECTOR-15 (text
 * presentation) and is converted only when the caller proves that this realm
 * paints it monochrome (`monochrome`); otherwise the authored code is kept.
 * Converted symbols must take the text colour, never a colour emoji face.
 *
 * SOURCES (the tables record their provenance and every deviation):
 *   - Symbol: Adobe "Symbol Encoding to Unicode" — symbol-encoding.ts.
 *   - Wingdings, Wingdings 2, Wingdings 3, Webdings: WG2 N4384 = L2/12-368
 *     section 7 — dingbat-encoding.ts.
 */

import {
  ADOBE_SYMBOL_UNICODE,
  SYMBOL_CORPORATE_USE_UNICODE,
} from './symbol-encoding.js';
import {
  WEBDINGS_ENCODING,
  WINGDINGS_2_ENCODING,
  WINGDINGS_3_ENCODING,
  WINGDINGS_ENCODING,
} from './dingbat-encoding.js';

/** The symbol-encoded families this module has tables for (exact, trimmed,
 * case-insensitive names). */
export type SymbolFontKey = 'symbol' | 'wingdings' | 'wingdings 2' | 'wingdings 3' | 'webdings';

const SYMBOL_FONT_KEYS: ReadonlySet<string> = new Set<SymbolFontKey>([
  'symbol', 'wingdings', 'wingdings 2', 'wingdings 3', 'webdings',
]);

/** The table key of `fontFamily`, or undefined. An exact name match: a Latin
 * face such as "SymbolMT" or "Wingdings Pro" is not one of these encodings. */
export function symbolFontKey(fontFamily: string | null | undefined): SymbolFontKey | undefined {
  const key = fontFamily?.trim().toLowerCase();
  return key && SYMBOL_FONT_KEYS.has(key) ? key as SymbolFontKey : undefined;
}

/** True when `fontFamily` is one of the symbol-encoded families above. */
export function isSymbolFontFamily(fontFamily: string | null | undefined): boolean {
  return symbolFontKey(fontFamily) !== undefined;
}

/** Symbol encoding → Unicode: every assigned code 0x20..0xFE with a standard
 * Unicode target (symbol-encoding.ts). */
export const SYMBOL_MAP: Readonly<Record<number, string>> = Object.freeze(Object.fromEntries(
  Object.entries({ ...ADOBE_SYMBOL_UNICODE, ...SYMBOL_CORPORATE_USE_UNICODE })
    .map(([code, unicode]) => [Number(code), String.fromCodePoint(unicode)]),
));

/** Decode a dingbat-encoding.ts table (codes 0x20..0xFF as base-36 code
 * point deltas, an empty token for no target) into a code → character table. */
function dingbatTable(encoding: string): Record<number, string> {
  const out: Record<number, string> = {};
  let code = 0x20;
  let cp = 0;
  for (const token of encoding.split(',')) {
    if (token) {
      cp += Number.parseInt(token, 36);
      out[code] = String.fromCodePoint(cp);
    }
    code++;
  }
  return out;
}

/** Microsoft Wingdings → Unicode (221 of 224 codes; dingbat-encoding.ts). */
export const WINGDINGS_MAP: Readonly<Record<number, string>> = Object.freeze(dingbatTable(WINGDINGS_ENCODING));

const FULL_LOOKUPS: ReadonlyMap<SymbolFontKey, Readonly<Record<number, string>>> = new Map([
  ['symbol', SYMBOL_MAP],
  ['wingdings', WINGDINGS_MAP],
  ['wingdings 2', dingbatTable(WINGDINGS_2_ENCODING)],
  ['wingdings 3', dingbatTable(WINGDINGS_3_ENCODING)],
  ['webdings', dingbatTable(WEBDINGS_ENCODING)],
]);

/**
 * The conversion of 0.89, kept for faces that could not be judged (no
 * measuring Canvas): the bullets, arrows and operators of the Symbol table,
 * and Wingdings bullets with their BMP stand-ins. Wingdings 2/3 and Webdings
 * had none.
 */
const LEGACY_LOOKUPS: ReadonlyMap<SymbolFontKey, Readonly<Record<number, string>>> = new Map<SymbolFontKey, Readonly<Record<number, string>>>([
  // The 0.89 Symbol subset: the same targets as the full table, by code.
  ['symbol', Object.fromEntries([0xa3, 0xa7, 0xa8, 0xa9, 0xaa, 0xab, 0xac, 0xad, 0xae, 0xaf, 0xb0, 0xb1, 0xb3, 0xb4, 0xb7, 0xb8, 0xb9]
    .map((code) => [code, SYMBOL_MAP[code]!]))],
  ['wingdings', {
    0x21: '✏', 0x22: '✂', 0x24: '\u{1F453}', 0x4a: '☺', 0x4b: '\u{1F610}', 0x4c: '☹',
    0x6c: '●', 0x6e: '■', 0x74: '◆', 0x76: '❖', 0x77: '◆', 0xa7: '▪',
    0xdf: '←', 0xe0: '→', 0xe1: '↑', 0xe2: '↓', 0xe3: '↖', 0xe4: '↗', 0xe5: '↙', 0xe6: '↘',
    0xfb: '✗', 0xfc: '✓', 0xfd: '☒', 0xfe: '☑',
  }],
]);

/** Targets with the Unicode Emoji property (UTS #51; ASCII keycap bases
 * excluded, which no table converts to) may be drawn from a colour emoji face
 * even with VS15 when no text face has them (CSS Fonts 4 §5.4 leaves the face
 * to the engine), so they are written with VS15 and converted only when
 * proven monochrome. Other characters, Extended_Pictographic reservations
 * included, have no emoji presentation. */
const EMOJI_CAPABLE = /(?![#*0-9])\p{Emoji}/u;
const EMOJI_PRESENTATION = /\p{Emoji_Presentation}/u;
const TEXT_PRESENTATION = String.fromCharCode(0xfe0e);

export interface SymbolTextOptions {
  /** Whether the font that paints the kept text draws this authored code (one
   * character: a bare code 0x20..0xFF or U+F020..U+F0FF) as its own symbol
   * glyph (symbol-face.ts). Undefined, or absent: it cannot be judged, and
   * only the 0.89 subset converts. Asked only for codes with a conversion. */
  readonly native?: (code: string) => boolean | undefined;
  /** Whether this realm paints `text` (an emoji-capable character followed
   * by VS15) in a monochrome face in the fallback route. Absent (no probe):
   * only the 0.89 subset's text-default symbols convert. */
  readonly monochrome?: (text: string) => boolean | undefined;
}

/** A maximal run of characters sharing a single rendering disposition. */
export interface SymbolTextSegment {
  /** The (possibly converted) text for this run. */
  readonly text: string;
  /**
   * True when the text was converted to Unicode. Such a segment must be drawn
   * in a Unicode fallback face, never the symbol family: a symbol cmap would
   * re-read a converted character such as U+00A9 as its own code 0xA9 (a
   * heart in Symbol). An unconverted segment keeps the symbol family.
   */
  readonly mapped: boolean;
}

/** Bare codes with the Bidi_Mirrored property: in right-to-left text a
 * shaper replaces them with their mirror partner, which in a symbol font is
 * an unrelated glyph ("\u00AB" is Symbol's left-right arrow, its partner
 * "\u00BB" the almost-equal sign). */
const BIDI_MIRRORED = /[()<>[\]{}\u00AB\u00BB]/u;

/** The Unicode text of one authored character (`mapped`), the form it is
 * kept in, or undefined: kept as authored. */
function convertCodePoint(
  ch: string,
  key: SymbolFontKey,
  options: SymbolTextOptions,
): Readonly<{ text: string; mapped: boolean }> | undefined {
  const converted = (text: string | undefined) => (text === undefined ? undefined : { text, mapped: true });
  const cp = ch.codePointAt(0)!;
  const pua = cp >= 0xf020 && cp <= 0xf0ff;
  const code = pua ? cp - 0xf000 : cp;
  if (!pua && (cp < 0x20 || cp > 0xff)) return undefined;
  const full = FULL_LOOKUPS.get(key)?.[code];
  const legacy = LEGACY_LOOKUPS.get(key)?.[code];
  if ((full ?? ch) === ch && (legacy ?? ch) === ch) return undefined;
  const native = options.native?.(ch);
  // A native bare code that bidi mirroring would replace is kept as its
  // F0-shifted form: the very glyph (that identity is what makes a bare code
  // native, symbol-face.ts), which no direction mirrors.
  if (native) return !pua && BIDI_MIRRORED.test(ch) ? { text: String.fromCharCode(0xf000 + code), mapped: false } : undefined;
  const target = native === undefined ? legacy : full;
  if (target === undefined || target === ch) return undefined;
  if (EMOJI_CAPABLE.test(target)) {
    const text = target + TEXT_PRESENTATION;
    const proof = options.monochrome?.(text);
    if (proof !== undefined) return converted(proof ? text : undefined);
    // Inconclusive (no readable canvas in this realm): the 0.89 subset keeps
    // its text-default symbols (▪, ✓, ☺), now with VS15; a character whose
    // default presentation is emoji is kept as authored.
    return converted(native === undefined && !EMOJI_PRESENTATION.test(target) ? text : undefined);
  }
  return converted(target);
}

/**
 * Convert every code point of `text` that the painted font does not draw
 * (see the module notes) and split the result into maximal same-disposition
 * runs, so the caller can switch the draw font at each boundary. A family
 * without a table yields one unmapped segment. An unmapped segment is the
 * authored text, except that a native bare code which bidi mirroring would
 * replace is written F0-shifted (same glyph, same font).
 */
export function symbolTextToUnicodeSegments(
  text: string,
  fontFamily: string | null | undefined,
  options: SymbolTextOptions = {},
): SymbolTextSegment[] {
  const key = symbolFontKey(fontFamily);
  if (!key || text.length === 0) return [{ text, mapped: false }];
  const out: SymbolTextSegment[] = [];
  let buf = '';
  let bufMapped: boolean | null = null;
  for (const ch of text) {
    const converted = convertCodePoint(ch, key, options);
    const mapped = converted?.mapped === true;
    const piece = converted?.text ?? ch;
    if (bufMapped === null || mapped === bufMapped) {
      bufMapped = mapped;
      buf += piece;
    } else {
      out.push({ text: buf, mapped: bufMapped });
      bufMapped = mapped;
      buf = piece;
    }
  }
  if (buf.length > 0) out.push({ text: buf, mapped: bufMapped ?? false });
  return out;
}

/** {@link symbolTextToUnicodeSegments}, joined. */
export function symbolFontToUnicode(
  char: string,
  fontFamily: string | null | undefined,
  options: SymbolTextOptions = {},
): string {
  return symbolTextToUnicodeSegments(char, fontFamily, options).map((part) => part.text).join('');
}

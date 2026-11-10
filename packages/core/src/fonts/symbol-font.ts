/**
 * Symbol / Wingdings / Webdings font-encoding → Unicode normalization.
 *
 * The classic "Symbol", "Wingdings", "Wingdings 2", "Wingdings 3" and
 * "Webdings" families do NOT use a Unicode cmap: each ships a private,
 * font-specific encoding where ordinary code points map to its own glyphs, and
 * the encodings DISAGREE (0xB7 is BULLET "•" in Symbol but a clock face in
 * Wingdings; 0xA7 is "♣" in Symbol but "▪" in Wingdings). OOXML stores run and
 * marker text as the *font's own* code points — either bare (an ASCII "a" in a
 * Symbol run is α) or, as Word and PowerPoint write inserted symbols, in the
 * Private Use Area U+F020–U+F0FF (a Symbol bullet is U+F0B7). So each family
 * has its own table, chosen from the requested font family.
 *
 * ECMA-376 / ISO-29500 does NOT define these fonts' glyph repertoires:
 *   - §17.3.2.26 (`w:rFonts`), §17.3.3.30 (`w:sym`) and §17.9.x (`w:lvlText`)
 *     for docx tell us WHICH font and WHICH code point, but the
 *     code-point→glyph mapping belongs to the font, not the spec.
 *   - §21.1.2.3.10 (`a:sym`) for pptx is likewise just a font + char reference.
 *
 * Without the font (always on non-Windows hosts) a font-coded PUA point has no
 * glyph and draws as tofu, and a bare code draws as the wrong Latin letter. We
 * normalize both forms to the Unicode character the font's glyph is, so the
 * intended glyph renders in any fallback face. Keyed solely on the requested
 * family and each font's published repertoire — not a per-sample heuristic.
 *
 * SOURCES (the tables record their provenance and every deviation):
 *   - Symbol:    Adobe "Symbol Encoding to Unicode" (unicode.org/Public/
 *                MAPPINGS/VENDORS/ADOBE/symbol.txt) — symbol-encoding.ts.
 *   - Wingdings, Wingdings 2, Wingdings 3, Webdings: the source reference of
 *                their encoding in Unicode 6.0/7.0 (WG2 N4384 = L2/12-368,
 *                section 7) — dingbat-encoding.ts.
 *
 * A code the source gives no Unicode character for passes through unchanged,
 * so an installed font still draws its native glyph; we never substitute a
 * different glyph.
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

/**
 * Symbol encoding → Unicode: every assigned code 0x20..0xFE that has a
 * standard Unicode target (symbol-encoding.ts records the source table, its
 * terms, and each deviation). That covers the Greek capitals and lowercase at
 * 0x41..0x5A / 0x61..0x7A, the operators, relations, arrows and set symbols in
 * 0xA0..0xFE, and the bracket pieces. Symbol text arrives either as these bare
 * codes (an ASCII "abcp" in a Symbol run is αβχπ) or F0-shifted (U+F061…), and
 * both resolve. 0xF0 and 0x7F..0x9F are unassigned and have no target.
 */
export const SYMBOL_MAP: Record<number, string> = Object.fromEntries(
  Object.entries({ ...ADOBE_SYMBOL_UNICODE, ...SYMBOL_CORPORATE_USE_UNICODE })
    .map(([code, unicode]) => [Number(code), String.fromCodePoint(unicode)]),
);

/** Expand a dingbat-encoding.ts string (codes 0x20..0xFF, one code point
 *  each, '-' for no target) into a code → character table. */
function dingbatTable(encoding: string): Record<number, string> {
  const out: Record<number, string> = {};
  let code = 0x20;
  for (const ch of encoding) {
    if (ch !== '-') out[code] = ch;
    code++;
  }
  return out;
}

/** Microsoft Wingdings → Unicode (221 of 224 codes; dingbat-encoding.ts). */
export const WINGDINGS_MAP: Record<number, string> = dingbatTable(WINGDINGS_ENCODING);

/** Add the PUA-shifted (0xF000 + code) variant for every entry in `base`. */
function withPua(base: Record<number, string>): Record<number, string> {
  const out: Record<number, string> = {};
  for (const key of Object.keys(base)) {
    const code = Number(key);
    out[code] = base[code];
    out[0xf000 + code] = base[code];
  }
  return out;
}

const LOOKUPS: ReadonlyMap<string, Record<number, string>> = new Map([
  ['symbol', withPua(SYMBOL_MAP)],
  ['wingdings', withPua(WINGDINGS_MAP)],
  ['wingdings 2', withPua(dingbatTable(WINGDINGS_2_ENCODING))],
  ['wingdings 3', withPua(dingbatTable(WINGDINGS_3_ENCODING))],
  ['webdings', withPua(dingbatTable(WEBDINGS_ENCODING))],
]);

/**
 * Normalize Symbol/Wingdings/Webdings code points to their Unicode
 * equivalents, one code point at a time (a single character in the common
 * case).
 *
 * The table is chosen by `fontFamily` via a TRIMMED EXACT match (case-
 * insensitive): "symbol", "wingdings", "wingdings 2", "wingdings 3" or
 * "webdings". Each family has its own encoding, so a family is never routed
 * through another's table. Both the bare code point (0xA7, 0xB7, …) and its
 * PUA-shifted form (0xF0A7, 0xF0B7, …) resolve.
 *
 * Returns the input unchanged when `fontFamily` is null/undefined or not one of
 * those families, and keeps each code point the selected table has no target
 * for (passthrough — the authored face, when installed, still draws it).
 *
 * @param char       the marker/run text as stored in the OOXML (often one
 *                   PUA code point such as U+F0B7)
 * @param fontFamily the requested font family (ascii / hAnsi axis)
 */
export function symbolFontToUnicode(
  char: string,
  fontFamily: string | null | undefined,
): string {
  if (!fontFamily) return char;
  const table = LOOKUPS.get(fontFamily.trim().toLowerCase());
  if (!table) return char;
  // Each code point is converted independently, so a multi-character marker
  // (a `w:lvlText` such as "1." in a Symbol run) keeps every character.
  let out = '';
  for (const ch of char) out += table[ch.codePointAt(0)!] ?? ch;
  return out;
}

/**
 * True when `fontFamily` is one of the font-specific (non-Unicode-cmap)
 * encodings that {@link symbolFontToUnicode} knows how to normalize: exactly
 * "Symbol" or "Webdings", or any "Wingdings" family (§17.3.2.26 /
 * §21.1.2.3.10).
 *
 * The shared "is this a symbol font?" gate, used by both docx and pptx so they
 * don't hand-roll divergent regexes. It does not match a Latin face like
 * "SymbolMT" (the "symbol" check is exact, not a substring).
 */
export function isSymbolFontFamily(
  fontFamily: string | null | undefined,
): boolean {
  if (!fontFamily) return false;
  const lower = fontFamily.toLowerCase();
  return lower === 'symbol' || lower === 'webdings' || lower.includes('wingdings');
}

/** A maximal run of characters sharing a single rendering disposition. */
export interface SymbolTextSegment {
  /** The (possibly normalized) text for this run. */
  text: string;
  /**
   * True when at least one character in `text` was mapped from the symbol
   * font's private encoding to a Unicode equivalent. Such a segment must be
   * drawn in a generic fallback face (the Symbol/Wingdings font, if present,
   * would re-interpret the Unicode code point as the WRONG glyph); an unmapped
   * segment keeps the requested symbol font so an installed Symbol/Wingdings
   * still draws its native glyph.
   */
  mapped: boolean;
}

/**
 * String-level companion to {@link symbolFontToUnicode}: normalize every
 * character of `text` through the font's private encoding and split the result
 * into maximal same-disposition runs so the caller can switch the draw font at
 * each mapped/unmapped boundary.
 *
 * When `fontFamily` is not a symbol font (per {@link isSymbolFontFamily}), the
 * whole string is returned as a single `{ text, mapped: false }` segment
 * unchanged — the caller then needs no special handling. Iterates by code point
 * so astral targets (e.g. Wingdings 0x24 → U+1F453) survive intact.
 */
export function symbolTextToUnicodeSegments(
  text: string,
  fontFamily: string | null | undefined,
): SymbolTextSegment[] {
  if (!isSymbolFontFamily(fontFamily) || text.length === 0) {
    return [{ text, mapped: false }];
  }
  const out: SymbolTextSegment[] = [];
  let buf = '';
  let bufMapped: boolean | null = null;
  for (const ch of text) {
    const normalized = symbolFontToUnicode(ch, fontFamily);
    const mapped = normalized !== ch;
    if (bufMapped === null || mapped === bufMapped) {
      bufMapped = mapped;
      buf += normalized;
    } else {
      out.push({ text: buf, mapped: bufMapped });
      bufMapped = mapped;
      buf = normalized;
    }
  }
  if (buf.length > 0) out.push({ text: buf, mapped: bufMapped ?? false });
  return out;
}

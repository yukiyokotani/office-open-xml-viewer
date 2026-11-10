/**
 * Symbol / Wingdings font-encoding → Unicode normalization.
 *
 * The classic "Symbol" and "Wingdings" families do NOT use a Unicode cmap: they
 * ship a private, font-specific encoding where ordinary code points map to
 * pictographic glyphs. Crucially the two fonts use DIFFERENT encodings, so the
 * same code point is a different glyph in each (e.g. 0xB7 is BULLET "•" in
 * Symbol but a clock face in Wingdings; 0xA7 is BLACK CLUB SUIT "♣" in Symbol
 * but a small filled square "▪" in Wingdings). OOXML stores marker / run text as
 * the *font's own* code points — and Word commonly emits them in the Private Use
 * Area (U+F020–U+F0FF), so a Symbol bullet is serialized as U+F0B7 and a
 * Wingdings square as U+F0A7.
 *
 * Because the two encodings disagree on the same code points, a single shared
 * table cannot be correct for both fonts. We therefore keep TWO font-specific
 * tables ({@link SYMBOL_MAP}, {@link WINGDINGS_MAP}) and pick the right one from
 * the requested font family.
 *
 * ECMA-376 / ISO-29500 does NOT define these fonts' glyph repertoires:
 *   - §17.9.x  (numbering level text, `w:lvlText`) and §17.3.2.26 (`w:rPr`
 *     `w:rFonts`) for docx tell us WHICH font and WHICH code point, but the
 *     code-point→glyph mapping belongs to the font implementation, not the spec.
 *   - §21.1.2.3.10 (`a:sym`) for pptx is likewise just a font + char reference.
 *
 * When such a marker is drawn in a generic fallback face (the Symbol/Wingdings
 * font is frequently unavailable, especially on non-Windows hosts), the PUA code
 * point has no glyph and renders as tofu (□). We normalize the well-known
 * Symbol/Wingdings code points — both their bare form (0xA7, 0xB7, …) and their
 * PUA-shifted form (0xF0A7, 0xF0B7, …) — to the equivalent Unicode characters
 * so the intended glyph (•, ▪, →, …) renders in any fallback font.
 *
 * This is a documented font-encoding → Unicode conversion, NOT a per-sample
 * heuristic: it is keyed solely on the requested font family ("symbol" /
 * "wingdings") and each font's published code-point repertoire.
 *
 * SOURCES (authoritative, cited per entry below):
 *   - Symbol:   Adobe "Symbol Encoding to Unicode", unicode.org/Public/MAPPINGS/
 *               VENDORS/ADOBE/symbol.txt (field 1 = Unicode, field 2 = Symbol
 *               code point). This is the Adobe Symbol PostScript encoding;
 *               the full table lives in symbol-encoding.ts.
 *   - Wingdings: the Microsoft Wingdings cmap as standardized in the Unicode
 *               proposals L2/11-052 and L2/11-344 (the glyph repertoire that
 *               added the Wingdings pictographs to Unicode 7.0). The byte→Unicode
 *               targets below were cross-checked against the PostScript glyph
 *               names in the shipping Wingdings.ttf (e.g. 0x21 "pencil",
 *               0xA7 "square4", 0x24 "readingglasses").
 *
 * POLICY (carried over from PR #519): a code point is mapped ONLY when its target
 * is confirmed from the source above. Glyphs whose only faithful Unicode target
 * is an astral pictograph that renders as tofu in ordinary fallback fonts (e.g.
 * the Wingdings clock faces U+1F550…, the office/hand dingbats U+1F5xx) are left
 * as passthrough: emitting the bare PUA code point is no worse, and we never
 * substitute a *wrong* but available glyph. Better tofu than a wrong glyph.
 */

import {
  ADOBE_SYMBOL_UNICODE,
  SYMBOL_CORPORATE_USE_UNICODE,
} from './symbol-encoding.js';

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

/**
 * Microsoft Wingdings cmap → Unicode (subset).
 *
 * Each byte's glyph is identified from the shipping Wingdings.ttf PostScript
 * glyph names (e.g. 0xA7 "square4", 0x21 "pencil"); the Unicode 7.0 proposals
 * L2/11-052 / L2/11-344 catalogue those glyphs but at DEDICATED ASTRAL code
 * points (barb arrows U+1F868…, bold check/X U+1F5F6…) that tofu in ordinary
 * fonts. So the targets below are deliberate faithful BMP substitutions (→, ✓,
 * ▪, …), NOT those canonical astral code points. Entries whose only faithful
 * target is an astral pictograph (the office/hand/clock dingbats U+1F5xx /
 * U+1F55x) are OMITTED — passthrough, since the bare PUA point is no worse.
 */
export const WINGDINGS_MAP: Record<number, string> = {
  // 0x21 "pencil" → U+270F PENCIL; 0x22 "scissors" → U+2702 BLACK SCISSORS. Both
  // are faithful BMP dingbats that render in any ordinary fallback font (unlike
  // the Unicode-7.0 canonical unifications U+1F589 / U+2702, where the pencil is
  // an astral PUA-adjacent glyph). The Dingbats ✏/✂ are the semantically exact
  // glyphs, so this is a faithful substitution, not a guess.
  // 0x23 is the SECOND Wingdings scissors variant ("scissorscutting"); its only
  // distinct faithful target (U+2701/U+2703) is an obscure half-blade scissors
  // that tofus in most fonts and is easily confused with 0x22, so it is left as
  // passthrough rather than mapped to a near-duplicate ✂.
  0x21: '✏', 0x22: '✂',
  // 0x24 "readingglasses" → U+1F453 EYEGLASSES. Astral with no BMP equivalent;
  // unlike the omitted clock/hand dingbats we DO map it because U+1F453 is the
  // semantically exact glyph — it renders in an emoji-capable fallback and is
  // never a *wrong* glyph (passthrough would always tofu the bare PUA point).
  0x24: '\u{1F453}',
  // 0x4A "smileface" → U+263A; 0x4B "neutralface" → U+1F610 (astral, exact —
  // same rationale as 0x24); 0x4C "frownface" → U+2639.
  0x4a: '☺', 0x4b: '\u{1F610}', 0x4c: '☹',
  // 0x76 "xrhombus" → U+2756 BLACK DIAMOND MINUS WHITE X (restored from PR #519)
  0x76: '❖',
  // 0xA7 "square4" → U+25AA BLACK SMALL SQUARE (the Wingdings list bullet)
  0xa7: '▪',
  // 0x6C "circle6" → U+25CF BLACK CIRCLE; 0x6E "square6" → U+25A0 BLACK SQUARE.
  // 0x74 "lozenge6" and 0x77 "rhombus4" are both SOLID/black diamond-family
  // glyphs (a tall lozenge and a wider rhombus); BMP doesn't distinguish the
  // proportion, so both approximate to U+25C6 BLACK DIAMOND. (0x77 was wrongly
  // U+25C7 WHITE DIAMOND — a fill error — before this fix.)
  0x6c: '●', 0x6e: '■', 0x74: '◆', 0x77: '◆',
  // 0xFB "xmarkbld" → U+2717 BALLOT X; 0xFC "checkbld" → U+2713 CHECK MARK;
  // 0xFD "boxxmarkbld" → U+2612 BALLOT BOX WITH X;
  // 0xFE "boxcheckbld" → U+2611 BALLOT BOX WITH CHECK (restored from PR #519)
  0xfb: '✗', 0xfc: '✓', 0xfd: '☒', 0xfe: '☑',
  // Wingdings barb2 arrow block — 0xDF "barb2left" … 0xE6 "barb2se". Mapped to
  // Unicode arrows so they render in any font when Wingdings is unavailable.
  0xdf: '←', 0xe0: '→', 0xe1: '↑', 0xe2: '↓',
  0xe3: '↖', 0xe4: '↗', 0xe5: '↙', 0xe6: '↘',
};

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

const SYMBOL_LOOKUP = withPua(SYMBOL_MAP);
const WINGDINGS_LOOKUP = withPua(WINGDINGS_MAP);

/**
 * Whether the host's "Symbol" face draws Symbol's own font codes. Microsoft's
 * Symbol (always on Windows) carries a symbol cmap (OpenType platform 3
 * encoding 0) at U+F020..U+F0FF, and text shaping also resolves bare codes
 * 0x20..0xFF through it, so the authored face draws the intended glyph
 * natively and must keep drawing it. Apple's Symbol (macOS) is Unicode-encoded
 * and has no glyph for those codes, and most other hosts have no Symbol at
 * all; only there does the text need the Unicode conversion.
 * https://learn.microsoft.com/en-us/typography/opentype/spec/cmap
 *
 * Capability probe: each probe code is measured in the authored face (backed
 * by an absent family) and in the absent family alone. Symbol's glyphs for
 * these codes (α β γ π ∑ ∂) have different advances. A face that lacks them
 * draws one missing-glyph box (or the same fallback) for every code, so its
 * advances are all equal; and a fallback face that happens to own them gives
 * the absent family the same advances. So the codes count as drawn natively
 * only when their advances in the authored face both vary and differ from the
 * absent family's. Without a Canvas (Node without a DOM, unit tests) the probe
 * reports false and the text is converted. Checked once per realm: system
 * faces do not change during a session.
 */
type SymbolFaceProbe = () => boolean;

const SYMBOL_PROBE_CODES = ['\uF061', '\uF062', '\uF067', '\uF070', '\uF0E5', '\uF0B6'];
const ABSENT_PROBE_FAMILY = 'ooxml-symbol-probe-absent-face';

function canvasSymbolFaceProbe(): boolean {
  let ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D | null = null;
  try {
    if (typeof OffscreenCanvas !== 'undefined') ctx = new OffscreenCanvas(1, 1).getContext('2d');
    else if (typeof document !== 'undefined') ctx = document.createElement('canvas').getContext('2d');
  } catch {
    ctx = null;
  }
  if (!ctx || typeof ctx.measureText !== 'function') return false;
  const advances = (font: string): number[] => {
    ctx!.font = font;
    return SYMBOL_PROBE_CODES.map((code) => ctx!.measureText(code).width);
  };
  const authored = advances(`100px "Symbol", "${ABSENT_PROBE_FAMILY}"`);
  const absent = advances(`100px "${ABSENT_PROBE_FAMILY}"`);
  if (!authored.every(Number.isFinite)) return false;
  const varies = authored.some((width) => Math.abs(width - authored[0]) > 0.5);
  const ownAdvances = authored.some((width, i) => Math.abs(width - absent[i]) > 0.5);
  return varies && ownAdvances;
}

let symbolFaceProbe: SymbolFaceProbe = canvasSymbolFaceProbe;
let symbolFaceDrawsCodes: boolean | undefined;

function hostSymbolFaceDrawsCodes(): boolean {
  symbolFaceDrawsCodes ??= symbolFaceProbe();
  return symbolFaceDrawsCodes;
}

/** Test seam: replace the host Symbol-face probe (`undefined` restores the
 *  Canvas probe) and forget the cached answer. Not part of the public API. */
export function setSymbolFaceProbeForTesting(probe: SymbolFaceProbe | undefined): void {
  symbolFaceProbe = probe ?? canvasSymbolFaceProbe;
  symbolFaceDrawsCodes = undefined;
}

/**
 * Normalize Symbol/Wingdings code points to their Unicode equivalents, one
 * code point at a time (a single character in the common case).
 *
 * The lookup table is chosen by `fontFamily` via a TRIMMED EXACT match (case-
 * insensitive): exactly "symbol" selects the Adobe Symbol encoding
 * ({@link SYMBOL_MAP}); exactly "wingdings" selects the Wingdings 1 cmap
 * ({@link WINGDINGS_MAP}). Both the bare code point (0xA7, 0xB7, …) and its
 * PUA-shifted form (0xF0A7, 0xF0B7, …) resolve.
 *
 * "Wingdings 2", "Wingdings 3" and "Webdings" use DIFFERENT private encodings —
 * the same code point is a different glyph than in Wingdings 1 — so they are
 * deliberately NOT matched and pass through unchanged (we have no table for
 * them). Routing them through the Wingdings 1 table would emit a *wrong* glyph;
 * passthrough lets the real font draw the intended glyph when installed, and a
 * bare PUA point is no worse than a wrong substitution. Add a dedicated table if
 * those repertoires are ever needed.
 *
 * Returns `char` unchanged when `fontFamily` is null/undefined, is not one of the
 * two mapped families, or has no mapping in the selected table (passthrough —
 * better the bare code point than a wrong glyph). Symbol text is also returned
 * unchanged when the host's Symbol face is symbol-encoded and so draws the
 * codes itself (see the probe above).
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
  const lower = fontFamily.trim().toLowerCase();
  const table =
    lower === 'wingdings'
      ? WINGDINGS_LOOKUP
      : lower === 'symbol'
        ? (hostSymbolFaceDrawsCodes() ? null : SYMBOL_LOOKUP)
        : null;
  if (!table) return char;
  // Each code point is converted independently, so a multi-character marker
  // (a `w:lvlText` such as "1." in a Symbol run) keeps every character.
  let out = '';
  for (const ch of char) out += table[ch.codePointAt(0)!] ?? ch;
  return out;
}

/**
 * True when `fontFamily` is one of the font-specific (non-Unicode-cmap)
 * encodings that {@link symbolFontToUnicode} knows how to normalize — currently
 * exactly "Symbol" and any "Wingdings" family (§17.3.2.26 / §21.1.2.3.10).
 *
 * The shared "is this a symbol font?" gate, used by both docx and pptx so they
 * don't hand-roll divergent regexes. It deliberately does NOT match "Webdings"
 * (core has no Webdings table yet — a Webdings PUA point would pass the gate
 * only to return unchanged) nor a Latin face like "SymbolMT" (the "symbol"
 * check is exact, not a substring), so a caller wanting either must add it
 * explicitly and accept passthrough.
 */
export function isSymbolFontFamily(
  fontFamily: string | null | undefined,
): boolean {
  if (!fontFamily) return false;
  const lower = fontFamily.toLowerCase();
  return lower === 'symbol' || lower.includes('wingdings');
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

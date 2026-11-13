/**
 * Symbol-encoded text in XLSX paint (core symbol-font.ts, symbol-face.ts;
 * issue #2177 for DOCX/PPTX). A face such as Symbol, Wingdings 1-3 or
 * Webdings draws its own glyphs at the authored codes: SpreadsheetML names it
 * as a cell or run font (ECMA-376 §18.8.22 `font/name`, §18.4.7 `rFont`),
 * DrawingML as a shape run's Latin or symbol face (§21.1.2.3.7 `a:latin`,
 * §21.1.2.3.10 `a:sym`). Where the face that paints is not that symbol face,
 * the authored Latin letters would paint as letters, so the codes the core
 * oracle proves the painting font does not draw convert to their Unicode
 * equivalents, with the same tables and verdicts as DOCX and PPTX: an
 * F0-shifted code is kept when the font draws it, a bare code when the font
 * draws it as the very glyph of its F0-shifted form, any other code converts
 * (an emoji-capable target only when proven to take the text colour). A code
 * the realm cannot judge keeps the 0.89 behaviour: a bare code paints as
 * authored, an F0-shifted one converts through the 0.89 subset.
 *
 * Library policy (not Excel behaviour): one text paints in one face. A cell
 * value, a rich run, a shape run piece and a slicer label are each measured
 * and painted with a single Canvas font, which the plain cell paths (overflow,
 * fill, distributed, rotated, stacked, shrink-to-fit, wrap) rely on. A text
 * whose codes are all kept, or all converted, has one face by itself. A text
 * that mixes codes the symbol face draws with codes it lacks (a partial face)
 * converts every convertible code and paints in the Unicode fallback face:
 * converted text must never paint in the symbol family, whose cmap would
 * re-read a target such as U+00A9 as its own symbol code.
 */
import { isSymbolFontFamily, symbolTextToUnicodeSegments, type SymbolFaceOracle } from '@silurus/ooxml-core';

/** How one authored text paints. */
export interface SymbolPaint {
  /** True: the text paints in the Unicode fallback face (some code
   * converted); false: in the authored face. */
  readonly unicode: boolean;
  /** The painted form of the text this was made for, or of any of its
   * substrings (wrap pieces, clipped prefixes). */
  paint(text: string): string;
}

const identity = (text: string) => text;
/** Text without symbol-encoded codes: painted as authored. */
export const AUTHORED_PAINT: SymbolPaint = Object.freeze({ unicode: false, paint: identity });

const TEXT_PRESENTATION = /︎/gu;

/**
 * How `text`, authored in `family`, paints when kept in `keptFont` (the
 * Canvas font of the authored face, `[style] [weight] <size>px <families>`),
 * judged by `oracle` in the painting realm. `unicodeFont` gives the Canvas
 * font converted text paints with, for the colour proof of a target.
 */
export function symbolPaint(
  oracle: SymbolFaceOracle,
  text: string,
  family: string | null | undefined,
  keptFont: string,
  unicodeFont: (converted: string) => string,
): SymbolPaint {
  if (!text || !isSymbolFontFamily(family)) return AUTHORED_PAINT;
  const pieces = new Map<string, string>();
  const native: string[] = [];
  let unicode = false;
  for (const ch of new Set(text)) {
    const cp = ch.codePointAt(0) ?? 0;
    const pua = cp >= 0xf020 && cp <= 0xf0ff;
    if (!pua && (cp < 0x20 || cp > 0xff)) continue;
    // A code without a conversion paints alike in either face, and is not
    // judged at all (spaces, digits of Symbol).
    if (!symbolTextToUnicodeSegments(ch, family, { native: () => false, monochrome: () => true })[0]?.mapped) continue;
    // A bare code the realm cannot judge paints as authored, as in 0.89.
    if (!pua && oracle.draws(ch, keptFont) === undefined) continue;
    const segment = oracle.segments(ch, family, keptFont, unicodeFont)[0];
    if (!segment) continue;
    if (segment.mapped) unicode = true;
    else if (oracle.draws(ch, keptFont) === true) native.push(ch);
    if (segment.text !== ch) pieces.set(ch, segment.text);
  }
  if (unicode && native.length > 0) {
    // A partial face: the codes it draws convert too (module notes), with the
    // colour proof of symbol-face.ts in the Unicode face.
    const monochrome = (converted: string) => {
      const verdict = oracle.monochrome(converted, unicodeFont(converted));
      const base = converted.replace(TEXT_PRESENTATION, '');
      return verdict !== true || base === converted ? verdict : oracle.monochrome(base, unicodeFont(base));
    };
    for (const ch of native) {
      const segment = symbolTextToUnicodeSegments(ch, family, { native: () => false, monochrome })[0];
      if (segment?.mapped) pieces.set(ch, segment.text);
      else pieces.delete(ch);
    }
  }
  if (pieces.size === 0) return AUTHORED_PAINT;
  return Object.freeze({
    unicode,
    paint: (value: string) => {
      let out = '';
      for (const ch of value) out += pieces.get(ch) ?? ch;
      return out;
    },
  });
}

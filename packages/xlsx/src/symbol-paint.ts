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
 * A text can need two faces: a partial face draws some of its codes and
 * lacks others. Each converted code paints in the Unicode fallback face and
 * every other character (a code the face draws, one it cannot judge, one
 * whose conversion is refused, any character outside the symbol encoding)
 * keeps the authored face, as the core leaves those unmapped (`segments`).
 * So no glyph the authored face may draw moves to another face, and no
 * converted character paints in the symbol family, whose cmap would re-read
 * a target such as U+00A9 as its own symbol code. A text with both kinds
 * paints as several stretches, each in its own Canvas font.
 */
import { isSymbolFontFamily, type SymbolFaceOracle } from '@silurus/ooxml-core';

/** One maximal stretch of a text that paints in one face. */
export interface SymbolSegment {
  /** The authored text of the stretch. */
  readonly authored: string;
  /** True: it paints in the Unicode fallback face; false: in the authored face. */
  readonly unicode: boolean;
}

/** How one authored text paints. */
export interface SymbolPaint {
  /** The text's stretches in order; exactly one when it paints in one face. */
  readonly segments: readonly SymbolSegment[];
  /** The painted form of any text made of the authored text's characters
   *  (per code point: its conversion, its F0-shifted form or itself). */
  paint(text: string): string;
}

const identity = (text: string) => text;

/** A text painted as authored, in its authored face. */
export function authoredPaint(text: string): SymbolPaint {
  return { segments: [{ authored: text, unicode: false }], paint: identity };
}

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
  if (!text || !isSymbolFontFamily(family)) return authoredPaint(text);
  const pieces = new Map<string, string>();
  /** Codes that convert: they paint in the Unicode fallback face. */
  const converted = new Set<string>();
  for (const ch of new Set(text)) {
    const cp = ch.codePointAt(0) ?? 0;
    const pua = cp >= 0xf020 && cp <= 0xf0ff;
    if (!pua && (cp < 0x20 || cp > 0xff)) continue;
    // A bare code the realm cannot judge paints as authored, as in 0.89.
    if (!pua && oracle.draws(ch, keptFont) === undefined) continue;
    const segment = oracle.segments(ch, family, keptFont, unicodeFont)[0];
    if (segment?.mapped) converted.add(ch);
    if (segment && segment.text !== ch) pieces.set(ch, segment.text);
  }
  if (pieces.size === 0) return authoredPaint(text);
  const segments: SymbolSegment[] = [];
  for (const ch of text) {
    const unicode = converted.has(ch);
    const previous = segments[segments.length - 1];
    if (previous && previous.unicode === unicode) segments[segments.length - 1] = { authored: previous.authored + ch, unicode };
    else segments.push({ authored: ch, unicode });
  }
  return Object.freeze({
    segments,
    paint: (value: string) => {
      let out = '';
      for (const ch of value) out += pieces.get(ch) ?? ch;
      return out;
    },
  });
}

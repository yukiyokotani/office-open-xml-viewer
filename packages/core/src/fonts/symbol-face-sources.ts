/**
 * Held symbol encodings of served font-source faces (symbol-face.ts),
 * computed by the font-source session (LoadedFontSources.symbolEncodings),
 * which already loads the CSS descriptor helpers; the formats only read them.
 */
import { parseUnicodeRange } from './css-source.js';
import { fontFaceWeightCovers } from './office-fallback.js';
import { symbolCodeSet, symbolCodeSetHas, type SymbolFaceEncoding } from './symbol-font.js';

/**
 * The encoding of a served family's faces for one weight/style, when every
 * face CSS matching can pick for that tuple has a read cmap (CSS Fonts 4
 * §4.5, §5.3: faces whose `unicode-range` covers a scalar are tried, the
 * last-declared first, until one has a glyph for it).
 *   - A PUA code is owned when any covering face's cmap draws it.
 *   - A bare code is owned when the first covering face that has a glyph for
 *     it draws it as a symbol code; a Unicode face that maps it draws the
 *     Unicode character instead.
 * Undefined when no face matches the tuple exactly (matching then picks a
 * nearer face or synthesizes; the caller measures instead) or a matching
 * face's bytes were not read.
 */
export function symbolEncodingOfFaces(
  faces: readonly Readonly<{
    weight: string; style: string; stretch?: string; unicodeRange: string; symbolEncoding?: SymbolFaceEncoding;
  }>[],
  weight: number,
  style: 'normal' | 'italic',
): SymbolFaceEncoding | undefined {
  // Text is requested at the normal width; a face of another stretch is
  // matched only when the family has no normal one, which is left to measurement.
  const normalWidth = (stretch: string | undefined) => !stretch || /^(?:normal|100%)$/iu.test(stretch.trim());
  const matching = faces.filter((face) => fontFaceWeightCovers(face.weight, weight)
    && face.style.trim().toLowerCase() === style && normalWidth(face.stretch));
  if (matching.length === 0) return undefined;
  const ranges: (readonly (readonly [number, number])[])[] = [];
  const encodings: Extract<SymbolFaceEncoding, { kind: 'known' }>[] = [];
  for (const face of matching) {
    const parsed = parseUnicodeRange(face.unicodeRange);
    if (face.symbolEncoding?.kind !== 'known' || !parsed) return undefined;
    ranges.push(parsed);
    encodings.push(face.symbolEncoding);
  }
  const pua: number[] = [];
  const bare: number[] = [];
  const covering = (scalar: number) => encodings.filter((_, index) =>
    ranges[index]!.some(([lo, hi]) => lo <= scalar && scalar <= hi));
  for (let code = 0; code < 256; code++) {
    if (covering(0xf000 + code).some((encoding) => symbolCodeSetHas(encoding.pua, code))) pua.push(code);
    // The last-declared covering face that draws the bare code at all decides:
    // as its symbol code (owned), or as the Unicode character (not owned).
    for (const encoding of covering(code).reverse()) {
      if (symbolCodeSetHas(encoding.bare, code)) { bare.push(code); break; }
      if (symbolCodeSetHas(encoding.latin, code)) break;
    }
  }
  return Object.freeze({ kind: 'known', pua: symbolCodeSet(pua), bare: symbolCodeSet(bare) });
}

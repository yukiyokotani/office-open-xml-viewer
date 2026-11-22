// GDI SYMBOL_CHARSET bytes are font codes, unlike the ordinary Unicode text
// consumed by symbol-font.ts. Windows symbol cmap entries use the private
// range beginning at F020 (OpenType cmap, platform 3 encoding 0):
// https://learn.microsoft.com/en-us/typography/opentype/spec/cmap
// A code is converted to Unicode only from the selected family's published
// repertoire. Unknown glyphs/fonts are never guessed or decoded as ANSI.

import { ADOBE_SYMBOL_UNICODE } from './symbol-encoding.js';

// Symbol: the published Adobe Symbol repertoire (symbol-encoding.ts). Metafile
// text keeps the table's Corporate-Use codes (bracket/integral/brace pieces,
// extenders, serif/sans copyright marks) font-coded: MathType-style metafiles
// stack those pieces at the authored face's exact metrics, so an installed
// Symbol still draws them rather than a Unicode fallback's unrelated pieces.

const SYMBOL_ENCODING: Readonly<Record<number, string>> = Object.fromEntries(
  Object.entries(ADOBE_SYMBOL_UNICODE).map(([code, unicode]) => [code, String.fromCodePoint(unicode)]),
);

// Mozilla's published MT Extra encoding (the reduced Design Science face):
// https://www-archive.mozilla.org/projects/mathml/fonts/encoding/mtextra
// The five ellipsis positions also agree with Design Science 4.31's enhanced
// font cmap and outlines. This is a bounded font-code conversion, not a claim
// that every same-named version shares the full repertoire: older/enhanced
// versions differ elsewhere. No fonts or font-specific glyph indices are
// embedded.
const MT_EXTRA_ENCODING: Readonly<Record<number, string>> = {
  0x20: ' ',
  0x4b: '…', // U+2026 horizontal (baseline) ellipsis
  0x4c: '⋯', // U+22EF math-axis ellipsis, distinct from U+2026
  0x4d: '⋮', // U+22EE vertical ellipsis
  0x4e: '⋰', // U+22F0 up-right diagonal ellipsis
  0x4f: '⋱', // U+22F1 down-right diagonal ellipsis
};

function symbolTable(face: string): Readonly<Record<number, string>> | null {
  const family = face.trim().toLowerCase();
  return family === 'symbol' ? SYMBOL_ENCODING
    : family === 'mt extra' ? MT_EXTRA_ENCODING : null;
}

/** True when `face` has a published symbol repertoire here. */
export function hasMetafileSymbolRepertoire(face: string): boolean {
  return symbolTable(face) !== null;
}

/** The SYMBOL_CHARSET face a Windows font mapper substitutes for a symbol
 *  face the system lacks: Symbol is the symbol-charset face Windows ships, so
 *  the charset match (weighted above the face name) selects it. */
export const SYMBOL_CHARSET_SUBSTITUTE_FACE = 'Symbol';

/** The Unicode character for one symbol-font code (0x00..0xFF) of `face`, or
 *  undefined when the family has no published repertoire or the code has no
 *  standard Unicode target. Callers keep such codes font-coded (U+F000+code)
 *  in the authored face, as Windows does for a symbol cmap. */
export function metafileSymbolCodeToUnicode(code: number, face: string): string | undefined {
  return symbolTable(face)?.[code];
}

/** Decode a fully known font-coded string, or null for an unsupported encoding.
 *  EMF UTF-16 only remaps private F000..F0FF entries; WMF remaps its byte codes.
 *  Callers must draw decoded text with a Unicode fallback font, not the source
 *  symbol font (which could reinterpret the new Unicode points). */
export function decodeMetafileSymbolText(
  text: string,
  face: string,
  privateUseOnly = false,
): string | null {
  const table = symbolTable(face);
  if (!table) return null;
  let result = '';
  for (const char of text) {
    const code = char.codePointAt(0)!;
    if (privateUseOnly && (code < 0xf000 || code > 0xf0ff)) {
      result += char;
      continue;
    }
    const target = table[privateUseOnly ? code - 0xf000 : code];
    if (target === undefined) return null;
    result += target;
  }
  return result;
}

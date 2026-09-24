/**
 * Spreadsheet cell font stacks: the CSS font-family list the XLSX renderer
 * paints a cell font with. Shared so that every spreadsheet producer measures
 * the Normal font's maximum digit width through the same stack the grid is
 * painted with (ECMA-376 §18.3.1.13), including the painted fallback when the
 * authored face is unavailable.
 */
import {
  classifyCjkFont,
  classifyFontGeneric,
  cjkFallbackChain,
  NON_CJK_SANS_FALLBACKS,
  NON_CJK_SERIF_FALLBACKS,
} from '../fonts/scripts';

// Default font stack. Calibri is the workbook default font in Excel; on
// systems without Office (macOS / Linux) the browser would otherwise fall
// back to Arial / Helvetica, which is meaningfully wider than Calibri at
// every weight/size combination. Carlito is the Google-released, metric-
// compatible Calibri clone (same advance widths and ascender / descender
// metrics) and is loaded opt-in by `XlsxWorkbook.load({ useGoogleFonts:
// true })`. Listing it in the cascade means: Calibri (Windows / Office)
// → Carlito (loaded webfont) → Arial → sans-serif. Caladea is the same
// for Cambria.
// The two trailing Noto Arabic faces are generic Arabic-script fallbacks:
// when the primary Latin faces (Calibri / Carlito / Arial) lack a requested
// glyph, the browser advances down the cascade per-glyph, so any Arabic
// codepoint resolves to a real web font (loaded by `XlsxWorkbook.load`'s
// useGoogleFonts path) instead of an oversized OS Arabic face. Latin glyphs
// still bind to the earlier faces, so Latin rendering is unchanged.
// The trailing non-CJK Noto faces (Hebrew / Thai / Devanagari, plus "Noto Sans"
// for Cyrillic) extend the same per-glyph fallback idea to the other
// non-Latin, non-CJK scripts: any such codepoint resolves to a real web font
// (loaded opt-in via useGoogleFonts) instead of an OS face or tofu. CJK is NOT
// appended here — shared Han glyphs differ in shape per language, so the
// correct Noto CJK is chosen per cell from the cell's font name; see
// fontStackFor() / cssTailFor().
const NON_CJK_SANS_TAIL = NON_CJK_SANS_FALLBACKS.map((n) => `"${n}"`).join(', ');
const NON_CJK_SERIF_TAIL = NON_CJK_SERIF_FALLBACKS.map((n) => `"${n}"`).join(', ');
export const DEFAULT_FONT_FAMILY =
  `"Calibri", "Carlito", "Cambria", "Caladea", Arial, "Noto Naskh Arabic", "Noto Sans Arabic", ${NON_CJK_SANS_TAIL}, sans-serif`;
// Serif counterpart of DEFAULT_FONT_FAMILY. A Latin *serif* cell font the host
// lacks (Century, Garamond, …) must degrade to a serif — Excel renders such a
// cell with a serif, not the sans default. Cambria is Office's serif; Caladea is
// its metric-compatible clone (loaded opt-in via useGoogleFonts), then web-safe
// serifs, ending in the `serif` generic.
const DEFAULT_SERIF_FONT_FAMILY =
  `"Cambria", "Caladea", "Times New Roman", "Liberation Serif", "Noto Naskh Arabic", "Noto Sans Arabic", ${NON_CJK_SERIF_TAIL}, serif`;
// Monospace counterpart: a monospaced cell font the host lacks degrades to a
// monospace generic rather than the proportional sans default.
const DEFAULT_MONO_FONT_FAMILY = `"Courier New", "Liberation Mono", monospace`;

/**
 * CSS font-family TAIL (everything after the cell's named face) for an xlsx
 * cell. For a CJK cell font the matching Noto CJK leads (so shared Han glyphs
 * take the document language's shapes; see core/fonts/scripts.ts), followed by
 * the standard Latin/Arabic/non-CJK fallbacks. A non-CJK cell font picks the
 * default chain by its generic class ({@link classifyFontGeneric}) so a Latin
 * serif/mono face the host lacks degrades to the matching generic. Exported for
 * unit testing.
 */
export function cssTailFor(name: string | null | undefined): string {
  const cjk = name ? classifyCjkFont(name) : null;
  const generic = classifyFontGeneric(name); // 'serif' | 'sans' | 'mono'
  if (!cjk) {
    // Non-CJK (Latin) cell font: choose the default chain by generic class so a
    // Latin serif/mono face the host lacks degrades to the matching generic
    // (Excel renders serif/mono, not the sans default).
    if (generic === 'serif') return DEFAULT_SERIF_FONT_FAMILY;
    if (generic === 'mono') return DEFAULT_MONO_FONT_FAMILY;
    return DEFAULT_FONT_FAMILY;
  }
  const serif = generic === 'serif';
  const cjkPart = cjkFallbackChain(cjk, serif ? 'serif' : 'sans')
    .map((n) => `"${n}"`)
    .join(', ');
  const tail = serif ? NON_CJK_SERIF_TAIL : NON_CJK_SANS_TAIL;
  const genericKeyword = serif ? 'serif' : 'sans-serif';
  // CJK Noto leads, then Latin/metric substitutes, Arabic, non-CJK scripts.
  return `${cjkPart}, "Calibri", "Carlito", "Cambria", "Caladea", Arial, "Noto Naskh Arabic", "Noto Sans Arabic", ${tail}, ${genericKeyword}`;
}

/** Full CSS font-family list for a cell font name (named face first). */
export function fontStackFor(name: string | null | undefined): string {
  return name ? `"${name}", ${cssTailFor(name)}` : DEFAULT_FONT_FAMILY;
}


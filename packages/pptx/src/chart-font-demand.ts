import {
  excelDateTimeSectionNameTexts,
  isDateFormatSection,
  localizedExcelShortDateTexts,
  splitFormatSections,
} from '@silurus/ooxml-core';

/**
 * The texts a DrawingML chart may paint, for `unicode-range` font demand
 * (library loading policy, not OOXML rendering behaviour). The chart painter
 * (core `renderChart`) lays its labels out at paint time, so the exact label
 * strings (tick values, elided or wrapped labels, formatted data labels) are
 * not known before render. Every scalar they can contain is, though:
 *
 * - authored text: every string of the canonical chart model (titles, rich
 *   runs, category and series caches, legend entries, data-label fields, the
 *   format codes whose literals a formatted label copies). Face names are not
 *   painted and are left out;
 * - generated numbers and dates: §18.8.30 formatting writes ASCII digits,
 *   signs, separators and exponents (printable ASCII below), the code's own
 *   literals (above) and, for a date / time section, the formatter's name
 *   tables (core `excelDateTimeSectionNameTexts`: month, weekday, AM/PM,
 *   Japanese weekday, era and quarter names). Every string under a key naming
 *   a format (`/format/i`, array items inheriting their array's key) is
 *   tested as a code. Excel built-in format 14 in a
 *   data table paints the viewer's locale short date (`Intl`), complete only
 *   for a Gregorian, Latin-digit locale; any other locale declines;
 * - the painter's own literals ({@link CHART_SYNTHESIZED_TEXT}): elision,
 *   histogram bin bounds and trendline equations.
 *
 * `'all'` (every subset) when that set is not provable: a region map (its
 * labels come from bundled geodata names), a metafile or unknown image fill
 * (a metafile plays its own text records), or a non-Gregorian / native-digit
 * locale short date. Strings are collected by the caller's scalar collector,
 * which still declines RTL, PUA and non-inert scalars.
 */
export function chartPaintTexts(chart: unknown): readonly string[] | 'all' {
  const texts: string[] = [CHART_SYNTHESIZED_TEXT];
  let localizedShortDate = false;
  const pending: unknown[] = [chart];
  const seen = new Set<object>();
  const formatArrays = new Set<unknown[]>();
  while (pending.length) {
    const value = pending.pop();
    if (value === null || typeof value !== 'object' || seen.has(value) || ArrayBuffer.isView(value)) continue;
    seen.add(value);
    const record = value as Record<string, unknown>;
    if (!Array.isArray(value)) {
      if (record.chartType === 'regionMap') return 'all';
      if (typeof record.imagePath === 'string' && !paintsTextFreeImage(record.imagePath, record.mimeType)) return 'all';
      if (record.catFormatBuiltinId === 14) localizedShortDate = true;
    }
    // A string under a format key (`formatCode`, `valAxisFormatCode`,
    // `catFormatCodes[i]`, `numberFormat.formatCode`, …) may be a §18.8.30
    // code; the parent's key carries over to array items.
    const formatArray = Array.isArray(value) && formatArrays.has(value);
    for (const key in record) {
      const child = record[key];
      if (typeof child === 'string') {
        if (CHART_FACE_KEY.test(key)) continue;
        texts.push(child);
        if (!formatArray && !FORMAT_KEY.test(key)) continue;
        for (const section of splitFormatSections(child)) {
          if (!isDateFormatSection(section)) continue;
          texts.push(...excelDateTimeSectionNameTexts(section));
          // Under ja-JP `rr` writes the era name too (`gggee`), which the
          // name table above reads only from `g`: ask it as if `g` were present.
          if (/r/i.test(section)) texts.push(...excelDateTimeSectionNameTexts(`${section}g`));
        }
      } else if (child !== null && typeof child === 'object') {
        if (Array.isArray(child) && FORMAT_KEY.test(key)) formatArrays.add(child);
        pending.push(child);
      }
    }
  }
  if (localizedShortDate) {
    const shortDate = localizedExcelShortDateTexts();
    if (!shortDate.complete) return 'all';
    texts.push(...shortDate.texts);
  }
  return texts;
}

/** Image parts that paint no text with a served face: rasters (decoded by the
 * browser) and SVG (decoded through an `<img>`, which cannot reach document
 * fonts). A metafile (WMF/EMF, played by the library's players, which draw
 * text records with served faces) or an unknown type is not admitted.
 *
 * The type is the parser's extension mapping (`mime_from_ext`), as the
 * package declares the part. The decoder also sniffs content, so a metafile
 * stored under a raster extension (a mislabelled part) still plays; its text
 * is not demanded and paints in a host fallback where no loaded subset covers
 * it (no late load: an undemanded subset is never registered). Sniffing every
 * image part during preflight would inflate every picture before first
 * paint; DOCX demand makes the same trade. */
export function paintsTextFreeImage(path: unknown, mimeType: unknown): boolean {
  if (typeof path === 'string' && METAFILE_PATH.test(path)) return false;
  return typeof mimeType === 'string' && TEXT_FREE_IMAGE_TYPES.has(mimeType.toLowerCase());
}

const TEXT_FREE_IMAGE_TYPES: ReadonlySet<string> = new Set([
  'image/png', 'image/jpeg', 'image/jpg', 'image/gif', 'image/bmp', 'image/webp', 'image/tiff', 'image/svg+xml',
]);
const METAFILE_PATH = /\.(?:wmf|emf|wmz|emz)$/i;

/** Keys holding a §18.8.30 format code: every ChartModel code field
 * (`*FormatCode`, `*FormatCodes`, `formatCode` of a number-format object). */
const FORMAT_KEY = /format/i;

/** Keys naming a face (core chart-font-sources CHART_FACE_KEY): not painted. */
const CHART_FACE_KEY = /^(?:fontFace|[A-Za-z]+FontFace|themeMajorFontLatin|themeMinorFontLatin)$/;

/** Printable ASCII (numbers, signs, separators, generated series and
 * trendline names) plus the non-ASCII literals the chart painter writes: the
 * elision ellipsis (text-elide, data-label-layout), histogram bin bounds
 * `≤ ≥` (histogram-binning), and the trendline equation minus sign and `R²`
 * (trendline). U+FFFD: elision's binary search measures UTF-16 prefixes that
 * may end in a lone surrogate, which Canvas measures as the replacement
 * character. */
export const CHART_SYNTHESIZED_TEXT = `${String.fromCharCode(...Array.from({ length: 0x7f - 0x20 }, (_, i) => 0x20 + i))}…≤≥−²\ufffd`;

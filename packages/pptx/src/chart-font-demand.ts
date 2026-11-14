import { localizedExcelShortDateTexts } from '@silurus/ooxml-core';

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
 *   signs, separators, exponents, English month and day names and AM/PM
 *   (printable ASCII below), plus the code's own literals (above), plus the
 *   Japanese weekday, era and quarter names of a date section
 *   ({@link CHART_DATE_SYNTHESIZED_TEXT}). Excel built-in format 14 in a data
 *   table paints the viewer's locale short date (`Intl`), complete only for a
 *   Gregorian, Latin-digit locale; any other locale declines;
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
  let dateSection = false;
  let localizedShortDate = false;
  const pending: unknown[] = [chart];
  const seen = new Set<object>();
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
    for (const [key, child] of Object.entries(record)) {
      if (typeof child === 'string') {
        if (CHART_FACE_KEY.test(key)) continue;
        texts.push(child);
        if (FORMAT_CODE_KEY.test(key) && formatCodeMayFormatDate(child)) dateSection = true;
      } else if (child !== null && typeof child === 'object') {
        if (FORMAT_CODE_KEY.test(key) && Array.isArray(child)
          && child.some((code) => typeof code === 'string' && formatCodeMayFormatDate(code))) dateSection = true;
        pending.push(child);
      }
    }
  }
  if (dateSection) texts.push(...CHART_DATE_SYNTHESIZED_TEXT);
  if (localizedShortDate) {
    const shortDate = localizedExcelShortDateTexts();
    if (!shortDate.complete) return 'all';
    texts.push(...shortDate.texts);
  }
  return texts;
}

/** Image parts that paint no text with a served face: rasters (decoded by the
 * browser) and SVG (decoded through an `<img>`, which cannot reach document
 * fonts). A metafile (WMF/EMF, decoded by the library's players, which draw
 * text records with served faces) or an unknown type is not admitted. The
 * type is the parser's extension mapping (`mime_from_ext`); a metafile part
 * mislabelled with a raster extension is not detected (as in DOCX demand). */
export function paintsTextFreeImage(path: unknown, mimeType: unknown): boolean {
  if (typeof path === 'string' && METAFILE_PATH.test(path)) return false;
  return typeof mimeType === 'string' && TEXT_FREE_IMAGE_TYPES.has(mimeType.toLowerCase());
}

const TEXT_FREE_IMAGE_TYPES: ReadonlySet<string> = new Set([
  'image/png', 'image/jpeg', 'image/jpg', 'image/gif', 'image/bmp', 'image/webp', 'image/tiff', 'image/svg+xml',
]);
const METAFILE_PATH = /\.(?:wmf|emf|wmz|emz)$/i;

/** Keys naming a face (core chart-font-sources CHART_FACE_KEY): not painted. */
const CHART_FACE_KEY = /^(?:fontFace|[A-Za-z]+FontFace|themeMajorFontLatin|themeMinorFontLatin)$/;
/** Keys holding §18.8.30 format codes (`formatCode`, `valAxisFormatCode`,
 * `catFormatCodes`, …). */
const FORMAT_CODE_KEY = /formatCodes?$/i;

/**
 * Whether a format code may contain a date/time section, so its output may
 * contain the Japanese weekday, era or quarter names. An over-approximation:
 * after removing General, quoted and escaped literals and bracketed
 * metadata, any date-capable letter (`y m d h s g e a b q`, except the
 * scientific `E+`/`E-`) counts.
 */
function formatCodeMayFormatDate(code: string): boolean {
  const stripped = code
    .replace(/"[^"]*"?/g, '')
    .replace(/\\./g, '')
    .replace(/[_*]./g, '')
    .replace(/\[[^\]]*\]?/g, '')
    .replace(/general/gi, '')
    .replace(/[eE][+-]/g, '');
  return /[ymdhsgeabqYMDHSGEABQ]/.test(stripped);
}

/** Printable ASCII (numbers, signs, separators, English month and day names,
 * AM/PM, generated series and trendline names) plus the non-ASCII literals
 * the chart painter writes: the elision ellipsis (text-elide,
 * data-label-layout), histogram bin bounds `≤ ≥` (histogram-binning), and
 * the trendline equation minus sign and `R²` (trendline). */
export const CHART_SYNTHESIZED_TEXT = `${String.fromCharCode(...Array.from({ length: 0x7f - 0x20 }, (_, i) => 0x20 + i))}…≤≥−²`;

/** The non-ASCII names a §18.8.30 date section may paint (core
 * excel-number-format: `aaa`/`aaaa` weekdays, `g`/`gg`/`ggg` eras, quarters). */
export const CHART_DATE_SYNTHESIZED_TEXT: readonly string[] = Object.freeze([
  '日月火水木金土', '曜日',
  '令和平成昭和大正明治',
  '第１２３４四半期',
]);

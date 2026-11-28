// Decimal and digit-group separators painted by spreadsheet number formats.
//
// ECMA-376 §18.8.30 defines `.` and `,` in a format code as the decimal
// point and the thousands separator tokens. [MS-OI29500] §2.1.712 states that
// grammar for en-US and has Office adjust it for international settings, with
// INTL-CHAR-DECIMAL-SEP as the fractional-seconds marker. Viewer policy
// (#2346): paint those tokens with the host locale's separators, the locale
// built-in format 14 already follows (`formatLocalizedExcelShortDate`):
// the main thread's `navigator.language`. Neither specification names that
// locale; it stands in for the reader's regional settings, which a browser
// library cannot read. A worker adopts the main thread's value from its parse
// request (`adoptHostLocale`) rather than reading its own: browser locale
// emulation (a test harness's per-page locale) does not reach dedicated
// workers, and paint and font demand must agree across the two realms.
//
// Scope: separators only. Which text a site paints and which format code
// applies to it are untouched; the one output difference is the decimal and
// group characters of numbers (and the font demand for them), and en-US
// output is unchanged. Formatters paint the host symbols themselves. A chart
// category painted as its raw string has no formatter: where that string is
// the invariant text of a number, its `.` is repainted through
// `localizedCategoryText` (chart-number-format.ts), keyed by the provenance
// the parser records for that list (`categoriesNumeric`).
//
// Library policy and its limits:
// - Only the two symbols are localized. Digits stay ASCII and groups stay
//   three digits wide: the symbols come from the locale's Latin-digit
//   numbering system (`numberingSystem: 'latn'`), so ar-EG paints `1,234.5`
//   rather than Arabic separators between ASCII digits, and en-IN keeps
//   `1,234,567` rather than lakh grouping.
// - The symbols are Intl's, as is: fr-FR groups with U+202F NARROW NO-BREAK
//   SPACE (CLDR), where some Windows builds use U+00A0 or U+0020; the
//   difference is one space glyph. Font demand collects the painted
//   symbols (`numberSymbolTexts`).

export interface NumberSymbols {
  readonly decimal: string;
  readonly group: string;
}

/** en-US separators: what a format code stores, and today's default. */
export const INVARIANT_NUMBER_SYMBOLS: NumberSymbols = Object.freeze({ decimal: '.', group: ',' });

const symbolCache = new Map<string, NumberSymbols>();
let adoptedLocale: string | undefined;

/** The host locale number symbols follow: the value adopted from the main
 * thread in a worker, else this realm's `navigator.language`. Main-thread
 * loaders send it with every worker parse request. */
export function hostLocale(): string | undefined {
  return adoptedLocale ?? (typeof navigator === 'undefined' ? undefined : navigator.language);
}

/** Adopt the main thread's host locale in a worker (a parse request's
 * `hostLocale`). A non-string value leaves the current locale in place. */
export function adoptHostLocale(locale: unknown): void {
  if (typeof locale === 'string' && locale) adoptedLocale = locale;
}

/** Adopt `message.hostLocale` when a worker message carries one. */
export function adoptMessageHostLocale(message: unknown): void {
  if (message && typeof message === 'object' && 'hostLocale' in message) {
    adoptHostLocale((message as { hostLocale?: unknown }).hostLocale);
  }
}

/** The decimal and group separators of `locale` (host locale by default). */
export function hostNumberSymbols(locale = hostLocale()): NumberSymbols {
  const key = locale ?? '';
  const cached = symbolCache.get(key);
  if (cached) return cached;
  let symbols = INVARIANT_NUMBER_SYMBOLS;
  try {
    const parts = new Intl.NumberFormat(locale, { numberingSystem: 'latn', useGrouping: true })
      .formatToParts(1234567.5);
    const decimal = parts.find((part) => part.type === 'decimal')?.value;
    const group = parts.find((part) => part.type === 'group')?.value;
    // A locale that names the same symbol for both would make the output
    // ambiguous; keep the invariant pair then.
    if (decimal && group && decimal !== group) {
      symbols = decimal === '.' && group === ',' ? INVARIANT_NUMBER_SYMBOLS : Object.freeze({ decimal, group });
    }
  } catch {
    // An invalid tag (RangeError) keeps the invariant separators.
  }
  symbolCache.set(key, symbols);
  return symbols;
}

/** Group an ASCII integer digit run in threes with `group`. */
export function groupDigits(intDigits: string, group: string): string {
  return intDigits.replace(/\B(?=(\d{3})+(?!\d))/g, group);
}

/** The separator texts a number format can paint under `symbols`, for font
 * demand: the invariant pair needs no entry beyond the format's own text. */
export function numberSymbolTexts(symbols: NumberSymbols = hostNumberSymbols()): readonly string[] {
  return symbols === INVARIANT_NUMBER_SYMBOLS ? [] : [symbols.decimal + symbols.group];
}

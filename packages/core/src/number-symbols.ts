// Decimal and digit-group separators painted by spreadsheet number formats.
//
// ECMA-376 §18.8.30 stores `.` and `,` in a format code whatever the
// author's locale: they stand for "the decimal separator" and "the
// thousands separator", which Excel takes from the OS regional settings,
// not from the file (`#,##0.00` paints `1 234,50` on a fr-FR system and
// `1.234,50` on de-DE). The host locale is the viewer's equivalent of those
// settings, as it already is for built-in format 14
// (`formatLocalizedExcelShortDate`), so both resolve the same locale:
// `navigator.language`, which a worker sees as well as the main thread.
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

/** The decimal and group separators of `locale` (host locale by default). */
export function hostNumberSymbols(
  locale = typeof navigator === 'undefined' ? undefined : navigator.language,
): NumberSymbols {
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

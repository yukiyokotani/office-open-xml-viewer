// CSS `unicode-range` parsing, a light module of its own: DOCX layout reads
// the declared ranges of loaded faces (application and served) on every
// document, while the rest of css-source.ts (the stylesheet scanner and the
// CSS adapters) loads only with the font-source session.

/**
 * Parse a CSS `unicode-range` value (CSS Fonts 4 §4.5: comma-separated
 * `U+hex`, `U+hex-hex` or `U+hex?` wildcard). Returns null unless every
 * component is recognised and within U+0–U+10FFFF with start ≤ end: a range
 * is an exclusion hint, never a support certificate, so an unparseable value
 * means "covers everything".
 */
export function parseUnicodeRange(value: string | undefined): readonly (readonly [number, number])[] | null {
  if (!value?.trim()) return null;
  const ranges: [number, number][] = [];
  for (const component of value.split(',')) {
    const part = component.trim();
    let lo: number, hi: number;
    const span = /^U\+([0-9A-F]{1,6})(?:-([0-9A-F]{1,6}))?$/i.exec(part);
    const wildcard = /^U\+([0-9A-F]{0,5})(\?{1,6})$/i.exec(part);
    if (span) {
      lo = parseInt(span[1], 16);
      hi = parseInt(span[2] ?? span[1], 16);
    } else if (wildcard && wildcard[1].length + wildcard[2].length <= 6) {
      lo = parseInt(wildcard[1] + '0'.repeat(wildcard[2].length), 16);
      hi = parseInt(wildcard[1] + 'F'.repeat(wildcard[2].length), 16);
    } else {
      return null;
    }
    if (lo > hi || hi > 0x10ffff) return null;
    ranges.push([lo, hi]);
  }
  return ranges;
}

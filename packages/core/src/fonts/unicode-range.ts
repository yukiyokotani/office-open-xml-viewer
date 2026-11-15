// CSS `unicode-range` parsing, a light module of its own: DOCX layout reads
// the declared ranges of loaded faces (application and served) on every
// document, while the rest of css-source.ts (the stylesheet scanner and the
// CSS adapters) loads only with the font-source session.

const isHexDigit = (unit: number): boolean =>
  (unit >= 0x30 && unit <= 0x39) || (unit >= 0x41 && unit <= 0x46) || (unit >= 0x61 && unit <= 0x66);
/** White space around a component: what `String.prototype.trim` removes. */
const isSpace = (unit: number): boolean =>
  unit === 0x20 || (unit >= 0x9 && unit <= 0xd) || unit === 0xa0 || unit === 0x1680 || (unit >= 0x2000 && unit <= 0x200a)
  || unit === 0x2028 || unit === 0x2029 || unit === 0x202f || unit === 0x205f || unit === 0x3000 || unit === 0xfeff;

/**
 * {@link parseUnicodeRange} as flat `[lo0, hi0, lo1, hi1, …]` pairs, scanned
 * in one pass without a string or array per range (a CJK family's
 * stylesheet declares some 17,000 ranges).
 */
export function parseUnicodeRangePairs(value: string | undefined): Uint32Array | null {
  if (value === undefined) return null;
  const length = value.length;
  const pairs: number[] = [];
  let at = 0;
  for (;;) {
    while (at < length && isSpace(value.charCodeAt(at))) at++;
    // `U+` (either case).
    if (at + 1 >= length || (value.charCodeAt(at) | 0x20) !== 0x75 || value.charCodeAt(at + 1) !== 0x2b) return null;
    at += 2;
    let lo = 0, digits = 0;
    for (; at < length && digits < 6; at++, digits++) {
      const unit = value.charCodeAt(at);
      if (!isHexDigit(unit)) break;
      lo = lo * 16 + (unit <= 0x39 ? unit - 0x30 : (unit | 0x20) - 0x57);
    }
    let hi = lo;
    const next = at < length ? value.charCodeAt(at) : -1;
    if (next === 0x3f) {
      // `U+hex?`: each `?` stands for any hex digit, six digits in all at most.
      let wild = 0;
      for (; at < length && value.charCodeAt(at) === 0x3f; at++) wild++;
      if (digits + wild > 6) return null;
      lo *= 16 ** wild;
      hi = lo + 16 ** wild - 1;
    } else {
      if (digits === 0) return null;
      if (next === 0x2d) {
        at++;
        hi = 0;
        let endDigits = 0;
        for (; at < length && endDigits < 6; at++, endDigits++) {
          const unit = value.charCodeAt(at);
          if (!isHexDigit(unit)) break;
          hi = hi * 16 + (unit <= 0x39 ? unit - 0x30 : (unit | 0x20) - 0x57);
        }
        if (endDigits === 0) return null;
      }
    }
    while (at < length && isSpace(value.charCodeAt(at))) at++;
    if (lo > hi || hi > 0x10ffff) return null;
    pairs.push(lo, hi);
    if (at >= length) return Uint32Array.from(pairs);
    if (value.charCodeAt(at) !== 0x2c) return null;
    at++;
  }
}

/**
 * Parse a CSS `unicode-range` value (CSS Fonts 4 §4.5: comma-separated
 * `U+hex`, `U+hex-hex` or `U+hex?` wildcard). Returns null unless every
 * component is recognised and within U+0–U+10FFFF with start ≤ end: a range
 * is an exclusion hint, never a support certificate, so an unparseable value
 * means "covers everything".
 */
export function parseUnicodeRange(value: string | undefined): readonly (readonly [number, number])[] | null {
  const pairs = parseUnicodeRangePairs(value);
  if (!pairs) return null;
  const ranges: [number, number][] = [];
  for (let index = 0; index < pairs.length; index += 2) ranges.push([pairs[index]!, pairs[index + 1]!]);
  return ranges;
}

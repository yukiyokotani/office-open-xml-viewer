import { describe, expect, it } from 'vitest';
import { parseUnicodeRange, parseUnicodeRangePairs } from './unicode-range.js';

/** The grammar as two regular expressions per trimmed component (the parser
 * before the one-pass scanner), kept as the reference. */
function reference(value: string | undefined): [number, number][] | null {
  if (!value?.trim()) return null;
  const ranges: [number, number][] = [];
  for (const component of value.split(',')) {
    const part = component.trim();
    let lo: number, hi: number;
    const span = /^U\+([0-9A-F]{1,6})(?:-([0-9A-F]{1,6}))?$/i.exec(part);
    const wildcard = /^U\+([0-9A-F]{0,5})(\?{1,6})$/i.exec(part);
    if (span) {
      lo = parseInt(span[1]!, 16);
      hi = parseInt(span[2] ?? span[1]!, 16);
    } else if (wildcard && wildcard[1]!.length + wildcard[2]!.length <= 6) {
      lo = parseInt(wildcard[1] + '0'.repeat(wildcard[2]!.length), 16);
      hi = parseInt(wildcard[1] + 'F'.repeat(wildcard[2]!.length), 16);
    } else {
      return null;
    }
    if (lo > hi || hi > 0x10ffff) return null;
    ranges.push([lo, hi]);
  }
  return ranges;
}

describe('parseUnicodeRangePairs', () => {
  it('returns the ranges as flat pairs', () => {
    expect(parseUnicodeRangePairs('U+0-7F, u+30??,U+4e00')).toEqual(Uint32Array.of(0, 0x7f, 0x3000, 0x30ff, 0x4e00, 0x4e00));
    expect(parseUnicodeRangePairs('U+??????')).toBeNull();
    expect(parseUnicodeRangePairs('U+10????')).toEqual(Uint32Array.of(0x100000, 0x10ffff));
    expect(parseUnicodeRangePairs(undefined)).toBeNull();
  });

  it.each([
    '', ' ', ',', 'U+0,', ',U+0', 'U+0,,U+1', 'U+', 'U+-5', 'U+5-', 'U+5-3', 'U+1234567', 'U+12 34', 'U+1 - 2', 'U+1?-2',
    'U+0-7F?', 'U+?', 'U+???????', 'U+123456?', 'U+11????', 'U+0-110000', 'U+G', 'X+0', 'U-0', 'U+0x', 'U+0-10FFFF',
    'u+a-B', '\n\tU+20 , U+30　', 'U+20᠎', 'U+20\u0085', '﻿U+20', 'U+0-0', 'U+000000-00000F',
  ])('agrees with the reference grammar on %j', (value) => {
    expect(parseUnicodeRange(value)).toEqual(reference(value));
  });

  it('agrees with the reference grammar on generated values', () => {
    const pieces = ['U', 'u', '+', '-', '?', ',', ' ', '\t', '\n', '0', '1', 'F', 'f', 'a', '9', 'G', 'x', ' ', '　',
      'U+', 'U+1F', '10FFFF', '110000', '﻿', ' ', '᠎', '\u0085', '\v', '\f', '\r'];
    let seed = 12345;
    const random = () => (seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x80000000;
    let parsed = 0;
    for (let round = 0; round < 200_000; round++) {
      let value = '';
      for (let length = 1 + Math.floor(random() * 12); length > 0; length--) value += pieces[Math.floor(random() * pieces.length)];
      const expected = reference(value);
      if (expected) parsed++;
      const actual = parseUnicodeRange(value);
      if (JSON.stringify(actual) !== JSON.stringify(expected)) expect({ value, actual }).toEqual({ value, actual: expected });
    }
    // The generator reaches the accepting paths, not only rejections.
    expect(parsed).toBeGreaterThan(1000);
  });
});

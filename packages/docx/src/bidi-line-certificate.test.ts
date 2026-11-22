import { describe, expect, it } from 'vitest';
import { certifyNativeSegmentLineOrder } from '@silurus/ooxml-core/internal/bidi-line';
import { computeLineVisualOrder, lineOrderPolicy, type LineOrderModel } from './bidi-line.js';

// #2130 added a native certificate that skips the full UAX #9 resolution for
// some RTL lines. Whenever it certifies a line, its order and per-segment
// direction must equal the canonical resolution under the same production
// policy. Seeded random lines cover the piece kinds the certificate classifies
// (R/L/AN/EN text, CS/ES/ET/ON neutrals, spaces, tabs, objects) and the HL1
// rtl mark, digitsAsAN and authored-isolate hooks.
const PIECES = [
  'אבג', 'אבג ', 'שלום עולם', 'مرحبا', 'مرحبا ', 'abc', 'abc ', 'Word', 'x y',
  '12', '12.50', '1,234', '٣٤', '١٢.٥', '12 ', ' ', '  ', '.', ',', ':', '-', '+',
  '$', '%', '#', '(', ')', '[', ']', '"', '/', 'א.', '1א', 'a1', 'אb', 'א 1', '12-34',
  'א́', '\u{10400}', '\u{1E900}',
];

function mulberry32(seed: number) {
  return () => {
    seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function randomLine(random: () => number): unknown[] {
  const length = 1 + Math.floor(random() * 8);
  return Array.from({ length }, () => {
    const kind = random();
    if (kind < 0.1) return { isTab: true };
    if (kind < 0.13) return {};
    const segment: { text: string; rtl?: boolean; digitsAsAN?: boolean } = {
      text: PIECES[Math.floor(random() * PIECES.length)]!,
    };
    if (random() < 0.5) segment.rtl = true;
    if (random() < 0.3) segment.digitsAsAN = true;
    return segment;
  });
}

describe('native line-order certificate (#2130)', () => {
  it.each<LineOrderModel>(['line-level', 'word'])('agrees with full UAX #9 resolution, %s model', (model) => {
    const random = mulberry32(model === 'word' ? 2130 : 9);
    const policy = lineOrderPolicy(model);
    let certified = 0;
    for (let i = 0; i < 40000; i++) {
      const line = randomLine(random);
      const certificate = certifyNativeSegmentLineOrder(line, true, policy);
      if (!certificate) continue;
      certified++;
      const expected = computeLineVisualOrder(line, true, model);
      expect({ order: certificate.order, rtl: certificate.rtl }, JSON.stringify(line))
        .toEqual({ order: [...expected.order], rtl: [...expected.rtl] });
    }
    // The certificate must actually be exercised, not refuse everything.
    expect(certified).toBeGreaterThan(2000);
  });
});

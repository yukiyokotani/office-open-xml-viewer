import { describe, expect, it } from 'vitest';
import {
  RTL_GATE,
  hasStrongRtl,
  OBJECT_PLACEHOLDER,
  buildVisualOrder,
} from './line-order.js';
import { getDefaultBidiEngine } from './engine.js';
import { certifyNativeSegmentLineOrder, computeSegmentLineVisualOrder } from './segment-line.js';

describe('native anchor certificate', () => {
  it('retains exact tab, numeric HL1 and space-only slice directions', () => {
    const segments = [{ isTab: true }, { text: '12.50', rtl: true },
      { text: ' ' }, { text: 'אבגד ', rtl: true }];
    const options = {
      isTab: (segment: unknown) => 'isTab' in (segment as object),
      isRtlMarked: (segment: unknown) => (segment as { rtl?: boolean }).rtl === true,
      isRtlAmbiguous: (character: string) => character === '.',
      directionFromAnyOdd: true,
    };
    expect(certifyNativeSegmentLineOrder(segments, true, options))
      .toEqual({ order: [3, 2, 1, 0], rtl: [true, true, false, true] });
    expect(computeSegmentLineVisualOrder(segments, true, options))
      .toEqual({ order: [3, 2, 1, 0], rtl: [true, true, false, true] });
    const prefixed = [{ text: 'AB' }, { text: 'CD' }, segments[0], segments[1],
      { text: ' ', rtl: true }, segments[3]];
    const wordOptions = {
      ...options,
      isolateDirection: (segment: unknown): 'ltr' | 'rtl' | undefined => {
        const s = segment as { text?: string; rtl?: boolean };
        return s.text === undefined ? undefined : s.rtl ? 'rtl' : 'ltr';
      },
      isolateRequiresUniformSlices: true,
    };
    const expected = { order: [5, 4, 3, 2, 0, 1], rtl: [false, false, true, true, false, true] };
    expect(certifyNativeSegmentLineOrder(prefixed, true, wordOptions)).toEqual(expected);
    expect(computeSegmentLineVisualOrder(prefixed, true, wordOptions)).toEqual(expected);
  });

  it('keeps the canonical UTF-16 HL1 boundary for astral L text', () => {
    const segments = [{ text: '\u{10400}' }, { text: 'א' }];
    const options = {
      isRtlMarked: () => true,
      // Canonical hooks receive individual code units, never this whole scalar.
      isRtlAmbiguous: (character: string) => character === '\u{10400}',
      directionFromAnyOdd: true,
    };
    const visual = certifyNativeSegmentLineOrder(segments, true, options)
      ?? computeSegmentLineVisualOrder(segments, true, options);
    expect(visual).toEqual({ order: [1, 0], rtl: [false, true] });
  });
});

describe('hasStrongRtl / RTL_GATE', () => {
  it('detects Arabic, Hebrew and RTL controls; ignores Latin/digits/neutrals', () => {
    expect(hasStrongRtl('مرحبا')).toBe(true); // Arabic
    expect(hasStrongRtl('שלום')).toBe(true); // Hebrew
    expect(hasStrongRtl('‏')).toBe(true); // RLM control
    expect(hasStrongRtl('Hello, world 123')).toBe(false);
    expect(hasStrongRtl('')).toBe(false);
  });

  it('detects Plane-1 RTL blocks (Adlam)', () => {
    expect(hasStrongRtl('\u{1E900}')).toBe(true); // Adlam letter
  });

  it('exposes the same regex used for the gate', () => {
    expect(RTL_GATE.test('שלום')).toBe(true);
    expect(RTL_GATE.test('abc')).toBe(false);
  });
});

describe('OBJECT_PLACEHOLDER', () => {
  it('is U+FFFC (OBJECT REPLACEMENT CHARACTER)', () => {
    expect(OBJECT_PLACEHOLDER).toBe('￼');
    expect(OBJECT_PLACEHOLDER.codePointAt(0)).toBe(0xfffc);
  });
});

describe('buildVisualOrder', () => {
  it('keeps LTR segments in logical order with even levels', () => {
    // "Hello " + "world", base LTR → levels all 0.
    const full = 'Hello world';
    const segStart = [0, 6];
    const { levels, paragraphLevel } = getDefaultBidiEngine().computeLevels(full, 'ltr');
    const { order, segLevels } = buildVisualOrder(levels, paragraphLevel, segStart);
    expect(order).toEqual([0, 1]);
    expect((segLevels[0] & 1) === 1).toBe(false);
    expect((segLevels[1] & 1) === 1).toBe(false);
  });

  it('reverses pure-RTL segments (visual L→R = logical last first)', () => {
    // Hebrew "שלום " + "עולם", base RTL.
    const s0 = 'שלום ';
    const full = s0 + 'עולם';
    const segStart = [0, s0.length];
    const { levels, paragraphLevel } = getDefaultBidiEngine().computeLevels(full, 'rtl');
    const { order, segLevels } = buildVisualOrder(levels, paragraphLevel, segStart);
    expect(order).toEqual([1, 0]);
    expect((segLevels[0] & 1) === 1).toBe(true);
  });

  it('falls back to the paragraph level for an X9-removed first unit', () => {
    // A lone object placeholder (neutral, removed by X9 in isolation) takes the
    // paragraph level rather than 255.
    const full = OBJECT_PLACEHOLDER;
    const { levels, paragraphLevel } = getDefaultBidiEngine().computeLevels(full, 'rtl');
    const { segLevels } = buildVisualOrder(levels, paragraphLevel, [0]);
    expect(segLevels[0]).toBe(paragraphLevel);
    expect(segLevels[0]).not.toBe(255);
  });
});

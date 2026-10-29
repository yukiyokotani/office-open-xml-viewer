import { describe, expect, it } from 'vitest';
import { contextualGraphemeAdvances } from './advance.js';
import type { LayoutTextSeg } from './model.js';
import type { TextLayoutService, TextShapeRequest } from '../layout/text.js';

/** Width of a window: per-character advances with an A–V pair interaction,
 * plus (when scoped) a term read from the surrounding run. */
function windowWidth(text: string, run: string, scoped: boolean): number {
  let advancePt = 0;
  for (let index = 0; index < text.length; index += 1) {
    advancePt += text.charCodeAt(index) % 7 + 1;
    if (index > 0 && text[index - 1] === 'A' && text[index] === 'V') advancePt -= 3;
  }
  return scoped ? advancePt + run.length / 100 : advancePt;
}

function countingService(scoped: boolean): { service: TextLayoutService; shapes: () => number } {
  let shapes = 0;
  const service = {
    fingerprint: 'counting',
    localMetrics: {},
    resolve: () => { throw new Error('unused'); },
    shape: (request: Readonly<TextShapeRequest>) => {
      shapes += 1;
      return {
        advancePt: windowWidth(request.text, request.substituteContext?.text ?? request.text, scoped),
      } as ReturnType<TextLayoutService['shape']>;
    },
  } as unknown as TextLayoutService;
  return { service, shapes: () => shapes };
}

function segment(service: TextLayoutService | undefined, text: string, context: string): LayoutTextSeg {
  return {
    text,
    fontSize: 10,
    ...(service ? { textLayoutService: service } : {}),
    textShapeRequest: {
      text,
      fontSizePt: 10,
      fonts: { ascii: 'Test' },
      substituteContext: { text: context, offset: context.indexOf(text) },
    },
  } as unknown as LayoutTextSeg;
}

function advances(seg: LayoutTextSeg, start = 0, end = seg.text.length, run = seg.text, scoped = false): {
  advances: readonly number[];
  work: number[];
} {
  const work: number[] = [];
  const clusters = contextualGraphemeAdvances(seg, start, end, 10,
    (text) => windowWidth(text, run, scoped),
    (units) => { work.push(units); });
  return { advances: clusters.map((cluster) => cluster.advancePx), work };
}

describe('WORD_CONDENSED_SPACING_ADVANCE_FLOOR contextual advance memo', () => {
  it('shapes each window of a run once, with unmemoised advances and work charges', () => {
    const text = 'WAVES of AVAIL';
    // Reference: the same windows measured without a shaping service.
    const expected = advances(segment(undefined, text, text));
    const { service, shapes } = countingService(false);
    const first = advances(segment(service, text, text));
    const afterFirst = shapes();
    // Fresh segments of the same run and a prefix probe shape nothing new.
    const again = advances(segment(service, text, text));
    const prefix = advances(segment(service, text, text), 0, 8);
    expect(first).toEqual(expected);
    expect(again).toEqual(expected);
    expect(prefix.advances).toEqual(expected.advances.slice(0, 8));
    expect(shapes()).toBe(afterFirst);
    // Overlapping windows inside one range are shaped once.
    expect(afterFirst).toBeLessThan(expected.work.length);
  });

  it('never reuses a window across runs whose context can change its shaping', () => {
    const { service } = countingService(true);
    const shortRun = 'AV';
    const longRun = 'AV '.repeat(40);
    const inShortRun = advances(segment(service, 'AV', shortRun));
    const inLongRun = advances(segment(service, 'AV', longRun));
    expect(inShortRun).toEqual(advances(segment(undefined, 'AV', shortRun), 0, 2, shortRun, true));
    expect(inLongRun).toEqual(advances(segment(undefined, 'AV', longRun), 0, 2, longRun, true));
    expect(inLongRun.advances).not.toEqual(inShortRun.advances);
  });

  it('does not memoise windows whose shape inputs are not primitives', () => {
    const { service, shapes } = countingService(false);
    const plain = segment(service, 'AB', 'AB');
    advances(plain);
    const warmed = shapes();
    const malformed = segment(service, 'AB', 'AB');
    (malformed.textShapeRequest as { fonts: unknown }).fonts = { ascii: ['Test'] };
    advances(malformed);
    expect(shapes()).toBeGreaterThan(warmed);
  });

  it('leaves sparse or malformed shape inputs to the service', () => {
    const { service, shapes } = countingService(false);
    const holes = (style: unknown): LayoutTextSeg => {
      const seg = segment(service, 'AB', 'AB');
      (seg.textShapeRequest as { style: unknown }).style = style;
      return seg;
    };
    // eslint-disable-next-line no-sparse-arrays
    advances(holes([1, , 2]));
    const warmed = shapes();
    // eslint-disable-next-line no-sparse-arrays
    advances(holes([1, 2, ,]));
    expect(shapes()).toBeGreaterThan(warmed);

    // Validation stays the service's: no key is read from a null font table.
    const validating = {
      ...service,
      shape: (request: Readonly<TextShapeRequest>) => {
        if (request.fontSizePt < 0) throw new RangeError('fontSizePt must be a finite non-negative number');
        return service.shape(request);
      },
    } as TextLayoutService;
    const nullFonts = segment(validating, 'AB', 'AB');
    (nullFonts.textShapeRequest as { fonts: unknown }).fonts = null;
    expect(() => contextualGraphemeAdvances(nullFonts, 0, 2, -1, () => 0, () => {})).toThrow(RangeError);
    const throwingFonts = segment(validating, 'AB', 'AB');
    (throwingFonts.textShapeRequest as { fonts: unknown }).fonts = Object.freeze({
      get ascii(): string { throw new TypeError('font accessor'); },
    });
    expect(() => contextualGraphemeAdvances(throwingFonts, 0, 2, -1, () => 0, () => {})).toThrow(RangeError);
  });
});

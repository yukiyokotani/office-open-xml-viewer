import { describe, expect, it } from 'vitest';
import type { ChartLabelBox } from '../types/chart.js';
import { mergeChartLabelBoxes } from './label-box.js';

const SHADOW = {
  color: '112233', alpha: 0.5, blur: 12_700, dist: 25_400, dir: 0,
};

describe('chart label box component precedence', () => {
  const linked = {
    fill: 'FFFFFF',
    style: { shadows: [SHADOW], effectAuthored: true },
  } satisfies ChartLabelBox;

  it('inherits a lower effect when the point authors only box paint', () => {
    const merged = mergeChartLabelBoxes({ fill: 'ABCDEF' }, linked);
    expect(merged?.fill).toBe('ABCDEF');
    expect(merged?.style).toBe(linked.style);
  });

  it.each([
    { effectAuthored: true },
    { effectUnsupported: true },
  ])('lets an authored point effect suppress the lower effect: %o', style => {
    const merged = mergeChartLabelBoxes({ fill: 'ABCDEF', style }, linked);
    expect(merged?.style).toBe(style);
  });
});

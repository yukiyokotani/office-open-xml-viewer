import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  MAX_HISTOGRAM_BINS,
  MAX_HISTOGRAM_INPUT_POINTS,
  planHistogramBins,
} from './histogram-binning.js';

describe('ChartEx histogram bin planning', () => {
  it('honors authored bin count and the closed interval side', () => {
    const values = [0, 1, 2, 3, 4];
    expect(planHistogramBins(values, { binCount: 2, intervalClosed: 'l' })).toMatchObject({
      kind: 'bins',
      counts: [2, 3],
    });
    expect(planHistogramBins(values, { binCount: 2, intervalClosed: 'r' })).toMatchObject({
      kind: 'bins',
      counts: [3, 2],
    });
  });

  it('assigns authored underflow and overflow boundaries consistently', () => {
    const values = [-1, 0, 1, 2, 3, 4, 5];
    expect(planHistogramBins(values, {
      binCount: 2,
      intervalClosed: 'l',
      underflow: 0,
      overflow: 4,
    })).toMatchObject({ kind: 'bins', counts: [1, 2, 2, 2] });
    expect(planHistogramBins(values, {
      binCount: 2,
      intervalClosed: 'r',
      underflow: 0,
      overflow: 4,
    })).toMatchObject({ kind: 'bins', counts: [2, 2, 2, 1] });
  });

  it('caps the final authored-size label at an explicit overflow boundary', () => {
    expect(planHistogramBins([0, 3, 4], {
      binSize: 3,
      intervalClosed: 'l',
      overflow: 4,
    })).toMatchObject({
      kind: 'bins',
      categories: ['[0, 3)', '[3, 4)', '≥ 4'],
      counts: [1, 1, 1],
    });
  });

  it('coarsens an extreme authored size into a bounded bin plan', () => {
    const result = planHistogramBins([0, 1_000_000_000_000], { binSize: 0.000_000_000_001 });
    expect(result.kind).toBe('bins');
    if (result.kind !== 'bins') return;
    expect(result.counts).toHaveLength(MAX_HISTOGRAM_BINS);
    expect(result.counts.reduce((sum, count) => sum + count, 0)).toBe(2);
  });

  it('keeps finite observations when their derived range overflows binary64', () => {
    expect(planHistogramBins([-1e308, 1e308], { binCount: 10 })).toMatchObject({
      kind: 'bins',
      counts: [2],
    });
  });

  it('keeps every observation when a derived bin width would underflow', () => {
    expect(planHistogramBins([0, Number.MIN_VALUE], { binCount: 2 })).toMatchObject({
      kind: 'bins',
      counts: [1, 1],
    });
  });

  it('rejects input beyond the parser cache ceiling before scanning it', () => {
    const oversized = new Array<number | null>(MAX_HISTOGRAM_INPUT_POINTS + 1);
    expect(planHistogramBins(oversized, {})).toEqual({ kind: 'tooManyInputPoints' });
  });
});

describe('ChartEx histogram automatic width (measured Office rule)', () => {
  const plan = (v: number[], o = {}, code?: string) => {
    const r = planHistogramBins(v, o, code);
    if (r.kind !== 'bins') throw new Error('expected bins');
    return r;
  };
  const even = Array.from({ length: 50 }, (_, i) => i * 3); // 0..147

  it('uses 3.5*sample sigma/n^(1/3) rounded to 2 significant digits', () => {
    // sigma ~ 43.7, n = 50 -> 41.55 -> 42; range 147 -> 4 bins.
    const r = plan(even, { intervalClosed: 'r' });
    expect(r.categories).toEqual(['[0, 42]', '(42, 84]', '(84, 126]', '(126, 168]']);
    expect(r.counts.reduce((a, b) => a + b, 0)).toBe(50);
  });

  it('uses the sample (n-1) sigma, not the population sigma', () => {
    // sample sigma of [0,10] = 7.07 -> width 3.5*7.07/1.26 = 19.64 -> 20;
    // population sigma would give 13.9 -> two bins.
    expect(plan([0, 10]).categories).toEqual(['[0, 20]']);
    // [1,2,3,4,100]: sample 89.28 -> 89 (population 79.9 -> 80).
    expect(plan([1, 2, 3, 4, 100], { intervalClosed: 'r' }).categories).toEqual(['[1, 90]', '(90, 179]']);
  });

  it('falls back to width 5 for a single point or zero deviation', () => {
    expect(plan([5]).categories).toEqual(['[5, 10]']);
    expect(plan([7, 7, 7]).categories).toEqual(['[7, 12]']);
    expect(plan([7, 7, 7]).counts).toEqual([3]);
  });

  it('keeps a maximum on a boundary in the last bin', () => {
    const r = plan([0, 10], { binSize: 5, intervalClosed: 'r' });
    expect(r.categories).toEqual(['[0, 5]', '(5, 10]']);
    expect(r.counts).toEqual([1, 1]);
    const l = plan([0, 10], { binSize: 5, intervalClosed: 'l' });
    expect(l.categories).toEqual(['[0, 5)', '[5, 10]']);
    expect(l.counts).toEqual([1, 1]);
  });

  it('labels r-closed and l-closed bins and avoids float noise', () => {
    const v = [1, 1, 2, 2, 2, 3, 3, 4, 4, 5, 6, 7];
    const r = plan(v, { intervalClosed: 'r' });
    expect(r.categories).toEqual(['[1, 3.9]', '(3.9, 6.8]', '(6.8, 9.7]']);
    expect(r.counts).toEqual([7, 4, 1]);
    const l = plan(v, { intervalClosed: 'l' });
    expect(l.categories).toEqual(['[1, 3.9)', '[3.9, 6.8)', '[6.8, 9.7]']);
    expect(l.counts).toEqual([7, 4, 1]);
  });

  it('computes sigma over all points and truncates at the overflow', () => {
    const r = plan(even, { intervalClosed: 'r', underflow: 30, overflow: 100 });
    // width from all 50 points is 42; edges 30, 72, 100 (truncated).
    expect(r.categories).toEqual(['≤ 30', '(30, 72]', '(72, 100]', '> 100']);
    expect(r.counts).toEqual([11, 14, 9, 16]);
  });

  it('honors binCount (range/count) and binSize', () => {
    expect(plan(even, { binCount: 7, intervalClosed: 'r' }).categories).toEqual([
      '[0, 21]', '(21, 42]', '(42, 63]', '(63, 84]', '(84, 105]', '(105, 126]', '(126, 147]',
    ]);
    expect(plan(even, { binSize: 60, intervalClosed: 'r' }).categories).toEqual([
      '[0, 60]', '(60, 120]', '(120, 180]',
    ]);
  });

  it('formats edges with the value format code', () => {
    expect(plan([0, 10], { binSize: 5 }, '0.00').categories).toEqual(['[0.00, 5.00)', '[5.00, 10.00]']);
  });

  it('keeps large integer edges distinct (15-digit and epoch-millisecond values)', () => {
    const big = plan([123456789012345, 123456789012399], { binSize: 10 });
    expect(big.categories).toEqual([
      '[123456789012345, 123456789012355)',
      '[123456789012355, 123456789012365)',
      '[123456789012365, 123456789012375)',
      '[123456789012375, 123456789012385)',
      '[123456789012385, 123456789012395)',
      '[123456789012395, 123456789012405]',
    ]);
    expect(big.counts).toEqual([1, 0, 0, 0, 0, 1]);
    const t0 = 1_700_000_000_000;
    const ms = plan([t0, t0 + 3_600_000, t0 + 7_200_000], { binSize: 3_600_000, intervalClosed: 'r' });
    expect(ms.categories).toEqual([
      '[1700000000000, 1700003600000]',
      '(1700003600000, 1700007200000]',
    ]);
    expect(ms.counts).toEqual([2, 1]);
  });

  it('classifies values by the decimal boundary, not by an epsilon', () => {
    // 3.0000000001 is strictly above the edge 3 and belongs to (3, 4].
    const r = plan([0, 3.0000000001, 5], { binSize: 1, intervalClosed: 'r' });
    expect(r.categories.slice(0, 5)).toEqual(['[0, 1]', '(1, 2]', '(2, 3]', '(3, 4]', '(4, 5]']);
    expect(r.counts).toEqual([1, 0, 0, 1, 1]);
    const l = plan([0, 2.9999999999, 5], { binSize: 1, intervalClosed: 'l' });
    expect(l.counts.slice(0, 4)).toEqual([1, 0, 1, 0]);
  });

  it('handles negative values and a range crossing zero', () => {
    const r = plan([-30, -21, -12, -3, 0, 6, 12, 18, 30], { intervalClosed: 'r' });
    expect(r.categories[0]).toMatch(/^\[-30, /);
    expect(r.counts.reduce((a, b) => a + b, 0)).toBe(9);
    expect(r.categories.every(label => !label.includes('e') && !label.includes('0000'))).toBe(true);
    const fixed = plan([-4, -2, 0, 2, 4], { binSize: 4, intervalClosed: 'r' });
    expect(fixed.categories).toEqual(['[-4, 0]', '(0, 4]']);
    expect(fixed.counts).toEqual([3, 2]);
  });

  it('shows decimal edges and applies a fixed value format', () => {
    const v = [0.21, 0.35, 0.4, 0.52, 0.66, 0.7, 0.83, 0.9, 1.05, 1.2];
    const auto = plan(v, { intervalClosed: 'r' });
    expect(auto.categories.every(label => !/\d{6,}/.test(label))).toBe(true);
    const fixed = plan(v, { intervalClosed: 'r' }, '0.000');
    expect(fixed.categories[0]).toMatch(/^\[0\.210, \d\.\d{3}\]$/);
    expect(fixed.categories[1]).toMatch(/^\(\d\.\d{3}, \d\.\d{3}\]$/);
    expect(fixed.counts).toEqual(auto.counts);
  });

  it('prefers the edge format code over the series format', () => {
    const r = planHistogramBins([0, 10], { binSize: 5, edgeFormatCode: '0.0' }, '0.000');
    if (r.kind !== 'bins') throw new Error('bins');
    expect(r.categories).toEqual(['[0.0, 5.0)', '[5.0, 10.0]']);
  });

  it('keeps underflow and overflow bins with a width computed over all points', () => {
    const v = [-40, -25, -10, 0, 5, 10, 15, 20, 30, 45, 60, 90];
    const r = plan(v, { intervalClosed: 'r', underflow: -10, overflow: 45 });
    expect(r.categories[0]).toBe('≤ -10');
    expect(r.categories[r.categories.length - 1]).toBe('> 45');
    expect(r.categories[1].startsWith('(-10, ')).toBe(true);
    expect(r.categories[r.categories.length - 2].endsWith(', 45]')).toBe(true);
    expect(r.counts[0]).toBe(3);
    expect(r.counts[r.counts.length - 1]).toBe(2);
    expect(r.counts.reduce((a, b) => a + b, 0)).toBe(12);
  });

  it('spans the range with binCount and binSize on negative data', () => {
    const v = [-12, -6, 0, 6, 12, 18, 24];
    expect(plan(v, { binCount: 4, intervalClosed: 'r' }).categories).toEqual([
      '[-12, -3]', '(-3, 6]', '(6, 15]', '(15, 24]',
    ]);
    expect(plan(v, { binSize: 20, intervalClosed: 'l' }).categories).toEqual([
      '[-12, 8)', '[8, 28]',
    ]);
  });
});

describe('ChartEx histogram General edge labels under a comma-decimal host locale', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('paints the decimal comma and keeps ", " between edges (Excel fr-FR control)', () => {
    vi.stubGlobal('navigator', { language: 'fr-FR', languages: ['fr-FR'] });
    const data = [0.1, 0.35, 0.6, 0.85, 1.1, 1.35, 1.6, 1.85, 2.1, 2.35];
    const result = planHistogramBins(data, { binSize: 0.5, intervalClosed: 'r' });
    expect(result.kind === 'bins' ? result.categories : []).toEqual([
      '[0,1, 0,6]', '(0,6, 1,1]', '(1,1, 1,6]', '(1,6, 2,1]', '(2,1, 2,6]',
    ]);
  });
});

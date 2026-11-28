import { describe, expect, it } from 'vitest';
import type { BorderSegment, TextDecorationLayout } from '../layout/types.js';
import type {
  CanvasPaintContext,
  PaintCanvas2D,
} from './types.js';
import { paintStrokeSegment } from './canvas-border.js';

function recordingContext(operations: unknown[]): PaintCanvas2D {
  return {
    globalAlpha: 1,
    fillStyle: '',
    strokeStyle: '',
    lineWidth: 1,
    font: '',
    textAlign: 'left',
    textBaseline: 'alphabetic',
    direction: 'ltr',
    letterSpacing: '0px',
    fontKerning: 'auto',
    fillRect(x, y, width, height) { operations.push(['fillRect', x, y, width, height]); },
    strokeRect() {},
    setLineDash(pattern) { operations.push(['dash', pattern]); },
    fillText() {},
    translate() {},
    rotate() {},
    scale() {},
    transform() {},
    drawImage() {},
    save() { operations.push(['save']); },
    restore() { operations.push(['restore']); },
    beginPath() {},
    rect() {},
    clip() {},
    moveTo() {},
    lineTo() {},
    stroke() { operations.push(['stroke', this.lineWidth]); },
    fill() {},
  };
}

const dashedSegment = (widthPt: number, dashPatternPt: readonly number[]): BorderSegment => ({
  edge: 'top',
  from: { xPt: 0, yPt: 4 },
  to: { xPt: 100, yPt: 4 },
  color: '#000000',
  widthPt,
  authoredStyle: 'dashed',
  style: 'dashed',
  dashPatternPt,
});

describe('retained Canvas border paint', () => {
  it('recomputes authored dash cadence when a minimum CSS width clamps a hairline', () => {
    const operations: unknown[] = [];
    const context: CanvasPaintContext = {
      ctx: recordingContext(operations),
      scale: 1,
      dpr: 1,
      resources: { paint() {} },
    };

    paintStrokeSegment(dashedSegment(0.25, [0.75, 0.5]), context, 0.5);

    expect(operations).toContainEqual(['dash', [1.5, 1]]);
    expect(operations).toContainEqual(['stroke', 0.5]);
  });

  it('keeps the retained dash cadence when the authored width exceeds the minimum', () => {
    const operations: unknown[] = [];
    const context: CanvasPaintContext = {
      ctx: recordingContext(operations),
      scale: 1,
      dpr: 1,
      resources: { paint() {} },
    };

    paintStrokeSegment(dashedSegment(1, [3, 2]), context, 0.5);

    expect(operations).toContainEqual(['dash', [3, 2]]);
    expect(operations).toContainEqual(['stroke', 1]);
  });

  it('bounds retained wavy paths with a bevel join instead of an unretained miter', () => {
    const operations: unknown[] = [];
    const ctx = recordingContext(operations);
    Object.defineProperty(ctx, 'lineJoin', {
      configurable: true,
      get: () => 'miter',
      set: (value) => operations.push(['lineJoin', value]),
    });
    const context: CanvasPaintContext = {
      ctx,
      scale: 1,
      dpr: 1,
      resources: { paint() {} },
    };
    const wavy: TextDecorationLayout = {
      kind: 'underline',
      from: { xPt: 0, yPt: 10 },
      to: { xPt: 8, yPt: 10 },
      color: '#000000',
      widthPt: 2,
      style: 'wavy',
      path: [
        { xPt: 0, yPt: 9 },
        { xPt: 2, yPt: 11 },
        { xPt: 4, yPt: 9 },
      ],
    };

    paintStrokeSegment(wavy, context);

    expect(operations).toContainEqual(['lineJoin', 'bevel']);
    expect(operations).toContainEqual(['save']);
    expect(operations).toContainEqual(['stroke', 2]);
    expect(operations).toContainEqual(['restore']);
    const operationIndex = (expected: unknown): number => operations.findIndex(
      (operation) => JSON.stringify(operation) === JSON.stringify(expected),
    );
    expect(operationIndex(['save']))
      .toBeLessThan(operationIndex(['lineJoin', 'bevel']));
    expect(operationIndex(['lineJoin', 'bevel']))
      .toBeLessThan(operationIndex(['stroke', 2]));
    expect(operationIndex(['stroke', 2]))
      .toBeLessThan(operationIndex(['restore']));
  });

  it('does not override the join for a two-point wavy segment', () => {
    const operations: unknown[] = [];
    const ctx = recordingContext(operations);
    Object.defineProperty(ctx, 'lineJoin', {
      configurable: true,
      get: () => 'miter',
      set: (value) => operations.push(['lineJoin', value]),
    });
    const context: CanvasPaintContext = {
      ctx,
      scale: 1,
      dpr: 1,
      resources: { paint() {} },
    };

    paintStrokeSegment({
      kind: 'underline',
      from: { xPt: 0, yPt: 10 },
      to: { xPt: 8, yPt: 10 },
      color: '#000000',
      widthPt: 2,
      style: 'wavy',
      path: [{ xPt: 0, yPt: 9 }, { xPt: 8, yPt: 11 }],
    }, context);

    expect(operations).not.toContainEqual(['lineJoin', 'bevel']);
    expect(operations).not.toContainEqual(['save']);
    expect(operations).not.toContainEqual(['restore']);
  });

  it('paints thin-thick compound borders as distinct unequal rails', () => {
    const operations: unknown[] = [];
    const context: CanvasPaintContext = {
      ctx: recordingContext(operations),
      scale: 1,
      dpr: 4,
      resources: { paint() {} },
    };

    paintStrokeSegment({
      edge: 'top', from: { xPt: 0, yPt: 10 }, to: { xPt: 100, yPt: 10 },
      color: '#000000', widthPt: 3, authoredStyle: 'thinThickSmallGap',
      style: 'compound', dashPatternPt: [],
    }, context);

    const fills = operations.filter((operation) =>
      Array.isArray(operation) && operation[0] === 'fillRect') as unknown[][];
    expect(fills).toHaveLength(2);
    expect(fills[0]?.[4]).toBeLessThan(fills[1]?.[4] as number);
    expect(operations).not.toContainEqual(['stroke', 3]);
  });
});

describe('paragraph wave borders (WORD_PARAGRAPH_BORDER_FLOW_RESERVATION)', () => {
  const tracePoints = (segment: BorderSegment) => {
    const points: Array<readonly [number, number]> = [];
    const widths: number[] = [];
    const ctx = recordingContext([]);
    ctx.moveTo = (x: number, y: number) => { points.push([x, y]); };
    ctx.lineTo = (x: number, y: number) => { points.push([x, y]); };
    ctx.stroke = function stroke() { widths.push(this.lineWidth); };
    paintStrokeSegment(segment, { ctx, scale: 1, dpr: 1, resources: { paint() {} } });
    return { points, widths };
  };
  const waveSegment = (from: { xPt: number; yPt: number }, to: { xPt: number; yPt: number },
    authoredStyle: 'wave' | 'doubleWave', widthPt: number, bandPt: number): BorderSegment => ({
    edge: from.yPt === to.yPt ? 'top' : 'left', from, to, color: '#000000', widthPt, authoredStyle, style: 'wavy',
    wave: { bandPt, rows: authoredStyle === 'wave' ? 1 : 2, hairlinePt: 0.375 },
  });

  it('fills the band with one hairline zigzag whatever the authored width', () => {
    for (const widthPt of [0.25, 6]) {
      const { points, widths } = tracePoints(waveSegment({ xPt: 0, yPt: 10 }, { xPt: 13, yPt: 10 }, 'wave', widthPt, 3));
      expect(widths).toEqual([0.375]);
      // A 3pt band about y = 10: vertices on its outer sides less half the
      // hairline, one band height apart, the last cut at the segment end.
      expect(points.map(([x]) => x)).toEqual([0, 3, 6, 9, 12, 13]);
      expect(points.map(([, y]) => y)).toEqual([8.6875, 11.3125, 8.6875, 11.3125, 8.6875, 8.6875 + 2.625 / 3]);
    }
  });

  it('stacks two zigzags in a double wave band and follows a vertical edge', () => {
    const { points } = tracePoints(waveSegment({ xPt: 20, yPt: 0 }, { xPt: 20, yPt: 5.25 }, 'doubleWave', 0.5, 5.25));
    const across = points.map(([x]) => x);
    expect(Math.min(...across) - 0.1875).toBeCloseTo(20 - 2.625, 10);
    expect(Math.max(...across) + 0.1875).toBeCloseTo(20 + 2.625, 10);
    // Each row is half the band: the outer row spans the left half only.
    expect(across.slice(0, 3)).toEqual([20 - 2.625 + 0.1875, 20 - 0.1875, 20 - 2.625 + 0.1875].map((x) => x));
    expect(points.map(([, y]) => y).slice(0, 3)).toEqual([0, 2.625, 5.25]);
  });
});

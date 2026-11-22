import { describe, expect, it } from 'vitest';
import { layoutLines, type LayoutTextSeg } from './line-layout.js';
import { createFontResolver } from './layout/font-service.js';
import {
  createTextLayoutService,
  type TextLayoutService,
  type TextShapeRequest,
} from './layout/text.js';

function measureContext(): CanvasRenderingContext2D {
  let font = '10px sans-serif';
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText(text: string) {
      return {
        width: [...text].length * 5,
        fontBoundingBoxAscent: 8,
        fontBoundingBoxDescent: 2,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

describe('line layout cluster acquisition', () => {
  it.each(['', ' '])('keeps justified fit acquisition bounded for an overlong word with tail %j', (tail) => {
    const text = 'a'.repeat(400) + tail;
    const layout = (compression: boolean) => {
      let measuredUtf16 = 0;
      const service = createTextLayoutService({
        fonts: createFontResolver([]),
        measurer: { fingerprint: 'overlong-word-acquisition', measure(request) {
          measuredUtf16 += request.text.length;
          return { advancePt: request.text.length * 5, ascentPt: 8, descentPt: 2 };
        } },
      });
      const segment: LayoutTextSeg = {
        text, bold: false, italic: false, underline: false, strikethrough: false,
        fontSize: 10, color: null, fontFamily: 'serif', vertAlign: null, measuredWidth: 0,
        textLayoutService: service,
        textShapeRequest: { text, fontSizePt: 10, fonts: { ascii: 'serif' }, measure: false },
      };
      const lines = layoutLines(measureContext(), [segment], 50, 0, 1, [], undefined, {}, 0,
        undefined, undefined, undefined, undefined, false, true, false,
        undefined, 'bounded', undefined, false, compression);
      return { measuredUtf16, lines: lines.map(line => line.segments.map(s => 'text' in s ? s.text : '').join('')) };
    };
    const previous = layout(false);
    const current = layout(true);
    expect(current.lines).toEqual(previous.lines);
    // Count shaping input, not wall time: a full grapheme pass over every
    // emergency suffix grows cubically and exceeds the aggregate path budget.
    expect(current.measuredUtf16).toBeLessThanOrEqual(previous.measuredUtf16 * 2);
  });

  it.each([
    { text: 'あいうえおかきくけこ', width: 16, wraps: true },
    { text: 'aa bb cc dd', width: 16, wraps: true },
    { text: 'aa bb', width: 60, wraps: false },
  ])('retains complete final clusters and face metrics for $text at $width', ({ text, width, wraps }) => {
    const requests: TextShapeRequest[] = [];
    const base = createTextLayoutService({
      fonts: createFontResolver([]),
      measurer: {
        fingerprint: 'line-layout-cluster-routing-v1',
        measure: (request) => ({
          advancePt: [...request.text].length * 5,
          ascentPt: 8,
          descentPt: 2,
        }),
      },
    });
    const textService: TextLayoutService = Object.freeze({
      fingerprint: base.fingerprint,
      localMetrics: base.localMetrics,
      resolve: (request: Parameters<TextLayoutService['resolve']>[0]) => base.resolve(request),
      shape: (request: Parameters<TextLayoutService['shape']>[0]) => {
        requests.push(request);
        return base.shape(request);
      },
    });
    const shapeRequest: TextShapeRequest = Object.freeze({
      text,
      fontSizePt: 10,
      fonts: { eastAsia: 'CJK Face' },
      measure: false,
    });
    const segment: LayoutTextSeg = {
      text,
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      fontSize: 10,
      color: null,
      fontFamily: 'CJK Face',
      vertAlign: null,
      measuredWidth: 0,
      textLayoutService: textService,
      textShapeRequest: shapeRequest,
    };

    const lines = layoutLines(measureContext(), [segment], width, 0, 1, [], undefined, {}, 0,
      undefined, undefined, undefined, undefined, false, text.includes(' '), false,
      undefined, 'bounded', undefined, false,
      text.includes(' ') ? { admissionOnly: true, lineEndIdeographicSpaceHangs: false } : false);
    const finalSegments = lines.flatMap((line) => line.segments)
      .filter((candidate): candidate is LayoutTextSeg => 'text' in candidate);
    const aggregateRequests = requests.filter((request) => request.clusterGeometry === false);
    const retainedRequests = requests.filter((request) => request.clusterGeometry === true);

    if (wraps) {
      expect(finalSegments.length).toBeGreaterThan(1);
      expect(aggregateRequests.some((request) => request.text !== text)).toBe(true);
    } else {
      expect(lines).toHaveLength(1);
      expect(finalSegments.map(segment => segment.text).join('')).toBe(text);
      expect(finalSegments[0]?.measuredWidth).toBe(25);
    }
    expect(retainedRequests.map((request) => request.text))
      .toEqual(finalSegments.map((candidate) => candidate.text));
    for (const candidate of finalSegments) {
      expect(candidate.selectedFaceFontBox).toEqual({ ascentPt: 8, descentPt: 2 });
      expect(candidate.selectedFaceInkBounds).toEqual({
        xMinPt: 0, xMaxPt: candidate.text.length * 5, ascentPt: 8, descentPt: 2,
      });
      expect(candidate.shapedClusters?.map((cluster) => cluster.range)).toEqual(
        [...candidate.text].map((_character, index) => ({ start: index, end: index + 1 })),
      );
    }
  });
});

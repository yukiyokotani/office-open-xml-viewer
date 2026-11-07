import { describe, expect, it } from 'vitest';
import { buildSegments, type LineLayoutEnvironment } from './line-layout.js';
import type { LayoutImageSeg } from './line-breaker/model.js';
import type { DocRun } from './types.js';

const image = (inlineEffectExtent?: [number, number, number, number]): DocRun => ({
  type: 'image', imagePath: 'word/media/p.png', mimeType: 'image/png',
  widthPt: 20, heightPt: 40, anchor: false,
  ...(inlineEffectExtent ? { inlineEffectExtent } : {}),
} as unknown as DocRun);

const imageSeg = (run: DocRun, environment: Partial<LineLayoutEnvironment> = {}): LayoutImageSeg => {
  const [seg] = buildSegments([run] as never, environment as LineLayoutEnvironment);
  if (!seg || !('imagePath' in seg)) throw new Error('expected an image segment');
  return seg;
};

describe('inline wp:effectExtent (ECMA-376 §20.4.2.6)', () => {
  it('widens the line box by the extent and keeps the authored object size', () => {
    const seg = imageSeg(image([2, 3, 5, 8]));
    expect([seg.widthPt, seg.heightPt]).toEqual([27, 51]);
    expect(seg.effectExtent).toEqual([2, 3, 5, 8]);
    expect(seg.objectSizePt).toEqual([20, 40]);
  });

  it('leaves an object without an extent unchanged', () => {
    const seg = imageSeg(image());
    expect([seg.widthPt, seg.heightPt]).toEqual([20, 40]);
    expect(seg.effectExtent).toBeUndefined();
  });

  it('clamps a negative extent box at zero but paints the authored size', () => {
    const seg = imageSeg(image([-30, -50, 0, 0]));
    expect([seg.widthPt, seg.heightPt]).toEqual([0, 0]);
    expect(seg.objectSizePt).toEqual([20, 40]);
  });

  it('keeps an inline shape on its own edges in a vertical section (not counter-turned)', () => {
    const [seg] = buildSegments([{
      type: 'shape', inline: true, widthPt: 20, heightPt: 40, inlineEffectExtent: [2, 3, 5, 8],
    }] as never, { uprightResourceOrientation: 'upright-physical' } as LineLayoutEnvironment);
    if (!seg || !('imagePath' in seg)) throw new Error('expected a shape segment');
    expect(seg.effectExtent).toEqual([2, 3, 5, 8]);
    expect([seg.widthPt, seg.heightPt]).toEqual([27, 51]);
  });

  it.each([
    ['upright-physical', [3, 5, 8, 2]],
    ['upright-physical-counter-clockwise', [8, 2, 3, 5]],
  ] as const)('maps physical edges onto the line axes of a %s section', (orientation, logical) => {
    const seg = imageSeg(image([2, 3, 5, 8]), { uprightResourceOrientation: orientation });
    expect(seg.effectExtent).toEqual(logical);
    expect([seg.widthPt, seg.heightPt]).toEqual([20 + logical[0] + logical[2], 40 + logical[1] + logical[3]]);
  });

  it('reserves the same box for a chart whether or not its resource is available', () => {
    const extent: [number, number, number, number] = [2, 3, 5, 8];
    const chart = imageSeg({ type: 'chart', widthPt: 20, heightPt: 40, anchor: false,
      inlineEffectExtent: extent } as unknown as DocRun);
    const unavailableChart = imageSeg({ type: 'unavailableDrawing', resourceKind: 'chart',
      widthPt: 20, heightPt: 40, inlineEffectExtent: extent } as unknown as DocRun);
    const unavailableImage = imageSeg({ type: 'unavailableDrawing', resourceKind: 'image',
      widthPt: 20, heightPt: 40, inlineEffectExtent: extent } as unknown as DocRun);
    expect([unavailableChart.widthPt, unavailableChart.heightPt]).toEqual([chart.widthPt, chart.heightPt]);
    expect(unavailableChart.effectExtent).toBeUndefined();
    expect([unavailableImage.widthPt, unavailableImage.heightPt]).toEqual([27, 51]);
  });
});

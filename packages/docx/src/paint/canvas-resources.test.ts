import { describe, expect, it } from 'vitest';
import type { SectionLayoutContext } from '../layout-context.js';
import { createPaintResourceRegistry } from '../layout/paint-resources.js';
import type { DocumentLayout, InlineResourceKind, LayoutRect } from '../layout/types.js';
import { createPaintResourceSession } from './resource-session.js';
import {
  createCanvasPaintResourcePainter,
  paintLayoutPage,
} from './canvas-page.js';

function resourceLayout(
  resourceKey: string,
  resourceKind: InlineResourceKind,
  bounds: LayoutRect,
): DocumentLayout {
  return {
    pages: [{
      pageIndex: 0,
      geometry: {
        xPt: 0, yPt: 0, widthPt: 100, heightPt: 200,
        contentTopPt: 10, contentBottomPt: 190,
      },
      flowDomains: [{
        id: 'body', kind: 'body', bounds: { xPt: 10, yPt: 10, widthPt: 80, heightPt: 180 },
      }],
      section: {} as SectionLayoutContext,
      sectionOccurrenceId: 'section:0', parityBlank: false, bookmarkStarts: [],
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:0' },
      sectionRegions: [{
        id: 'region:0', sectionOccurrenceId: 'section:0',
        coordinateSpace: { writingMode: 'horizontal-tb', logicalToPhysical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 } },
        blockStartPt: 10, blockEndPt: 190, flowDomainIds: ['body'],
        section: {} as SectionLayoutContext,
      }],
      layers: {
        paintOrder: [{ layer: 'body', nodeId: 'drawing-1', coordinateSpace: 'logical-body-points', logicalBlock: { blockStartPt: bounds.yPt, blockExtentPt: bounds.heightPt } }],
        background: [], behindText: [], header: [], notes: [], front: [], footer: [],
        body: [{
          kind: 'drawing',
          id: 'drawing-1',
          source: { story: 'body', storyInstance: 'body', path: [0] },
          flowBounds: bounds,
          inkBounds: bounds,
          advancePt: bounds.heightPt,
          ordinaryFlow: true,
          flowDomainId: 'body',
          commands: [{ kind: 'resource', resourceKey, resourceKind, rect: bounds }],
        }],
      },
      readingOrder: ['drawing-1'],
    }],
    diagnostics: [],
  };
}

function canvasTarget() {
  const ctx = {
    save() {}, restore() {}, setTransform() {}, transform() {}, clearRect() {},
  } as unknown as CanvasRenderingContext2D;
  const target = {
    width: 0, height: 0, getContext: () => ctx,
  } as unknown as HTMLCanvasElement;
  return { ctx, target };
}

describe('canvas paint resource adapter', () => {
  it('resolves a typed descriptor and per-render opaque handle before painting', async () => {
    const resourceKey = 'image:body:0';
    const bounds = { xPt: 10, yPt: 20, widthPt: 30, heightPt: 40 };
    const registry = createPaintResourceRegistry([{
      kind: 'image', resourceKey, partPath: 'word/media/image.png', mimeType: 'image/png',
      intrinsicSize: { widthPt: 30, heightPt: 40 },
    }]);
    const imageHandle = { imageBitmap: true };
    const session = createPaintResourceSession(registry, [{
      kind: 'image', resourceKey, handle: imageHandle,
    }]);
    const calls: unknown[] = [];
    const painter = createCanvasPaintResourcePainter(session, {
      image(resource, paintedBounds, ctx) {
        calls.push({ resource, paintedBounds, ctx });
      },
      chart() { throw new Error('unexpected chart'); },
      math() { throw new Error('unexpected math'); },
      'picture-bullet'() { throw new Error('unexpected picture bullet'); },
    });
    const { ctx, target } = canvasTarget();

    await paintLayoutPage(
      resourceLayout(resourceKey, 'image', bounds),
      0,
      target,
      { scale: 1, dpr: 1 },
      painter,
    );

    expect(calls).toEqual([{
      resource: {
        descriptor: registry.resolve(resourceKey, 'image'),
        handle: imageHandle,
      },
      paintedBounds: bounds,
      ctx,
    }]);
  });

  it('hard-fails when a retained resource has no bound painter', async () => {
    const { target } = canvasTarget();

    await expect(paintLayoutPage(
      resourceLayout('image:missing', 'image', { xPt: 0, yPt: 0, widthPt: 1, heightPt: 1 }),
      0,
      target,
      { scale: 1, dpr: 1 },
    )).rejects.toThrow(/Missing retained resource painter.*image:missing.*expected image/i);
  });

  it('hard-fails for missing handles and kind mismatches before the handler runs', () => {
    const registry = createPaintResourceRegistry([{
      kind: 'chart', resourceKey: 'chart:body:0', intrinsicSize: { widthPt: 1, heightPt: 1 },
      model: {} as never,
    }]);
    const session = createPaintResourceSession(registry, []);
    const painter = createCanvasPaintResourcePainter(session, {
      image() { throw new Error('handler must not run'); },
      chart() { throw new Error('handler must not run'); },
      math() { throw new Error('handler must not run'); },
      'picture-bullet'() { throw new Error('handler must not run'); },
    });
    const ctx = {} as CanvasRenderingContext2D;
    const bounds = { xPt: 0, yPt: 0, widthPt: 1, heightPt: 1 };

    expect(() => painter.paint('chart:body:0', 'chart', bounds, ctx))
      .toThrow(/Missing paint resource handle.*chart:body:0/);
    expect(() => painter.paint('chart:body:0', 'image', bounds, ctx))
      .toThrow(/kind mismatch.*expected image.*chart/i);
  });

});

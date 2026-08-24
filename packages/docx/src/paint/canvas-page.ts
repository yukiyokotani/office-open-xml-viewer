import type {
  DocumentLayout,
  LayoutPage,
  LayoutRect,
  PagePaintDrawingEntry,
  PagePaintEntry,
  PagePaintFrame,
  PaintNode,
  PaintResourceKind,
} from '../layout/types.js';
import { rasterizeColumnSeparator } from './column-separator-raster.js';
import { paintDrawingLayout } from './canvas-drawing.js';
import {
  paintDrawingWithOwnedTextBoxes,
  paintParagraphLayout,
} from './canvas-text.js';
import { paintTableLayout } from './canvas-table.js';
import { paintStrokeSegment } from './canvas-border.js';
import { applyCanvasTransform } from './canvas-transform.js';
import { canvasPaintFrame } from './deferred-paint-frame.js';
import { composeAffine, scaleAffine } from './affine.js';
import { paintPageBorderLayout } from './page-border.js';
import type { PaintResourceSession } from './resource-session.js';
import type {
  CanvasPaintContext,
  CanvasPaintResourceHandlers,
  CanvasPaintResourcePainter,
  PaintCanvas2D,
  PaintPageOptions,
} from './types.js';

const missingResourcePainter: CanvasPaintResourcePainter = Object.freeze({
  paint(resourceKey: string, kind: PaintResourceKind): never {
    throw new Error(
      `Missing retained resource painter for ${resourceKey}: expected ${kind}`,
    );
  },
});

export function createCanvasPaintResourcePainter(
  session: PaintResourceSession,
  handlers: CanvasPaintResourceHandlers,
): CanvasPaintResourcePainter {
  return Object.freeze({
    paint(
      resourceKey: string,
      kind: PaintResourceKind,
      bounds: LayoutRect,
      ctx: PaintCanvas2D,
    ): void {
      switch (kind) {
        case 'image':
          handlers.image(session.resolve(resourceKey, kind), bounds, ctx);
          return;
        case 'chart':
          handlers.chart(session.resolve(resourceKey, kind), bounds, ctx);
          return;
        case 'math':
          handlers.math(session.resolve(resourceKey, kind), bounds, ctx);
          return;
        case 'picture-bullet':
          handlers['picture-bullet'](session.resolve(resourceKey, kind), bounds, ctx);
          return;
        default: {
          const exhaustive: never = kind;
          throw new Error(`Unknown retained resource kind: ${String(exhaustive)}`);
        }
      }
    },
  });
}

function paintNode(node: PaintNode, context: CanvasPaintContext): void {
  switch (node.kind) {
    case 'drawing':
      paintDrawingLayout(node, context);
      return;
    case 'paragraph':
      paintParagraphLayout(node, context);
      return;
    case 'table':
      paintTableLayout(node, context, node.resolvedFloatingTables ?? []);
      return;
    case 'note': {
      node.separator.forEach((segment) => paintStrokeSegment(segment, context));
      const paintStory = () => node.story.blocks.forEach((block) => paintNode(block, context));
      if (!node.story.clipBounds) {
        paintStory();
        return;
      }
      const clip = node.story.clipBounds;
      context.ctx.save();
      try {
        context.ctx.beginPath();
        context.ctx.rect(clip.xPt, clip.yPt, clip.widthPt, clip.heightPt);
        context.ctx.clip();
        paintStory();
      } finally {
        context.ctx.restore();
      }
      return;
    }
    case 'textbox':
      throw new Error(`Unsupported page paint node kind: ${node.kind}`);
    default: {
      const exhaustive: never = node;
      throw new Error(`Unknown page paint node kind: ${String(exhaustive)}`);
    }
  }
}

function paintColumnSeparators(page: LayoutPage, context: CanvasPaintContext): void {
  const segments = page.columnSeparators;
  if (segments.length === 0) return;
  const { ctx } = context;
  ctx.save();
  ctx.strokeStyle = '#000000';
  for (const segment of segments) {
    const raster = rasterizeColumnSeparator(segment, context.scale, context.dpr);
    ctx.lineWidth = raster.widthPt;
    ctx.beginPath();
    ctx.moveTo(raster.segment.start.xPt, raster.segment.start.yPt);
    ctx.lineTo(raster.segment.end.xPt, raster.segment.end.yPt);
    ctx.stroke();
  }
  ctx.restore();
}

function paintInEntryRegion(
  entry: PagePaintEntry,
  context: CanvasPaintContext,
  regionByDomain: ReadonlyMap<string, LayoutPage['sectionRegions'][number]>,
  paint: (entryContext: CanvasPaintContext) => void,
): void {
  const region = regionByDomain.get(entry.flowDomainId);
  const matrix = entry.coordinateSpace === 'upright-physical'
    ? undefined
    : region?.coordinateSpace.logicalToPhysical;
  const frame = canvasPaintFrame(context.ctx, () => {
    if (matrix && (
      matrix.a !== 1 || matrix.b !== 0 || matrix.c !== 0
      || matrix.d !== 1 || matrix.e !== 0 || matrix.f !== 0
    )) applyCanvasTransform(context.ctx, matrix);
  });
  const entryContext: CanvasPaintContext = {
    ...context,
    ...(matrix ? {
      pointToCss: {
        a: matrix.a * context.scale,
        b: matrix.b * context.scale,
        c: matrix.c * context.scale,
        d: matrix.d * context.scale,
        e: matrix.e * context.scale,
        f: matrix.f * context.scale,
      },
    } : {}),
  };
  frame(() => paint(entryContext))();
}

function applyPaintFrame(frame: PagePaintFrame, ctx: PaintCanvas2D): void {
  if (frame.kind === 'transform') {
    const transform = (ctx as PaintCanvas2D & {
      transform?: (a: number, b: number, c: number, d: number, e: number, f: number) => void;
    }).transform;
    if (transform) {
      transform.call(
        ctx,
        frame.transform.a,
        frame.transform.b,
        frame.transform.c,
        frame.transform.d,
        frame.transform.e,
        frame.transform.f,
      );
    } else if (
      frame.transform.a === 1
      && frame.transform.b === 0
      && frame.transform.c === 0
      && frame.transform.d === 1
    ) {
      ctx.translate(frame.transform.e, frame.transform.f);
    } else {
      throw new Error('Canvas context cannot apply the retained page paint transform');
    }
    return;
  }
  ctx.beginPath();
  ctx.rect(
    frame.clip.xPt,
    frame.clip.yPt,
    frame.clip.widthPt,
    frame.clip.heightPt,
  );
  ctx.clip();
}

function paintDrawingEntry(
  entry: PagePaintDrawingEntry,
  context: CanvasPaintContext,
): void {
  let pointToCss = context.pointToCss ?? scaleAffine(context.scale);
  for (const frame of entry.frames) {
    if (frame.kind === 'transform') pointToCss = composeAffine(pointToCss, frame.transform);
  }
  const drawingContext: CanvasPaintContext = {
    ...context,
    pointToCss,
    layoutTranslationPt: entry.layoutTranslationPt,
    omitAnchoredDrawings: false,
  };
  let enteredFrames = 0;
  try {
    for (const retainedFrame of entry.frames) {
      context.ctx.save();
      enteredFrames += 1;
      applyPaintFrame(retainedFrame, context.ctx);
    }
    paintDrawingWithOwnedTextBoxes(entry.node, entry.textBoxes, drawingContext);
  } finally {
    while (enteredFrames > 0) {
      context.ctx.restore();
      enteredFrames -= 1;
    }
  }
}

/** Paint a completed page into an already initialized point-space surface. */
export function paintLayoutPageContent(
  page: LayoutPage,
  context: CanvasPaintContext,
): void {
  const regionByDomain = new Map(page.sectionRegions.flatMap((region) => (
    region.flowDomainIds.map((domainId) => [domainId, region] as const)
  )));
  const regionById = new Map(page.sectionRegions.map((region) => [region.id, region]));
  for (const domain of page.flowDomains) {
    if (domain.kind === 'footnote' || domain.kind === 'endnote') {
      const storyRegion = domain.sectionRegionId
        ? regionById.get(domain.sectionRegionId)
        : page.sectionRegions[0];
      if (!storyRegion) {
        throw new Error(
          `${domain.id} references missing page story region ${domain.sectionRegionId ?? '<default>'}`,
        );
      }
      regionByDomain.set(domain.id, storyRegion);
    }
  }
  const entries = page.layers.paintOrder;
  const firstNonLeadingEntry = entries.findIndex((entry) => (
    entry.sourceLayer !== 'background'
      && entry.sourceLayer !== 'behindText'
      && entry.sourceLayer !== 'header'
  ));
  const decorationIndex = firstNonLeadingEntry === -1 ? entries.length : firstNonLeadingEntry;
  const paintEntries = (retainedEntries: readonly PagePaintEntry[]): void => {
    for (const entry of retainedEntries) {
      paintInEntryRegion(entry, context, regionByDomain, (entryContext) => {
        if (entry.kind === 'drawing') {
          paintDrawingEntry(entry, entryContext);
        } else {
          paintNode(entry.node, {
            ...entryContext,
            omitAnchoredDrawings: entry.omitAnchoredDrawings,
          });
        }
      });
    }
  };
  if (page.pageBorder?.zOrder === 'back') {
    paintPageBorderLayout(page.pageBorder, context);
  }
  paintEntries(entries.slice(0, decorationIndex));
  paintColumnSeparators(page, context);
  paintEntries(entries.slice(decorationIndex));
  if (page.pageBorder?.zOrder !== 'back' && page.pageBorder) {
    paintPageBorderLayout(page.pageBorder, context);
  }
}

export async function paintLayoutPage(
  layout: DocumentLayout,
  pageIndex: number,
  target: HTMLCanvasElement | OffscreenCanvas,
  options: PaintPageOptions,
  resources: CanvasPaintResourcePainter = missingResourcePainter,
): Promise<void> {
  const page = layout.pages[pageIndex];
  if (!page) throw new RangeError(`Page ${pageIndex} is outside the layout`);
  const ctx = target.getContext('2d') as
    | CanvasRenderingContext2D
    | OffscreenCanvasRenderingContext2D
    | null;
  if (!ctx) throw new Error('Canvas 2D context is unavailable');

  const pixelScale = options.scale * options.dpr;
  target.width = Math.ceil(page.geometry.widthPt * pixelScale);
  target.height = Math.ceil(page.geometry.heightPt * pixelScale);
  ctx.save();
  try {
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, target.width, target.height);
    ctx.setTransform(pixelScale, 0, 0, pixelScale, 0, 0);
    paintLayoutPageContent(page, {
      ctx, scale: options.scale, dpr: options.dpr, resources,
    });
  } finally {
    ctx.restore();
  }
}

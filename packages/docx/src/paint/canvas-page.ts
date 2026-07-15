import type {
  DocumentLayout,
  LayoutRect,
  Matrix2DData,
  PaintNode,
  PaintResourceKind,
  ResolvedFloatingTablePlacementLayout,
  TableLayout,
} from '../layout/types.js';
import { composeAffine, inverseAffine, scaleAffine } from './affine.js';
import { orderedPagePaintEntries } from '../layout/page-graph.js';
import { paintDrawingLayout } from './canvas-drawing.js';
import { paintParagraphLayout } from './canvas-text.js';
import { paintTableLayout } from './canvas-table.js';
import type { PaintResourceSession } from './resource-session.js';
import type {
  CanvasPaintContext,
  CanvasPaintResourceHandlers,
  CanvasPaintResourcePainter,
  PaintCanvas2D,
  PaintPageOptions,
} from './types.js';
import type { TextRunPaintInfo } from './text-run-info.js';

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
  if (node.kind === 'drawing') paintDrawingLayout(node, context);
  else if (node.kind === 'paragraph') paintParagraphLayout(node, context);
  else if (node.kind === 'table') {
    const retained = node as TableLayout & {
      readonly resolvedFloatingTables?: readonly ResolvedFloatingTablePlacementLayout[];
    };
    paintTableLayout(node, context, retained.resolvedFloatingTables ?? []);
  }
}

function logicalRunCallback(
  callback: ((run: TextRunPaintInfo) => void) | undefined,
  matrix: Matrix2DData | undefined,
  scale: number,
): ((run: TextRunPaintInfo) => void) | undefined {
  if (!callback || !matrix) return callback;
  if (matrix.a === 1 && matrix.b === 0 && matrix.c === 0
    && matrix.d === 1 && matrix.e === 0 && matrix.f === 0) return callback;
  const orientation = matrix.a === 0 && matrix.b === 1
    && matrix.c === -1 && matrix.d === 0
    ? 'rotate(90deg)'
    : `matrix(${matrix.a}, ${matrix.b}, ${matrix.c}, ${matrix.d}, 0, 0)`;
  return (run) => callback({
    ...run,
    x: matrix.a * run.x + matrix.c * run.y + matrix.e * scale,
    y: matrix.b * run.x + matrix.d * run.y + matrix.f * scale,
    transform: run.transform ? `${orientation} ${run.transform}` : orientation,
  });
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
  const entries = orderedPagePaintEntries(page);
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
    const regionByDomain = new Map(page.sectionRegions.flatMap((region) => (
      region.flowDomainIds.map((domainId) => [domainId, region] as const)
    )));
    for (const entry of entries) {
      const region = regionByDomain.get(entry.node.flowDomainId);
      const logicalToPhysical = entry.coordinateSpace === 'logical-body-points'
        ? region?.coordinateSpace.logicalToPhysical
        : undefined;
      if (entry.coordinateSpace === 'logical-body-points' && !logicalToPhysical) {
        throw new Error(`Missing logical region transform for ${entry.node.id}`);
      }
      ctx.save();
      try {
        if (logicalToPhysical && (
          logicalToPhysical.a !== 1 || logicalToPhysical.b !== 0
          || logicalToPhysical.c !== 0 || logicalToPhysical.d !== 1
          || logicalToPhysical.e !== 0 || logicalToPhysical.f !== 0
        )) {
          ctx.transform(
            logicalToPhysical.a,
            logicalToPhysical.b,
            logicalToPhysical.c,
            logicalToPhysical.d,
            logicalToPhysical.e,
            logicalToPhysical.f,
          );
        }
        paintNode(entry.node, {
          ctx,
          scale: options.scale,
          dpr: options.dpr,
          resources,
          pointToCss: logicalToPhysical
            ? composeAffine(scaleAffine(options.scale), logicalToPhysical)
            : scaleAffine(options.scale),
          ...(logicalToPhysical ? { pageToLocal: inverseAffine(logicalToPhysical) ?? undefined } : {}),
          textRunTransform: { translateXPt: 0, translateYPt: 0, scale: options.scale },
          ...(options.onTextRun ? {
            onTextRun: logicalRunCallback(options.onTextRun, logicalToPhysical, options.scale),
            onPhysicalTextRun: options.onTextRun,
          } : {}),
        });
      } finally {
        ctx.restore();
      }
    }
  } finally {
    ctx.restore();
  }
}

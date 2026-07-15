import type {
  DocumentLayout,
  LayoutRect,
  PaintNode,
  PaintResourceKind,
} from '../layout/types.js';
import { orderedPagePaintNodes } from '../layout/page-graph.js';
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
  else if (node.kind === 'table') paintTableLayout(node, context);
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
  const nodes = orderedPagePaintNodes(page);
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
    for (const node of nodes) {
      paintNode(node, { ctx, scale: options.scale, dpr: options.dpr, resources });
    }
  } finally {
    ctx.restore();
  }
}

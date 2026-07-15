import {
  autoContrastColor,
  canvasFontString,
} from '@silurus/ooxml-core';
import type { ParagraphLayout, TextBoxLayout } from '../layout/types.js';
import type { CanvasPaintContext } from './types.js';
import { paintDrawingLayout } from './canvas-drawing.js';
import { paintStrokeSegment } from './canvas-border.js';
import {
  composeAffine,
  quarterTurnAffine,
  scaleAffine,
  translationAffine,
} from './affine.js';
import { textRunPaintInfo } from './text-run-info.js';

function validateTextSlices(placement: import('../layout/types.js').TextPlacement): void {
  if (placement.text.length !== placement.range.end - placement.range.start) {
    throw new Error('UTF-16 text range is inconsistent');
  }
  if (placement.clusters.length === 0) {
    throw new Error('Retained glyph slices are incomplete');
  }
  let cursor = placement.range.start;
  for (const cluster of placement.clusters) {
    const { advancePt, offset, range } = cluster;
    if (
      !Number.isFinite(advancePt) || advancePt < 0
      || !Number.isFinite(offset.xPt) || !Number.isFinite(offset.yPt)
      || range.start !== cursor || range.end <= range.start
      || range.end > placement.range.end
    ) {
      throw new Error('Retained glyph slices are incomplete');
    }
    cursor = range.end;
  }
  if (cursor !== placement.range.end) throw new Error('Retained glyph slices are incomplete');
  if (placement.paintOps.length === 0) throw new Error('Retained glyph slices are incomplete');
  let previousEnd = placement.range.start;
  for (const op of placement.paintOps) {
    const invalidTextMapping = op.sourceMapping !== 'kashida'
      && op.text.length !== op.range.end - op.range.start;
    const invalidGeometry = !Number.isFinite(op.offset.xPt) || !Number.isFinite(op.offset.yPt)
      || !Number.isFinite(op.letterSpacingPt)
      || !Number.isFinite(op.scaleX) || op.scaleX <= 0;
    const invalidRange = op.range.start !== previousEnd || op.range.end <= op.range.start
      || op.range.end > placement.range.end;
    if (invalidTextMapping || invalidGeometry || invalidRange) {
      throw new Error(
        `Retained glyph slices are incomplete (${invalidTextMapping ? 'text' : invalidGeometry ? 'geometry' : `range ${previousEnd}:${op.range.start}-${op.range.end}/${placement.range.end}`})`,
      );
    }
    previousEnd = op.range.end;
  }
  const trailing = placement.text.slice(previousEnd - placement.range.start);
  if (trailing !== '' && !/^\s+$/u.test(trailing)) {
    throw new Error('Retained glyph slices are incomplete');
  }
}

function resolvedTextColor(
  color: import('../layout/types.js').TextColorPolicy,
  context: CanvasPaintContext,
): string {
  if (color.kind === 'explicit') return color.color;
  if (color.kind === 'auto') return autoContrastColor(color.background ?? '#FFFFFF');
  return context.defaultTextColor ?? '#000000';
}

function textColor(
  placement: import('../layout/types.js').TextPlacement,
  context: CanvasPaintContext,
): string {
  return resolvedTextColor(placement.color, context);
}

function paintRetainedGlyph(
  operation: import('../layout/types.js').RetainedGlyphPaintOperation,
  context: CanvasPaintContext,
  upright = false,
): void {
  const { ctx } = context;
  ctx.fillStyle = resolvedTextColor(operation.color, context);
  ctx.font = canvasFontString(
    operation.fontRoute,
    operation.fontSizePt,
    operation.fontWeight,
    operation.fontStyle,
  );
  if (upright) {
    ctx.save();
    ctx.translate(operation.origin.xPt, operation.origin.yPt);
    ctx.rotate(-Math.PI / 2);
    ctx.fillText(operation.text, 0, 0);
    ctx.restore();
  } else {
    ctx.fillText(operation.text, operation.origin.xPt, operation.origin.yPt);
  }
}

function paintRetainedMarkPath(
  path: import('../layout/types.js').RetainedMarkPath,
  context: CanvasPaintContext,
): void {
  const { ctx } = context;
  ctx.beginPath();
  if (path.points.length > 0) {
    const first = path.points[0]!;
    ctx.moveTo(first.xPt, first.yPt);
    for (const point of path.points.slice(1)) ctx.lineTo(point.xPt, point.yPt);
  }
  if (path.stroke !== null) {
    ctx.strokeStyle = path.stroke;
    ctx.lineWidth = path.strokeWidthPt;
    ctx.stroke();
  }
  if (path.fill !== null) {
    ctx.fillStyle = path.fill;
    ctx.fill();
  }
}

/** Paints only retained point geometry. Text acquisition and measurement are not
 * available through this contract, so zoom cannot alter line partitioning. */
export function paintParagraphLayout(node: ParagraphLayout, context: CanvasPaintContext): void {
  const { ctx } = context;
  if (node.clipBounds) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(
      node.clipBounds.xPt,
      node.clipBounds.yPt,
      node.clipBounds.widthPt,
      node.clipBounds.heightPt,
    );
    ctx.clip();
  }
  try {
  const textBoxesById = new Map(node.textBoxes.map((textBox) => [textBox.id, textBox]));
  const ownedTextBoxIds = new Set(node.drawings.flatMap((drawing) => drawing.textBoxIds ?? []));
  const textBoxesFor = (drawing: import('../layout/types.js').DrawingLayout) =>
    (drawing.textBoxIds ?? []).flatMap((id) => {
      const textBox = textBoxesById.get(id);
      return textBox ? [textBox] : [];
    });
  const paintDrawingWithTextBoxes = (drawing: import('../layout/types.js').DrawingLayout): void => {
    const pageOwned = drawing.anchorLayer?.horizontalOwnership === 'page'
      && drawing.anchorLayer?.verticalOwnership === 'page';
    if (pageOwned && context.pageToLocal) {
      context.ctx.save();
      context.ctx.transform(
        context.pageToLocal.a,
        context.pageToLocal.b,
        context.pageToLocal.c,
        context.pageToLocal.d,
        context.pageToLocal.e,
        context.pageToLocal.f,
      );
      try {
        const physicalContext = {
          ...context,
          pointToCss: scaleAffine(context.scale),
          pageToLocal: undefined,
          onTextRun: context.onPhysicalTextRun,
        };
        paintDrawingLayout(drawing, physicalContext);
        for (const textBox of textBoxesFor(drawing)) {
          paintTextBoxLayout(textBox, physicalContext);
        }
      } finally {
        context.ctx.restore();
      }
      return;
    }
    const translation = context.layoutTranslationPt;
    const undoX = drawing.anchorLayer?.horizontalOwnership === 'page'
      ? -(translation?.xPt ?? 0) : 0;
    const undoY = drawing.anchorLayer?.verticalOwnership === 'page'
      ? -(translation?.yPt ?? 0) : 0;
    if (undoX !== 0 || undoY !== 0) {
      context.ctx.save();
      context.ctx.translate(undoX, undoY);
    }
    try {
      paintDrawingLayout(drawing, context);
      for (const textBox of textBoxesFor(drawing)) paintTextBoxLayout(textBox, context);
    } finally {
      if (undoX !== 0 || undoY !== 0) context.ctx.restore();
    }
  };
  const behind = node.drawings
    .filter((drawing) => drawing.anchorLayer?.behindDoc === true)
    .sort((a, b) => a.anchorLayer!.relativeHeight - b.anchorLayer!.relativeHeight
      || a.anchorLayer!.sourceOrder - b.anchorLayer!.sourceOrder);
  for (const drawing of behind) paintDrawingWithTextBoxes(drawing);
  for (const retained of node.lineNumbers ?? []) {
    for (const operation of retained.paintOps) {
      ctx.fillStyle = operation.color;
      ctx.font = operation.font;
      ctx.textAlign = operation.textAlign;
      ctx.textBaseline = 'alphabetic';
      ctx.fillText(operation.text, operation.origin.xPt, operation.origin.yPt);
    }
  }
  if (node.shading) {
    ctx.fillStyle = node.shading.color;
    ctx.fillRect(
      node.inkBounds.xPt,
      node.inkBounds.yPt,
      node.inkBounds.widthPt,
      node.inkBounds.heightPt,
    );
  }
  for (const line of node.lines) {
    for (const placement of line.placements) {
      if (placement.kind === 'resource') {
        if (!context.resources) {
          throw new Error(`Missing retained resource painter for ${placement.resourceKey}`);
        }
        if (context.textBoxVerticalMode) {
          const rotation = context.textBoxVerticalMode === 'vert270' ? Math.PI / 2 : -Math.PI / 2;
          ctx.save();
          ctx.translate(
            placement.bounds.xPt + placement.bounds.widthPt / 2,
            placement.bounds.yPt + placement.bounds.heightPt / 2,
          );
          ctx.rotate(rotation);
          context.resources.paint(
            placement.resourceKey,
            placement.resourceKind,
            {
              xPt: -placement.bounds.heightPt / 2,
              yPt: -placement.bounds.widthPt / 2,
              widthPt: placement.bounds.heightPt,
              heightPt: placement.bounds.widthPt,
            },
            ctx,
          );
          ctx.restore();
        } else {
          context.resources.paint(
            placement.resourceKey,
            placement.resourceKind,
            placement.bounds,
            ctx,
          );
        }
        continue;
      }
      if (placement.kind === 'tab') {
        if (placement.leader !== 'none') {
          if (!placement.leaderGlyphs) {
            throw new Error('Retained tab leader geometry is missing');
          }
          for (const operation of placement.leaderGlyphs) paintRetainedGlyph(operation, context);
        }
        continue;
      }
      if (placement.kind !== 'text') continue;
      validateTextSlices(placement);
      if (placement.unsupportedGeometry?.length) {
        throw new Error(
          `Unsupported retained typography geometry: ${placement.unsupportedGeometry.join(', ')}`,
        );
      }
      if (placement.highlightFragments) {
        for (const fragment of placement.highlightFragments) {
          ctx.fillStyle = fragment.color;
          ctx.fillRect(fragment.rect.xPt, fragment.rect.yPt, fragment.rect.widthPt, fragment.rect.heightPt);
        }
      } else if (placement.background || placement.highlight) {
        ctx.fillStyle = placement.highlight ?? placement.background ?? '#000000';
        ctx.fillRect(placement.bounds.xPt, placement.bounds.yPt, placement.bounds.widthPt, placement.bounds.heightPt);
      }
      ctx.fillStyle = textColor(placement, context);
      ctx.font = canvasFontString(
        placement.fontRoute,
        placement.fontSizePt,
        placement.fontWeight,
        placement.fontStyle,
      );
      ctx.textAlign = 'left';
      ctx.textBaseline = 'alphabetic';
      const previousLetterSpacing = ctx.letterSpacing;
      const previousKerning = ctx.fontKerning;
      for (const op of placement.paintOps) {
        ctx.direction = op.direction;
        ctx.fontKerning = op.kerning;
        const originXPt = placement.origin.xPt + op.offset.xPt;
        const originYPt = placement.origin.yPt + op.offset.yPt;
        if (op.glyphOrientation === 'upright') {
          ctx.save();
          ctx.translate(originXPt, originYPt);
          ctx.rotate(-Math.PI / 2);
          if (op.scaleX !== 1) ctx.scale(op.scaleX, 1);
          ctx.textAlign = 'center';
          ctx.textBaseline = 'middle';
          ctx.letterSpacing = `${op.letterSpacingPt}px`;
          ctx.fillText(op.text, 0, 0);
          ctx.restore();
        } else if (op.scaleX !== 1) {
          ctx.save();
          ctx.translate(originXPt, originYPt);
          ctx.scale(op.scaleX, 1);
          ctx.letterSpacing = `${op.letterSpacingPt / op.scaleX}px`;
          ctx.fillText(op.text, 0, 0);
          ctx.restore();
        } else {
          ctx.letterSpacing = `${op.letterSpacingPt}px`;
          ctx.fillText(op.text, originXPt, originYPt);
        }
      }
      ctx.letterSpacing = previousLetterSpacing;
      ctx.fontKerning = previousKerning;
      if (placement.ruby) {
        const uprightRuby = context.textBoxVerticalMode === 'eaVert'
          || context.textBoxVerticalMode === 'mongolianVert';
        for (const operation of placement.ruby.paintOps) {
          paintRetainedGlyph(operation, context, uprightRuby);
        }
      }
      for (const operation of placement.emphasis?.glyphs ?? []) {
        paintRetainedGlyph(operation, context);
      }
      for (const path of placement.emphasis?.paths ?? []) paintRetainedMarkPath(path, context);
      if (context.onTextRun) {
        const transform = context.textRunTransform ?? {
          translateXPt: 0,
          translateYPt: 0,
          scale: 1,
        };
        const scale = transform.scale;
        const letterSpacingPt = placement.paintOps[0]?.letterSpacingPt ?? 0;
        context.onTextRun(textRunPaintInfo({
          text: placement.text,
          x: (transform.translateXPt + placement.bounds.xPt) * scale,
          y: (transform.translateYPt + placement.bounds.yPt) * scale,
          w: placement.bounds.widthPt * scale,
          h: placement.bounds.heightPt * scale,
          fontSize: placement.fontSizePt * scale,
          font: canvasFontString(
            placement.fontRoute,
            placement.fontSizePt * scale,
            placement.fontWeight,
            placement.fontStyle,
          ),
          ...(letterSpacingPt !== 0 ? { letterSpacingPx: letterSpacingPt * scale } : {}),
          ...(placement.hyperlink ? {
            hyperlink: { kind: 'external' as const, url: placement.hyperlink },
          } : {}),
          ...(placement.tateChuYoko ? { eastAsianVert: true } : {}),
        }));
      }
      for (const decoration of placement.decorations) paintStrokeSegment(decoration, context);
      for (const border of placement.runBorderFragments ?? []) paintStrokeSegment(border, context);
    }
  }
  for (const border of node.borders) paintStrokeSegment(border, context);
  for (const drawing of node.drawings.filter((item) => !item.anchorLayer)) {
    paintDrawingWithTextBoxes(drawing);
  }
  const front = node.drawings
    .filter((drawing) => drawing.anchorLayer && !drawing.anchorLayer.behindDoc)
    .sort((a, b) => a.anchorLayer!.relativeHeight - b.anchorLayer!.relativeHeight
      || a.anchorLayer!.sourceOrder - b.anchorLayer!.sourceOrder);
  for (const drawing of front) {
    if (!context.deferFrontDrawing?.(drawing, textBoxesFor(drawing))) {
      paintDrawingWithTextBoxes(drawing);
    }
  }
  for (const textBox of node.textBoxes) {
    if (!ownedTextBoxIds.has(textBox.id)) paintTextBoxLayout(textBox, context);
  }
  } finally {
    if (node.clipBounds) ctx.restore();
  }
}

/** Paints an acquired text box. All line partitioning, glyph shaping and point
 * geometry are owned by acquisition; this function only traverses paint data. */
export function paintTextBoxLayout(node: TextBoxLayout, context: CanvasPaintContext): void {
  if (!node.verticalMode) {
    for (const paragraph of node.paragraphs) paintParagraphLayout(paragraph, context);
    return;
  }
  context.ctx.save();
  try {
    const center = translationAffine(
      node.flowBounds.xPt + node.flowBounds.widthPt / 2,
      node.flowBounds.yPt + node.flowBounds.heightPt / 2,
    );
    const turn = quarterTurnAffine(node.verticalMode === 'vert270' ? -1 : 1);
    const pointToCss = composeAffine(
      context.pointToCss ?? scaleAffine(context.scale),
      composeAffine(center, turn),
    );
    context.ctx.translate(
      node.flowBounds.xPt + node.flowBounds.widthPt / 2,
      node.flowBounds.yPt + node.flowBounds.heightPt / 2,
    );
    context.ctx.rotate(node.verticalMode === 'vert270' ? -Math.PI / 2 : Math.PI / 2);
    for (const paragraph of node.paragraphs) {
      paintParagraphLayout(paragraph, {
        ...context,
        pointToCss,
        textBoxVerticalMode: node.verticalMode,
      });
    }
  } finally {
    context.ctx.restore();
  }
}

/** Paints an absolute point-space text box into a CSS-pixel canvas viewport. */
export function paintPlacedTextBoxLayout(
  node: TextBoxLayout,
  context: CanvasPaintContext,
): void {
  context.ctx.save();
  try {
    context.ctx.scale(context.scale, context.scale);
    paintTextBoxLayout(node, {
      ...context,
      pointToCss: context.pointToCss ?? scaleAffine(context.scale),
    });
  } finally {
    context.ctx.restore();
  }
}

/** Paint a retained paragraph at a page placement using one point-to-CSS transform. */
export function paintPlacedParagraphLayout(
  node: ParagraphLayout,
  placement: Readonly<{ xPt: number; yPt: number }>,
  context: CanvasPaintContext,
): void {
  const dxPt = placement.xPt - node.flowBounds.xPt;
  const dyPt = placement.yPt - node.flowBounds.yPt;
  const pointToCss = composeAffine(
    context.pointToCss ?? scaleAffine(context.scale),
    translationAffine(dxPt, dyPt),
  );
  context.ctx.save();
  try {
    context.ctx.translate(dxPt * context.scale, dyPt * context.scale);
    context.ctx.scale(context.scale, context.scale);
    paintParagraphLayout(node, {
      ...context,
      pointToCss,
      textRunTransform: {
        translateXPt: dxPt,
        translateYPt: dyPt,
        scale: context.scale,
      },
      layoutTranslationPt: { xPt: dxPt, yPt: dyPt },
    });
  } finally {
    context.ctx.restore();
  }
}

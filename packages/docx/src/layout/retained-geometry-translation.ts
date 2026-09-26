import type { AnchorFrameResult } from './anchor-frame.js';
import type {
  BorderSegment, ClipPathData, DrawingLayout, DrawingPaintCommand, LayoutNodeId,
  FloatingTablePositionInput, LayoutRect, LineLayout, ParagraphLayout, ParagraphPlacement, PointPt, TableLayout,
  TextBoxLayout,
} from './types.js';

export interface LayoutTranslation { readonly xPt: number; readonly yPt: number }

/** Keep occurrence translation a dependency-free leaf: projection may call it,
 * so importing general geometry helpers here would reopen a runtime cycle seam. */
function unionTranslatedRects(rects: readonly LayoutRect[]): LayoutRect | null {
  if (rects.length === 0) return null;
  const left = Math.min(...rects.map((rect) => rect.xPt));
  const top = Math.min(...rects.map((rect) => rect.yPt));
  const right = Math.max(...rects.map((rect) => rect.xPt + rect.widthPt));
  const bottom = Math.max(...rects.map((rect) => rect.yPt + rect.heightPt));
  return {
    xPt: left,
    yPt: top,
    widthPt: right - left,
    heightPt: bottom - top,
  };
}

/** Initial tblpPr placement and retained occurrence translation must share this
 * predicate or their resolved frame and later axis ownership diverge. */
export function floatingTableAxesFollowHostFlow(
  positioning: FloatingTablePositionInput,
): Readonly<{ x: boolean; y: boolean }> {
  return {
    x: !positioning.horzSpecified
      || (positioning.horzAnchor !== 'page' && positioning.horzAnchor !== 'margin'),
    y: positioning.vertAnchor !== 'page' && positioning.vertAnchor !== 'margin',
  };
}

interface ParagraphTranslationContext {
  readonly memo: WeakMap<ParagraphLayout, { readonly key: string; readonly value: ParagraphLayout }>;
  readonly drawingMemo: WeakMap<DrawingLayout, { readonly key: string; readonly value: DrawingLayout }>;
}

export function translatePoint(point: PointPt, delta: LayoutTranslation): PointPt {
  return { ...point, xPt: point.xPt + delta.xPt, yPt: point.yPt + delta.yPt };
}

export function translateRect(rect: LayoutRect, delta: LayoutTranslation): LayoutRect {
  return { ...rect, xPt: rect.xPt + delta.xPt, yPt: rect.yPt + delta.yPt };
}

export function translateBorder(border: BorderSegment, delta: LayoutTranslation): BorderSegment {
  return { ...border, from: translatePoint(border.from, delta), to: translatePoint(border.to, delta) };
}

function translateClip(clip: ClipPathData, delta: LayoutTranslation): ClipPathData {
  return clip.kind === 'rect'
    ? { ...clip, rect: translateRect(clip.rect, delta) }
    : { ...clip, points: clip.points.map((point) => translatePoint(point, delta)) };
}

function translateDrawingCommand(command: DrawingPaintCommand, delta: LayoutTranslation): DrawingPaintCommand {
  if (command.kind === 'noop') return command;
  if (command.kind === 'drawingml-shape' || command.kind === 'drawingml-image-fill') return {
    ...command,
    plan: { ...command.plan, rect: {
      ...command.plan.rect,
      x: command.plan.rect.x + delta.xPt,
      y: command.plan.rect.y + delta.yPt,
    } },
  };
  return { ...command, rect: translateRect(command.rect, delta) };
}

export function translateDrawing(drawing: DrawingLayout, delta: LayoutTranslation): DrawingLayout {
  // Upright physical commands/text boxes are drawing-local. Only their shared
  // logical placement transform moves with the retained anchor occurrence.
  const commandDelta = drawing.orientation === 'upright-physical'
    ? { xPt: 0, yPt: 0 }
    : delta;
  return {
    ...drawing,
    flowBounds: translateRect(drawing.flowBounds, delta),
    inkBounds: translateRect(drawing.inkBounds, delta),
    ...(drawing.clipBounds ? { clipBounds: translateRect(drawing.clipBounds, delta) } : {}),
    ...(drawing.transform ? { transform: {
      ...drawing.transform, e: drawing.transform.e + delta.xPt, f: drawing.transform.f + delta.yPt,
    } } : {}),
    ...(drawing.clip ? { clip: translateClip(drawing.clip, delta) } : {}),
    commands: drawing.commands.map((command) => translateDrawingCommand(command, commandDelta)),
  };
}

function translateDrawingWithContext(
  drawing: DrawingLayout,
  delta: LayoutTranslation,
  context: ParagraphTranslationContext,
): DrawingLayout {
  const key = `${delta.xPt}\u0000${delta.yPt}`;
  const prior = context.drawingMemo.get(drawing);
  if (prior) {
    if (prior.key !== key) throw new Error('incompatible projection ownership');
    return prior.value;
  }
  const value = translateDrawing(drawing, delta);
  context.drawingMemo.set(drawing, { key, value });
  return value;
}

export function translatePlacement(
  placement: ParagraphPlacement,
  delta: LayoutTranslation,
  drawingTranslations?: ReadonlyMap<LayoutNodeId, LayoutTranslation>,
): ParagraphPlacement {
  if (placement.kind === 'text') return {
    ...placement,
    origin: translatePoint(placement.origin, delta), bounds: translateRect(placement.bounds, delta),
    decorations: placement.decorations.map((decoration) => ({
      ...decoration, from: translatePoint(decoration.from, delta), to: translatePoint(decoration.to, delta),
      ...(decoration.path ? { path: decoration.path.map((point) => translatePoint(point, delta)) } : {}),
    })),
    ...(placement.highlightFragments ? { highlightFragments: placement.highlightFragments.map((fragment) => ({
      ...fragment, rect: translateRect(fragment.rect, delta),
    })) } : {}),
    ...(placement.ruby ? { ruby: { ...placement.ruby, paintOps: placement.ruby.paintOps.map((operation) => ({
      ...operation, origin: translatePoint(operation.origin, delta),
    })) } } : {}),
    ...(placement.emphasis ? { emphasis: {
      ...placement.emphasis,
      ...(placement.emphasis.glyphs ? { glyphs: placement.emphasis.glyphs.map((glyph) => ({
        ...glyph, origin: translatePoint(glyph.origin, delta),
      })) } : {}),
      ...(placement.emphasis.paths ? { paths: placement.emphasis.paths.map((path) => ({
        ...path, points: path.points.map((point) => translatePoint(point, delta)),
      })) } : {}),
    } } : {}),
    ...(placement.runBorderFragments ? { runBorderFragments: placement.runBorderFragments.map((border) =>
      translateBorder(border, delta)) } : {}),
  };
  if (placement.kind === 'anchor-host') return {
    ...placement, bounds: translateRect(placement.bounds, delta), baselinePt: placement.baselinePt + delta.yPt,
  };
  if (placement.kind === 'drawing') return {
    ...placement, bounds: translateRect(placement.bounds, drawingTranslations?.get(placement.drawingId) ?? delta),
  };
  if (placement.kind === 'tab' && (placement.leaderGlyphs || placement.decorations)) return {
    ...placement,
    ...(placement.bounds ? { bounds: translateRect(placement.bounds, delta) } : {}),
    ...(placement.leaderGlyphs ? { leaderGlyphs: placement.leaderGlyphs.map((operation) => ({
      ...operation, origin: translatePoint(operation.origin, delta),
    })) } : {}),
    ...(placement.decorations ? { decorations: placement.decorations.map((decoration) => ({
      ...decoration,
      from: translatePoint(decoration.from, delta),
      to: translatePoint(decoration.to, delta),
      ...(decoration.path ? { path: decoration.path.map((point) => translatePoint(point, delta)) } : {}),
    })) } : {}),
  };
  return placement.bounds ? { ...placement, bounds: translateRect(placement.bounds, delta) } : placement;
}

export function translateLine(
  line: LineLayout,
  delta: LayoutTranslation,
  drawingTranslations?: ReadonlyMap<LayoutNodeId, LayoutTranslation>,
): LineLayout {
  return {
    ...line,
    bounds: translateRect(line.bounds, delta),
    baselinePt: line.baselinePt + delta.yPt,
    placements: line.placements.map((placement) => translatePlacement(placement, delta, drawingTranslations)),
    ...(line.barTabRules ? { barTabRules: line.barTabRules.map((rule) => ({
      ...rule,
      from: translatePoint(rule.from, delta),
      to: translatePoint(rule.to, delta),
    })) } : {}),
  };
}

function axisIsPageOwned(frame: AnchorFrameResult, axis: 'horizontal' | 'vertical'): boolean {
  const diagnostic = frame.axes[axis];
  return diagnostic.status === 'resolved'
    && ['page', 'margin', 'leftMargin', 'rightMargin', 'topMargin', 'bottomMargin'].includes(diagnostic.referenceFrame);
}

function translateAnchorFrame(frame: AnchorFrameResult, delta: LayoutTranslation): AnchorFrameResult {
  const xPt = axisIsPageOwned(frame, 'horizontal') ? 0 : delta.xPt;
  const yPt = axisIsPageOwned(frame, 'vertical') ? 0 : delta.yPt;
  const frameDelta = { xPt, yPt };
  const axes = {
    horizontal: frame.axes.horizontal.status === 'resolved' ? {
      ...frame.axes.horizontal,
      baseStartPt: frame.axes.horizontal.baseStartPt + xPt,
      baseEndPt: frame.axes.horizontal.baseEndPt + xPt,
      resolvedOriginPt: frame.axes.horizontal.resolvedOriginPt + xPt,
    } : frame.axes.horizontal,
    vertical: frame.axes.vertical.status === 'resolved' ? {
      ...frame.axes.vertical,
      baseStartPt: frame.axes.vertical.baseStartPt + yPt,
      baseEndPt: frame.axes.vertical.baseEndPt + yPt,
      resolvedOriginPt: frame.axes.vertical.resolvedOriginPt + yPt,
    } : frame.axes.vertical,
  };
  if (frame.status === 'unsupported') return { ...frame, axes };
  return {
    ...frame, axes,
    geometry: {
      ...frame.geometry,
      objectFrame: translateRect(frame.geometry.objectFrame, frameDelta),
      inkBounds: translateRect(frame.geometry.inkBounds, frameDelta),
      wrapBounds: frame.geometry.wrapBounds ? translateRect(frame.geometry.wrapBounds, frameDelta) : null,
      wrap: { ...frame.geometry.wrap, polygon: frame.geometry.wrap.polygon ? {
        ...frame.geometry.wrap.polygon,
        points: frame.geometry.wrap.polygon.points.map((point) => translatePoint(point, frameDelta)),
      } : null },
    },
  };
}

/** Translate retained paragraph geometry without invoking acquisition services. */
export function translateParagraphLayout(paragraph: ParagraphLayout, delta: LayoutTranslation): ParagraphLayout {
  return translateParagraphWithContext(paragraph, delta, {
    memo: new WeakMap(), drawingMemo: new WeakMap(),
  });
}

function translateParagraphWithContext(
  paragraph: ParagraphLayout,
  delta: LayoutTranslation,
  context: ParagraphTranslationContext,
): ParagraphLayout {
  const key = `${delta.xPt}\u0000${delta.yPt}`;
  const prior = context.memo.get(paragraph);
  if (prior) {
    if (prior.key !== key) throw new Error('incompatible projection ownership');
    return prior.value;
  }
  const anchorOwnership = new Map(paragraph.drawings.flatMap((drawing) =>
    drawing.anchorLayer ? [[drawing.anchorLayer.occurrenceId, drawing.anchorLayer] as const] : []));
  const textBoxTranslations = new Map<LayoutNodeId, LayoutTranslation>();
  const drawingTranslations = new Map<LayoutNodeId, LayoutTranslation>();
  for (const drawing of paragraph.drawings) {
    const drawingDelta = {
      xPt: drawing.anchorLayer?.horizontalOwnership === 'page' ? 0 : delta.xPt,
      yPt: drawing.anchorLayer?.verticalOwnership === 'page' ? 0 : delta.yPt,
    };
    drawingTranslations.set(drawing.id, drawingDelta);
    drawing.textBoxIds?.forEach((id) => textBoxTranslations.set(
      id,
      drawing.orientation === 'upright-physical'
        ? { xPt: 0, yPt: 0 }
        : drawingDelta,
    ));
  }
  const drawings = paragraph.drawings.map((drawing) => translateDrawingWithContext(
    drawing, drawingTranslations.get(drawing.id) ?? delta, context,
  ));
  const cellContainmentBounds = unionTranslatedRects(
    drawings
      .filter((drawing) => drawing.anchorLayer?.cellContainment === true)
      .map((drawing) => drawing.flowBounds),
  );
  const translated: ParagraphLayout = {
    ...paragraph,
    flowBounds: translateRect(paragraph.flowBounds, delta), inkBounds: translateRect(paragraph.inkBounds, delta),
    ...(paragraph.clipBounds ? { clipBounds: translateRect(paragraph.clipBounds, delta) } : {}),
    lines: paragraph.lines.map((line) => translateLine(line, delta, drawingTranslations)),
    borders: paragraph.borders.map((border) => translateBorder(border, delta)),
    drawings,
    ...(cellContainmentBounds ? { cellContainmentBounds } : {}),
    textBoxes: paragraph.textBoxes.map((textBox) =>
      translateTextBoxWithContext(textBox, textBoxTranslations.get(textBox.id) ?? delta, context)),
    exclusions: paragraph.exclusions.map((exclusion) => {
      const owner = exclusion.anchorOccurrenceId ? anchorOwnership.get(exclusion.anchorOccurrenceId) : undefined;
      const exclusionDelta = {
        xPt: owner?.horizontalOwnership === 'page' ? 0 : delta.xPt,
        yPt: exclusion.verticalOwnership === 'page' || owner?.verticalOwnership === 'page' ? 0 : delta.yPt,
      };
      return {
        ...exclusion, bounds: translateRect(exclusion.bounds, exclusionDelta),
        polygon: exclusion.polygon.map((point) => translatePoint(point, exclusionDelta)),
      };
    }),
    ...(paragraph.anchorCollisions ? {
      anchorCollisions: paragraph.anchorCollisions.map((entry) => ({
        ...entry,
        bounds: translateRect(entry.bounds, {
          xPt: entry.horizontalOwnership === 'page' ? 0 : delta.xPt,
          yPt: entry.verticalOwnership === 'page' ? 0 : delta.yPt,
        }),
      })),
    } : {}),
    ...(paragraph.anchorFrames ? { anchorFrames: paragraph.anchorFrames.map((frame) =>
      translateAnchorFrame(frame, delta)) } : {}),
    ...(paragraph.paragraphMark ? { paragraphMark: {
      ...paragraph.paragraphMark, bounds: translateRect(paragraph.paragraphMark.bounds, delta),
    } } : {}),
    ...(paragraph.lineNumbers ? { lineNumbers: paragraph.lineNumbers.map((lineNumber) => ({
      ...lineNumber, bounds: translateRect(lineNumber.bounds, delta),
      paintOps: lineNumber.paintOps.map((operation) => ({
        ...operation, origin: translatePoint(operation.origin, delta),
      })),
    })) } : {}),
  };
  context.memo.set(paragraph, { key, value: translated });
  return translated;
}

export function translateTextBox(textBox: TextBoxLayout, delta: LayoutTranslation): TextBoxLayout {
  return translateTextBoxWithContext(textBox, delta, {
    memo: new WeakMap(), drawingMemo: new WeakMap(),
  });
}

function translateTextBoxWithContext(
  textBox: TextBoxLayout,
  delta: LayoutTranslation,
  context: ParagraphTranslationContext,
): TextBoxLayout {
  const pageRelativeContent = textBox.verticalMode === undefined;
  return {
    ...textBox,
    flowBounds: translateRect(textBox.flowBounds, delta), inkBounds: translateRect(textBox.inkBounds, delta),
    ...(textBox.clipBounds ? { clipBounds: pageRelativeContent
      ? translateRect(textBox.clipBounds, delta) : textBox.clipBounds } : {}),
    ...(textBox.contentBounds ? { contentBounds: pageRelativeContent
      ? translateRect(textBox.contentBounds, delta) : textBox.contentBounds } : {}),
    transform: pageRelativeContent ? textBox.transform : {
      ...textBox.transform,
      e: textBox.transform.e + delta.xPt,
      f: textBox.transform.f + delta.yPt,
    },
    story: pageRelativeContent ? {
      ...textBox.story,
      flowBounds: translateRect(textBox.story.flowBounds, delta),
      inkBounds: translateRect(textBox.story.inkBounds, delta),
      ...(textBox.story.clipBounds
        ? { clipBounds: translateRect(textBox.story.clipBounds, delta) } : {}),
      blocks: textBox.story.blocks.map((block) => {
        if (block.kind === 'paragraph') {
          return translateParagraphWithContext(block, delta, context);
        }
        if (block.kind === 'table') return translateTableLayout(block, delta);
        throw new Error(`Text-box story contains unsupported retained node: ${block.kind}`);
      }),
    } : textBox.story,
  };
}

export function translateCompleteParagraphLayout(
  paragraph: ParagraphLayout,
  delta: LayoutTranslation,
): ParagraphLayout {
  return translateParagraphLayout(paragraph, delta);
}

export function translateTableLayout(table: TableLayout, delta: LayoutTranslation): TableLayout {
  return {
    ...table,
    flowBounds: translateRect(table.flowBounds, delta), inkBounds: translateRect(table.inkBounds, delta),
    ...(table.clipBounds ? { clipBounds: translateRect(table.clipBounds, delta) } : {}),
    borders: table.borders.map((border) => translateBorder(border, delta)),
    ...(table.compoundBorderFrames ? {
      compoundBorderFrames: table.compoundBorderFrames.map((frame) => ({
        ...frame,
        bounds: translateRect(frame.bounds, delta),
      })),
    } : {}),
    rows: table.rows.map((row) => ({
      ...row,
      flowBounds: translateRect(row.flowBounds, delta), inkBounds: translateRect(row.inkBounds, delta),
      ...(row.clipBounds ? { clipBounds: translateRect(row.clipBounds, delta) } : {}),
      cells: row.cells.map((cell) => ({
        ...cell,
        flowBounds: translateRect(cell.flowBounds, delta), inkBounds: translateRect(cell.inkBounds, delta),
        ...(cell.clipBounds ? { clipBounds: translateRect(cell.clipBounds, delta) } : {}),
        // A rotated cell (ECMA-376 §17.4.72) keeps its content in the local
        // text frame; only the frame-to-page transform moves with the table.
        ...(cell.verticalText ? {
          contentBounds: cell.contentBounds,
          verticalText: {
            ...cell.verticalText,
            transform: {
              ...cell.verticalText.transform,
              e: cell.verticalText.transform.e + delta.xPt,
              f: cell.verticalText.transform.f + delta.yPt,
            },
          },
        } : { contentBounds: translateRect(cell.contentBounds, delta) }),
        // Cell paint adds contentBounds/offsetPt; retained descendants are cell-local.
        blocks: cell.blocks,
      })),
    })),
  };
}

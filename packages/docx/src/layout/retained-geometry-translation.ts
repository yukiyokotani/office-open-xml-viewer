import type {
  AnchorAxisDiagnostic,
  AnchorFrameResult,
} from './anchor-frame.js';
import type {
  BorderSegment,
  ClipPathData,
  DrawingLayout,
  DrawingPaintCommand,
  LayoutNodeId,
  LayoutRect,
  LineLayout,
  ParagraphLayout,
  ParagraphPlacement,
  PointPt,
  TextBoxLayout,
} from './types.js';

export interface LayoutTranslation {
  readonly xPt: number;
  readonly yPt: number;
}

export function translatePoint(point: PointPt, delta: LayoutTranslation): PointPt {
  return { ...point, xPt: point.xPt + delta.xPt, yPt: point.yPt + delta.yPt };
}

export function translateRect(rect: LayoutRect, delta: LayoutTranslation): LayoutRect {
  return { ...rect, xPt: rect.xPt + delta.xPt, yPt: rect.yPt + delta.yPt };
}

function translateClip(clip: ClipPathData, delta: LayoutTranslation): ClipPathData {
  return clip.kind === 'rect'
    ? { ...clip, rect: translateRect(clip.rect, delta) }
    : { ...clip, points: clip.points.map((point) => translatePoint(point, delta)) };
}

function translateDrawingCommand(
  command: DrawingPaintCommand,
  delta: LayoutTranslation,
): DrawingPaintCommand {
  if (command.kind === 'noop') return command;
  if (command.kind === 'drawingml-shape') return {
    ...command,
    plan: {
      ...command.plan,
      rect: {
        ...command.plan.rect,
        x: command.plan.rect.x + delta.xPt,
        y: command.plan.rect.y + delta.yPt,
      },
    },
  };
  return { ...command, rect: translateRect(command.rect, delta) };
}

export function translateDrawing(
  drawing: DrawingLayout,
  delta: LayoutTranslation,
): DrawingLayout {
  return {
    ...drawing,
    flowBounds: translateRect(drawing.flowBounds, delta),
    inkBounds: translateRect(drawing.inkBounds, delta),
    ...(drawing.clipBounds ? { clipBounds: translateRect(drawing.clipBounds, delta) } : {}),
    ...(drawing.transform ? {
      transform: {
        ...drawing.transform,
        e: drawing.transform.e + delta.xPt,
        f: drawing.transform.f + delta.yPt,
      },
    } : {}),
    ...(drawing.clip ? { clip: translateClip(drawing.clip, delta) } : {}),
    commands: drawing.commands.map((command) => translateDrawingCommand(command, delta)),
  };
}

function translateAcquiredAnchorHostFrames(
  drawing: DrawingLayout,
  delta: LayoutTranslation,
): DrawingLayout {
  const anchor = drawing.anchorLayer;
  if (!anchor || anchor.coordinateSpace !== 'acquired-anchor-points') return drawing;
  const frames = anchor.normalization.logicalHostFrames;
  return {
    ...drawing,
    anchorLayer: {
      ...anchor,
      normalization: {
        ...anchor.normalization,
        logicalHostFrames: {
          paragraph: translateRect(frames.paragraph, delta),
          line: translateRect(frames.line, delta),
          character: translateRect(frames.character, delta),
        },
      },
    },
  };
}

export function translateBorder(border: BorderSegment, delta: LayoutTranslation): BorderSegment {
  return {
    ...border,
    from: translatePoint(border.from, delta),
    to: translatePoint(border.to, delta),
  };
}

export function translatePlacement(
  placement: ParagraphPlacement,
  delta: LayoutTranslation,
  drawingTranslations: ReadonlyMap<LayoutNodeId, LayoutTranslation> = new Map(),
): ParagraphPlacement {
  if (placement.kind === 'text') return {
    ...placement,
    origin: translatePoint(placement.origin, delta),
    bounds: translateRect(placement.bounds, delta),
    decorations: placement.decorations.map((decoration) => ({
      ...decoration,
      from: translatePoint(decoration.from, delta),
      to: translatePoint(decoration.to, delta),
      ...(decoration.path ? {
        path: decoration.path.map((point) => translatePoint(point, delta)),
      } : {}),
    })),
    ...(placement.highlightFragments ? {
      highlightFragments: placement.highlightFragments.map((fragment) => ({
        ...fragment,
        rect: translateRect(fragment.rect, delta),
      })),
    } : {}),
    ...(placement.ruby ? {
      ruby: {
        ...placement.ruby,
        paintOps: placement.ruby.paintOps.map((operation) => ({
          ...operation,
          origin: translatePoint(operation.origin, delta),
        })),
      },
    } : {}),
    ...(placement.emphasis ? {
      emphasis: {
        ...placement.emphasis,
        ...(placement.emphasis.glyphs ? {
          glyphs: placement.emphasis.glyphs.map((glyph) => ({
            ...glyph,
            origin: translatePoint(glyph.origin, delta),
          })),
        } : {}),
        ...(placement.emphasis.paths ? {
          paths: placement.emphasis.paths.map((path) => ({
            ...path,
            points: path.points.map((point) => translatePoint(point, delta)),
          })),
        } : {}),
      },
    } : {}),
    ...(placement.runBorderFragments ? {
      runBorderFragments: placement.runBorderFragments.map((border) => (
        translateBorder(border, delta)
      )),
    } : {}),
  };
  if (placement.kind === 'anchor-host') return {
    ...placement,
    bounds: translateRect(placement.bounds, delta),
    baselinePt: placement.baselinePt + delta.yPt,
  };
  if (placement.kind === 'drawing') return {
    ...placement,
    bounds: translateRect(
      placement.bounds,
      drawingTranslations.get(placement.drawingId) ?? delta,
    ),
  };
  if (placement.kind === 'tab' && placement.leaderGlyphs) return {
    ...placement,
    ...(placement.bounds ? { bounds: translateRect(placement.bounds, delta) } : {}),
    leaderGlyphs: placement.leaderGlyphs.map((operation) => ({
      ...operation,
      origin: translatePoint(operation.origin, delta),
    })),
  };
  return placement.bounds
    ? { ...placement, bounds: translateRect(placement.bounds, delta) }
    : placement;
}

export function translateLine(
  line: LineLayout,
  delta: LayoutTranslation,
  drawingTranslations: ReadonlyMap<LayoutNodeId, LayoutTranslation> = new Map(),
): LineLayout {
  return {
    ...line,
    bounds: translateRect(line.bounds, delta),
    baselinePt: line.baselinePt + delta.yPt,
    placements: line.placements.map((placement) => (
      translatePlacement(placement, delta, drawingTranslations)
    )),
  };
}

/** Translate host-owned retained geometry while preserving axes owned by the
 * page anchor solver. */
export function translateParagraphLayout(
  paragraph: ParagraphLayout,
  delta: LayoutTranslation,
): ParagraphLayout {
  return translateParagraphLayoutInternal(paragraph, delta, false);
}

function translateParagraphLayoutInternal(
  paragraph: ParagraphLayout,
  delta: LayoutTranslation,
  complete: boolean,
): ParagraphLayout {
  const anchorOwnership = new Map(paragraph.drawings.flatMap((drawing) => (
    drawing.anchorLayer ? [[drawing.anchorLayer.occurrenceId, drawing.anchorLayer] as const] : []
  )));
  const textBoxTranslations = new Map<LayoutNodeId, LayoutTranslation>();
  const drawingTranslations = new Map<LayoutNodeId, LayoutTranslation>();
  paragraph.drawings.forEach((drawing) => {
    const drawingDelta = {
      xPt: drawing.anchorLayer?.horizontalOwnership === 'page' ? 0 : delta.xPt,
      yPt: drawing.anchorLayer?.verticalOwnership === 'page' ? 0 : delta.yPt,
    };
    drawingTranslations.set(drawing.id, drawingDelta);
    drawing.textBoxIds?.forEach((id) => textBoxTranslations.set(id, drawingDelta));
  });
  const translated: ParagraphLayout = {
    ...paragraph,
    flowBounds: translateRect(paragraph.flowBounds, delta),
    inkBounds: translateRect(paragraph.inkBounds, delta),
    ...(paragraph.clipBounds ? { clipBounds: translateRect(paragraph.clipBounds, delta) } : {}),
    lines: paragraph.lines.map((line) => translateLine(line, delta, drawingTranslations)),
    borders: paragraph.borders.map((border) => translateBorder(border, delta)),
    drawings: paragraph.drawings.map((drawing) => translateAcquiredAnchorHostFrames(
      translateDrawing(drawing, drawingTranslations.get(drawing.id) ?? delta),
      delta,
    )),
    textBoxes: paragraph.textBoxes.map((textBox) => (
      translateTextBoxInternal(
        textBox,
        textBoxTranslations.get(textBox.id) ?? delta,
        complete,
      )
    )),
    exclusions: paragraph.exclusions.map((exclusion) => {
      const owner = exclusion.anchorOccurrenceId
        ? anchorOwnership.get(exclusion.anchorOccurrenceId)
        : undefined;
      const exclusionDelta = {
        xPt: owner?.horizontalOwnership === 'page' ? 0 : delta.xPt,
        yPt: exclusion.verticalOwnership === 'page'
          || owner?.verticalOwnership === 'page' ? 0 : delta.yPt,
      };
      return {
        ...exclusion,
        bounds: translateRect(exclusion.bounds, exclusionDelta),
        polygon: exclusion.polygon.map((point) => translatePoint(point, exclusionDelta)),
      };
    }),
    ...(paragraph.paragraphMark ? {
      paragraphMark: {
        ...paragraph.paragraphMark,
        bounds: translateRect(paragraph.paragraphMark.bounds, delta),
      },
    } : {}),
  };
  if (!complete) return translated;
  return {
    ...translated,
    ...(paragraph.lineNumbers ? {
      lineNumbers: paragraph.lineNumbers.map((line) => ({
        ...line,
        bounds: translateRect(line.bounds, delta),
        paintOps: line.paintOps.map((operation) => ({
          ...operation,
          origin: translatePoint(operation.origin, delta),
        })),
      })),
    } : {}),
    ...(paragraph.anchorFrames ? {
      anchorFrames: paragraph.anchorFrames.map((frame) => {
        const owner = anchorOwnership.get(frame.occurrenceId);
        return translateAnchorFrame(frame, {
          xPt: owner?.horizontalOwnership === 'page' ? 0 : delta.xPt,
          yPt: owner?.verticalOwnership === 'page' ? 0 : delta.yPt,
        });
      }),
    } : {}),
  };
}

export function translateTextBox(
  textBox: TextBoxLayout,
  delta: LayoutTranslation,
): TextBoxLayout {
  return translateTextBoxInternal(textBox, delta, false);
}

function translateTextBoxInternal(
  textBox: TextBoxLayout,
  delta: LayoutTranslation,
  complete: boolean,
): TextBoxLayout {
  const pageRelativeContent = textBox.verticalMode === undefined;
  return {
    ...textBox,
    flowBounds: translateRect(textBox.flowBounds, delta),
    inkBounds: translateRect(textBox.inkBounds, delta),
    ...(textBox.clipBounds ? { clipBounds: translateRect(textBox.clipBounds, delta) } : {}),
    ...(textBox.contentBounds ? {
      contentBounds: pageRelativeContent
        ? translateRect(textBox.contentBounds, delta)
        : textBox.contentBounds,
    } : {}),
    paragraphs: pageRelativeContent
      ? textBox.paragraphs.map((paragraph) => (
          translateParagraphLayoutInternal(paragraph, delta, complete)
        ))
      : textBox.paragraphs,
  };
}

function translateAxis(axis: AnchorAxisDiagnostic, deltaPt: number): AnchorAxisDiagnostic {
  return axis.status === 'resolved' ? {
    ...axis,
    baseStartPt: axis.baseStartPt + deltaPt,
    baseEndPt: axis.baseEndPt + deltaPt,
    resolvedOriginPt: axis.resolvedOriginPt + deltaPt,
  } : axis;
}

function translateAnchorFrame(
  frame: AnchorFrameResult,
  delta: LayoutTranslation,
): AnchorFrameResult {
  const axes = {
    horizontal: translateAxis(frame.axes.horizontal, delta.xPt),
    vertical: translateAxis(frame.axes.vertical, delta.yPt),
  };
  if (frame.status === 'unsupported') return { ...frame, axes };
  const polygon = frame.geometry.wrap.polygon;
  return {
    ...frame,
    axes,
    geometry: {
      ...frame.geometry,
      objectFrame: translateRect(frame.geometry.objectFrame, delta),
      inkBounds: translateRect(frame.geometry.inkBounds, delta),
      wrapBounds: frame.geometry.wrapBounds
        ? translateRect(frame.geometry.wrapBounds, delta)
        : null,
      wrap: {
        ...frame.geometry.wrap,
        polygon: polygon ? {
          ...polygon,
          points: polygon.points.map((point) => translatePoint(point, delta)),
        } : null,
      },
    },
  };
}

/** Complete final-occurrence translation. Page-relative text-box descendants
 * recurse through this mode; vertical text-box children retain local geometry. */
export function translateCompleteParagraphLayout(
  paragraph: ParagraphLayout,
  delta: LayoutTranslation,
): ParagraphLayout {
  return translateParagraphLayoutInternal(paragraph, delta, true);
}

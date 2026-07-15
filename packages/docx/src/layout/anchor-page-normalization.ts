import { resolveAnchorFrame, type AnchorReferenceFramesInput } from './anchor-frame.js';
import type { AnchorAcquisitionInput } from './anchor-input.js';
import { composeAffine, mapAffineRect, quarterTurnAffine, translationAffine } from './affine.js';
import { snapshotPlainData } from './plain-data.js';
import { resizeResolvedAnchorGeometry } from './anchor-derived-geometry.js';
import type {
  BorderSegment,
  DrawingLayout,
  DrawingPaintCommand,
  LayoutRect,
  Matrix2DData,
  PaintNode,
  PaintReadyTableLayout,
  ParagraphLayout,
  ParagraphPlacement,
  PointPt,
  TableLayout,
  TextBoxLayout,
} from './types.js';

export interface PageAnchorNormalizationContext {
  readonly currentToPage: Matrix2DData;
  readonly normalizedFor: Readonly<{
    physicalPageIndex: number;
    flowDomainId: string;
    regionId: string;
  }>;
  readonly destinationFrames: Readonly<Pick<
    AnchorReferenceFramesInput,
    'page' | 'margin' | 'column' | 'pageParity'
  >>;
}

export interface AcquiredAnchorNormalizationFacts {
  readonly acquisition: Readonly<AnchorAcquisitionInput>;
  readonly pageParity: 'odd' | 'even' | null;
  readonly physicalFrames: Readonly<Pick<
    AnchorReferenceFramesInput,
    'page' | 'margin' | 'column'
  >>;
  readonly logicalHostFrames: Readonly<{
    paragraph: NonNullable<AnchorReferenceFramesInput['paragraph']>;
    line: NonNullable<AnchorReferenceFramesInput['line']>;
    character: NonNullable<AnchorReferenceFramesInput['character']>;
  }>;
}

export function normalizeAnchorReferenceFrames(
  facts: AcquiredAnchorNormalizationFacts,
  currentToPage: Matrix2DData,
  destinationFrames: PageAnchorNormalizationContext['destinationFrames'],
): AnchorReferenceFramesInput {
  return {
    ...destinationFrames,
    paragraph: mapAffineRect(currentToPage, facts.logicalHostFrames.paragraph),
    line: mapAffineRect(currentToPage, facts.logicalHostFrames.line),
    character: mapAffineRect(currentToPage, facts.logicalHostFrames.character),
  };
}

const IDENTITY: Matrix2DData = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };

interface NormalizedAnchor {
  readonly drawing: DrawingLayout;
  readonly reframe: AnchorRectReframe;
  readonly frame: Extract<ReturnType<typeof resolveAnchorFrame>, { status: 'resolved' }>;
}

interface AnchorRectReframe {
  readonly source: LayoutRect;
  readonly destination: LayoutRect;
  readonly scaleX: number;
  readonly scaleY: number;
}

function anchorRectReframe(source: LayoutRect, destination: LayoutRect): AnchorRectReframe {
  return {
    source,
    destination,
    scaleX: source.widthPt === 0 ? 1 : destination.widthPt / source.widthPt,
    scaleY: source.heightPt === 0 ? 1 : destination.heightPt / source.heightPt,
  };
}

function reframePoint(point: PointPt, reframe: AnchorRectReframe): PointPt {
  return {
    ...point,
    xPt: reframe.destination.xPt
      + (point.xPt - reframe.source.xPt) * reframe.scaleX,
    yPt: reframe.destination.yPt
      + (point.yPt - reframe.source.yPt) * reframe.scaleY,
  };
}

function reframeRect(rect: LayoutRect, reframe: AnchorRectReframe): LayoutRect {
  const origin = reframePoint(rect, reframe);
  return {
    ...rect,
    ...origin,
    widthPt: rect.widthPt * reframe.scaleX,
    heightPt: rect.heightPt * reframe.scaleY,
  };
}

function reframeBorder(border: BorderSegment, reframe: AnchorRectReframe): BorderSegment {
  return {
    ...border,
    from: reframePoint(border.from, reframe),
    to: reframePoint(border.to, reframe),
  };
}

function reframeCommand(
  command: DrawingPaintCommand,
  reframe: AnchorRectReframe,
): DrawingPaintCommand {
  if (command.kind === 'noop') return command;
  if (command.kind === 'drawingml-shape') {
    const rect = reframeRect({
      xPt: command.plan.rect.x,
      yPt: command.plan.rect.y,
      widthPt: command.plan.rect.w,
      heightPt: command.plan.rect.h,
    }, reframe);
    return {
      ...command,
      plan: {
        ...command.plan,
        rect: { x: rect.xPt, y: rect.yPt, w: rect.widthPt, h: rect.heightPt },
      },
    };
  }
  return { ...command, rect: reframeRect(command.rect, reframe) };
}

function reframeAcquiredDrawingFrames(
  drawing: DrawingLayout,
  reframe: AnchorRectReframe,
): DrawingLayout['anchorLayer'] {
  const anchor = drawing.anchorLayer;
  if (!anchor || anchor.coordinateSpace !== 'acquired-anchor-points') return anchor;
  return {
    ...anchor,
    normalization: {
      ...anchor.normalization,
      logicalHostFrames: {
        paragraph: reframeRect(anchor.normalization.logicalHostFrames.paragraph, reframe),
        line: reframeRect(anchor.normalization.logicalHostFrames.line, reframe),
        character: reframeRect(anchor.normalization.logicalHostFrames.character, reframe),
      },
    },
  };
}

function reframeDrawingGeometry(
  drawing: DrawingLayout,
  reframe: AnchorRectReframe,
  authoritative?: Extract<ReturnType<typeof resolveAnchorFrame>, { status: 'resolved' }>,
): DrawingLayout {
  const transform = drawing.transform;
  return {
    ...drawing,
    flowBounds: authoritative?.geometry.objectFrame ?? reframeRect(drawing.flowBounds, reframe),
    inkBounds: authoritative?.geometry.inkBounds ?? reframeRect(drawing.inkBounds, reframe),
    ...(drawing.clipBounds ? { clipBounds: reframeRect(drawing.clipBounds, reframe) } : {}),
    ...(transform ? {
      transform: {
        a: transform.a * reframe.scaleX,
        b: transform.b * reframe.scaleY,
        c: transform.c * reframe.scaleX,
        d: transform.d * reframe.scaleY,
        e: reframePoint({ xPt: transform.e, yPt: transform.f }, reframe).xPt,
        f: reframePoint({ xPt: transform.e, yPt: transform.f }, reframe).yPt,
      },
    } : {}),
    ...(drawing.clip ? {
      clip: drawing.clip.kind === 'rect'
        ? { ...drawing.clip, rect: reframeRect(drawing.clip.rect, reframe) }
        : { ...drawing.clip, points: drawing.clip.points.map((point) => reframePoint(point, reframe)) },
    } : {}),
    commands: drawing.commands.map((command) => reframeCommand(command, reframe)),
    ...(drawing.anchorLayer ? { anchorLayer: reframeAcquiredDrawingFrames(drawing, reframe) } : {}),
  };
}

function reframePlacement(
  placement: ParagraphPlacement,
  reframe: AnchorRectReframe,
): ParagraphPlacement {
  if (placement.kind === 'text') return {
    ...placement,
    origin: reframePoint(placement.origin, reframe),
    bounds: reframeRect(placement.bounds, reframe),
    advancePt: placement.advancePt * reframe.scaleX,
    // Point-valued offsets and advances belong to the acquired object frame.
    // Font selection, authored font size, and dimensionless glyph scaling stay
    // acquisition facts because page admission must not reshape or repartition.
    clusters: placement.clusters.map((cluster) => ({
      ...cluster,
      offset: {
        xPt: cluster.offset.xPt * reframe.scaleX,
        yPt: cluster.offset.yPt * reframe.scaleY,
      },
      advancePt: cluster.advancePt * reframe.scaleX,
    })),
    paintOps: placement.paintOps.map((operation) => ({
      ...operation,
      offset: {
        xPt: operation.offset.xPt * reframe.scaleX,
        yPt: operation.offset.yPt * reframe.scaleY,
      },
      letterSpacingPt: operation.letterSpacingPt * reframe.scaleX,
    })),
    ...(placement.characterSpacingPt !== undefined ? {
      characterSpacingPt: placement.characterSpacingPt * reframe.scaleX,
    } : {}),
    ...(placement.fitText ? {
      fitText: {
        ...placement.fitText,
        perGapPt: placement.fitText.perGapPt * reframe.scaleX,
        trailingPadPt: placement.fitText.trailingPadPt * reframe.scaleX,
      },
    } : {}),
    ...(placement.positionPt !== undefined ? {
      positionPt: placement.positionPt * reframe.scaleY,
    } : {}),
    ...(placement.ownedTrailingSlackPt !== undefined ? {
      ownedTrailingSlackPt: placement.ownedTrailingSlackPt * reframe.scaleX,
    } : {}),
    decorations: placement.decorations.map((decoration) => ({
      ...decoration,
      from: reframePoint(decoration.from, reframe),
      to: reframePoint(decoration.to, reframe),
      ...(decoration.path ? {
        path: decoration.path.map((point) => reframePoint(point, reframe)),
      } : {}),
    })),
    ...(placement.highlightFragments ? {
      highlightFragments: placement.highlightFragments.map((fragment) => ({
        ...fragment, rect: reframeRect(fragment.rect, reframe),
      })),
    } : {}),
    ...(placement.ruby ? {
      ruby: {
        ...placement.ruby,
        advancePt: placement.ruby.advancePt * reframe.scaleX,
        paintOps: placement.ruby.paintOps.map((operation) => ({
          ...operation, origin: reframePoint(operation.origin, reframe),
        })),
      },
    } : {}),
    ...(placement.emphasis ? {
      emphasis: {
        ...placement.emphasis,
        ...(placement.emphasis.glyphs ? {
          glyphs: placement.emphasis.glyphs.map((glyph) => ({
            ...glyph, origin: reframePoint(glyph.origin, reframe),
          })),
        } : {}),
        ...(placement.emphasis.paths ? {
          paths: placement.emphasis.paths.map((path) => ({
            ...path, points: path.points.map((point) => reframePoint(point, reframe)),
          })),
        } : {}),
      },
    } : {}),
    ...(placement.runBorderFragments ? {
      runBorderFragments: placement.runBorderFragments.map((border) => (
        reframeBorder(border, reframe)
      )),
    } : {}),
  };
  if (placement.kind === 'anchor-host') return {
    ...placement,
    bounds: reframeRect(placement.bounds, reframe),
    baselinePt: reframe.destination.yPt
      + (placement.baselinePt - reframe.source.yPt) * reframe.scaleY,
  };
  if (placement.kind === 'tab') return {
    ...placement,
    advancePt: placement.advancePt * reframe.scaleX,
    ...(placement.bounds ? { bounds: reframeRect(placement.bounds, reframe) } : {}),
    ...(placement.leaderGlyphs ? {
      leaderGlyphs: placement.leaderGlyphs.map((operation) => ({
        ...operation, origin: reframePoint(operation.origin, reframe),
      })),
    } : {}),
  };
  return {
    ...placement,
    bounds: reframeRect(placement.bounds, reframe),
    advancePt: placement.advancePt * reframe.scaleX,
  };
}

function reframeParagraphGeometry(
  paragraph: ParagraphLayout,
  reframe: AnchorRectReframe,
): ParagraphLayout {
  // Page admission is intentionally measurement-free. Retained glyph choices,
  // font sizes, and line partitions therefore remain authored acquisition
  // results; only their point-space paint geometry follows the resized shape
  // coordinate frame. Reflow would require a separate measured reacquisition.
  const anchoredTextBoxes = new Set(paragraph.drawings.flatMap((drawing) => (
    drawing.anchorLayer ? drawing.textBoxIds ?? [] : []
  )));
  const anchoredExclusions = new Set(paragraph.drawings.flatMap((drawing) => (
    drawing.anchorLayer ? [drawing.anchorLayer.occurrenceId] : []
  )));
  return {
    ...paragraph,
    flowBounds: reframeRect(paragraph.flowBounds, reframe),
    inkBounds: reframeRect(paragraph.inkBounds, reframe),
    ...(paragraph.clipBounds ? { clipBounds: reframeRect(paragraph.clipBounds, reframe) } : {}),
    advancePt: paragraph.advancePt * reframe.scaleY,
    lines: paragraph.lines.map((line) => ({
      ...line,
      bounds: reframeRect(line.bounds, reframe),
      baselinePt: reframe.destination.yPt
        + (line.baselinePt - reframe.source.yPt) * reframe.scaleY,
      advancePt: line.advancePt * reframe.scaleY,
      placements: line.placements.map((placement) => reframePlacement(placement, reframe)),
    })),
    borders: paragraph.borders.map((border) => reframeBorder(border, reframe)),
    drawings: paragraph.drawings.map((drawing) => {
      if (!drawing.anchorLayer) return reframeDrawingGeometry(drawing, reframe);
      if (drawing.anchorLayer.coordinateSpace === 'physical-page-points') return drawing;
      return { ...drawing, anchorLayer: reframeAcquiredDrawingFrames(drawing, reframe) };
    }),
    textBoxes: paragraph.textBoxes.map((textBox) => (
      anchoredTextBoxes.has(textBox.id) ? textBox : reframeTextBoxGeometry(textBox, reframe)
    )),
    exclusions: paragraph.exclusions.map((exclusion) => (
      exclusion.anchorOccurrenceId && anchoredExclusions.has(exclusion.anchorOccurrenceId)
        ? exclusion
        : {
            ...exclusion,
            bounds: reframeRect(exclusion.bounds, reframe),
            polygon: exclusion.polygon.map((point) => reframePoint(point, reframe)),
          }
    )),
    ...(paragraph.paragraphMark ? {
      paragraphMark: {
        ...paragraph.paragraphMark,
        bounds: reframeRect(paragraph.paragraphMark.bounds, reframe),
      },
    } : {}),
    ...(paragraph.lineNumbers ? {
      lineNumbers: paragraph.lineNumbers.map((line) => ({
        ...line,
        bounds: reframeRect(line.bounds, reframe),
        paintOps: line.paintOps.map((operation) => ({
          ...operation, origin: reframePoint(operation.origin, reframe),
        })),
      })),
    } : {}),
  };
}

function reframeTextBoxGeometry(
  textBox: TextBoxLayout,
  reframe: AnchorRectReframe,
): TextBoxLayout {
  const flowBounds = reframeRect(textBox.flowBounds, reframe);
  if (textBox.verticalMode) {
    const oldContent = textBox.contentBounds ?? {
      xPt: -textBox.flowBounds.heightPt / 2,
      yPt: -textBox.flowBounds.widthPt / 2,
      widthPt: textBox.flowBounds.heightPt,
      heightPt: textBox.flowBounds.widthPt,
    };
    const contentBounds = {
      xPt: -flowBounds.heightPt / 2,
      yPt: -flowBounds.widthPt / 2,
      widthPt: flowBounds.heightPt,
      heightPt: flowBounds.widthPt,
    };
    const localReframe = anchorRectReframe(oldContent, contentBounds);
    return {
      ...textBox,
      flowBounds,
      inkBounds: reframeRect(textBox.inkBounds, reframe),
      ...(textBox.clipBounds ? { clipBounds: reframeRect(textBox.clipBounds, reframe) } : {}),
      contentBounds,
      paragraphs: textBox.paragraphs.map((paragraph) => (
        reframeParagraphGeometry(paragraph, localReframe)
      )),
    };
  }
  return {
    ...textBox,
    flowBounds,
    inkBounds: reframeRect(textBox.inkBounds, reframe),
    ...(textBox.clipBounds ? { clipBounds: reframeRect(textBox.clipBounds, reframe) } : {}),
    ...(textBox.contentBounds ? { contentBounds: reframeRect(textBox.contentBounds, reframe) } : {}),
    paragraphs: textBox.paragraphs.map((paragraph) => (
      reframeParagraphGeometry(paragraph, reframe)
    )),
  };
}

function normalizeDrawingAnchor(
  drawing: DrawingLayout,
  context: PageAnchorNormalizationContext,
): NormalizedAnchor | null {
  const anchor = drawing.anchorLayer;
  if (!anchor || anchor.coordinateSpace === 'physical-page-points') return null;
  const resolved = resolveAnchorFrame({
    acquisition: anchor.normalization.acquisition,
    frames: normalizeAnchorReferenceFrames(
      anchor.normalization,
      context.currentToPage,
      context.destinationFrames,
    ),
  });
  if (resolved.status !== 'resolved') {
    throw new Error(`Anchor ${anchor.occurrenceId} cannot be normalized in its destination frame`);
  }
  const sourceResult = resolveAnchorFrame({
    acquisition: anchor.normalization.acquisition,
    frames: {
      ...anchor.normalization.physicalFrames,
      ...anchor.normalization.logicalHostFrames,
      pageParity: anchor.normalization.pageParity,
    },
  });
  const result = sourceResult.status === 'resolved'
    ? resizeResolvedAnchorGeometry(
        resolved,
        reframeRect(
          drawing.flowBounds,
          anchorRectReframe(sourceResult.geometry.objectFrame, resolved.geometry.objectFrame),
        ),
      )
    : resolved;
  // The final solver owns size as well as origin. Mapping the retained payload
  // between its authored object frames preserves group-child coordinates while
  // ink/wrap geometry comes directly from the destination solver transaction.
  const reframe = anchorRectReframe(drawing.flowBounds, result.geometry.objectFrame);
  const reframed = reframeDrawingGeometry(drawing, reframe, result);
  return {
    reframe,
    frame: result,
    drawing: {
      ...reframed,
      anchorLayer: {
        occurrenceId: anchor.occurrenceId,
        behindDoc: anchor.behindDoc,
        relativeHeight: anchor.relativeHeight,
        sourceOrder: anchor.sourceOrder,
        horizontalOwnership: anchor.horizontalOwnership,
        verticalOwnership: anchor.verticalOwnership,
        coordinateSpace: 'physical-page-points',
        normalizedFor: context.normalizedFor,
      },
    },
  };
}

function textBoxFrame(
  textBox: TextBoxLayout,
  parentToPage: Matrix2DData,
): Matrix2DData {
  if (!textBox.verticalMode) return parentToPage;
  const center = translationAffine(
    textBox.flowBounds.xPt + textBox.flowBounds.widthPt / 2,
    textBox.flowBounds.yPt + textBox.flowBounds.heightPt / 2,
  );
  const turn = quarterTurnAffine(textBox.verticalMode === 'vert270' ? -1 : 1);
  return composeAffine(parentToPage, composeAffine(center, turn));
}

function normalizeTextBox(
  textBox: TextBoxLayout,
  context: PageAnchorNormalizationContext,
): TextBoxLayout {
  const currentToPage = textBoxFrame(textBox, context.currentToPage);
  return {
    ...textBox,
    paragraphs: textBox.paragraphs.map((paragraph) => normalizeParagraph(paragraph, {
      ...context,
      currentToPage,
    })),
  };
}

function normalizeParagraph(
  paragraph: ParagraphLayout,
  context: PageAnchorNormalizationContext,
): ParagraphLayout {
  const normalizedById = new Map<string, NormalizedAnchor>();
  const normalizedByOccurrence = new Map<string, NormalizedAnchor>();
  const drawings = paragraph.drawings.map((drawing) => {
    const normalized = normalizeDrawingAnchor(drawing, context);
    if (!normalized) return drawing;
    normalizedById.set(drawing.id, normalized);
    normalizedByOccurrence.set(drawing.anchorLayer!.occurrenceId, normalized);
    return normalized.drawing;
  });
  const ownedTextBoxes = new Map<string, NormalizedAnchor>();
  paragraph.drawings.forEach((drawing) => {
    const normalized = normalizedById.get(drawing.id);
    if (!normalized) return;
    drawing.textBoxIds?.forEach((id) => ownedTextBoxes.set(id, normalized));
  });
  const textBoxes = paragraph.textBoxes.map((textBox) => {
    const owner = ownedTextBoxes.get(textBox.id);
    const reframed = owner ? reframeTextBoxGeometry(textBox, owner.reframe) : textBox;
    return normalizeTextBox(reframed, {
      ...context,
      // An admitted anchored box is absolute physical geometry; its internal
      // vertical turn is composed from the physical page frame, not the host.
      currentToPage: owner ? IDENTITY : context.currentToPage,
    });
  });
  return {
    ...paragraph,
    lines: paragraph.lines.map((line) => ({
      ...line,
      placements: line.placements.map((placement) => {
        if (placement.kind !== 'drawing') return placement;
        const normalized = normalizedById.get(placement.drawingId);
        return normalized ? { ...placement, bounds: normalized.drawing.inkBounds } : placement;
      }),
    })),
    drawings,
    textBoxes,
    exclusions: paragraph.exclusions.map((exclusion) => {
      const normalized = exclusion.anchorOccurrenceId
        ? normalizedByOccurrence.get(exclusion.anchorOccurrenceId)
        : undefined;
      if (!normalized) return exclusion;
      const wrap = normalized.frame.geometry.wrap;
      const bounds = normalized.frame.geometry.wrapBounds;
      if (!bounds) return {
        ...exclusion,
        bounds: reframeRect(exclusion.bounds, normalized.reframe),
        polygon: exclusion.polygon.map((point) => reframePoint(point, normalized.reframe)),
      };
      return {
        ...exclusion,
        wrap: wrap.kind === 'none' ? exclusion.wrap : wrap.kind,
        bounds,
        polygon: wrap.polygon?.points ?? [
          { xPt: bounds.xPt, yPt: bounds.yPt },
          { xPt: bounds.xPt + bounds.widthPt, yPt: bounds.yPt },
          { xPt: bounds.xPt + bounds.widthPt, yPt: bounds.yPt + bounds.heightPt },
          { xPt: bounds.xPt, yPt: bounds.yPt + bounds.heightPt },
        ],
      };
    }),
    ...(paragraph.anchorFrames ? {
      anchorFrames: paragraph.anchorFrames.map((frame) => (
        normalizedByOccurrence.get(frame.occurrenceId)?.frame ?? frame
      )),
    } : {}),
  };
}

function childPlacement(
  child: ParagraphLayout | TableLayout,
  cell: TableLayout['rows'][number]['cells'][number],
  offsetPt: number,
): Readonly<{ xPt: number; yPt: number }> {
  return {
    xPt: cell.contentBounds.xPt + (child.kind === 'table' ? child.flowBounds.xPt : 0),
    yPt: cell.flowBounds.yPt + offsetPt + (child.kind === 'table' ? child.flowBounds.yPt : 0),
  };
}

function normalizeTable(
  table: TableLayout,
  context: PageAnchorNormalizationContext,
): TableLayout {
  const normalized: TableLayout = {
    ...table,
    rows: table.rows.map((row) => ({
      ...row,
      cells: row.cells.map((cell) => ({
        ...cell,
        blocks: cell.blocks.map((block) => {
          const target = childPlacement(block.layout, cell, block.offsetPt);
          const delta = {
            xPt: target.xPt - block.layout.flowBounds.xPt,
            yPt: target.yPt - block.layout.flowBounds.yPt,
          };
          const currentToPage = composeAffine(
            context.currentToPage,
            translationAffine(delta.xPt, delta.yPt),
          );
          return {
            ...block,
            layout: block.layout.kind === 'paragraph'
              ? normalizeParagraph(block.layout, { ...context, currentToPage })
              : normalizeTable(block.layout, { ...context, currentToPage }),
          };
        }),
      })),
    })),
  };
  if (!('paintReadyFloatingTables' in table)) return normalized;
  const paintReady = table as PaintReadyTableLayout;
  if (paintReady.paintReadyFloatingTables.kind !== 'resolved') {
    return { ...normalized, paintReadyFloatingTables: paintReady.paintReadyFloatingTables } as PaintReadyTableLayout;
  }
  const placements = paintReady.paintReadyFloatingTables.placements.map((placement) => {
    const delta = {
      xPt: placement.xPt - placement.child.flowBounds.xPt,
      yPt: placement.yPt - placement.child.flowBounds.yPt,
    };
    const currentToPage = composeAffine(
      context.currentToPage,
      translationAffine(delta.xPt, delta.yPt),
    );
    const child = normalizeTable(placement.child, { ...context, currentToPage });
    return {
      ...placement,
      child,
      source: { ...placement.source, child },
    };
  });
  return {
    ...normalized,
    paintReadyFloatingTables: {
      ...paintReady.paintReadyFloatingTables,
      placements,
    },
  } as PaintReadyTableLayout;
}

/** Normalize the acquired/projected hybrid anchor graph exactly once before a
 * node is admitted to `PagePaintNode`. This adapter owns the destination page,
 * column, and region matrix; Canvas never reconstructs them from ownership. */
export function normalizePagePaintNodeAnchors<T extends PaintNode>(
  node: T,
  context: PageAnchorNormalizationContext,
): T {
  const normalized = node.kind === 'paragraph'
    ? normalizeParagraph(node, context)
    : node.kind === 'table'
      ? normalizeTable(node, context)
      : node.kind === 'textbox'
        ? normalizeTextBox(node, context)
        : node;
  return snapshotPlainData(normalized, 'Normalized page anchor graph') as T;
}

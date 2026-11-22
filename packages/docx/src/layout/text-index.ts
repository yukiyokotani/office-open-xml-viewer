import { sourceOwnedTextPlacements, type SourceOwnedTextGeometry } from './text-source-ownership.js';
import {
  composeAffine,
  inverseMapAffinePoint,
  mapAffinePoint,
  quarterTurnAffine,
  translationAffine,
} from './affine.js';
import { sourceKey } from './source-key.js';
import { lineFrameTransform } from './line-frame.js';
import type {
  DocumentLayout,
  DrawingLayout,
  LayoutPage,
  LayoutRect,
  Matrix2DData,
  PageLayerRoot,
  PagePaintDrawingEntry,
  PaintNode,
  ParagraphLayout,
  PointPt,
  RasterPaintOccurrence,
  ResourcePlacement,
  SourceRef,
  TableLayout,
  TextBoxLayout,
  UprightResourceOrientation,
} from './types.js';

export interface TextRunGeometry {
  readonly placement: SourceOwnedTextGeometry;
  readonly pointToPage: Matrix2DData;
  /** The part of the run's local frame (the frame of `placement.bounds`)
   * that its ancestor paint clips leave visible: their intersection where a
   * clip is axis-aligned with the run, and the bounding box of the clipped run
   * box where it is not. Absent for an unclipped run. */
  readonly visibleBounds?: LayoutRect;
  /** Canonical structural source of the owning paragraph. */
  readonly source: ParagraphLayout['source'];
  /** Source `w14:paraId` of the owning paragraph, when authored. */
  readonly paragraphId?: string;
}

/** Retained rectangular Canvas clip in the coordinate space where it is applied. */
export interface ElementClipGeometry {
  readonly bounds: LayoutRect;
  readonly pointToPage: Matrix2DData;
}

/** Retained DrawingML occurrence projected into physical page points. */
export interface DrawingGeometry {
  readonly drawing: DrawingLayout;
  readonly textBoxes: readonly TextBoxLayout[];
  readonly pointToPage: Matrix2DData;
  /** Ancestor paint clips, in their retained local coordinate spaces. */
  readonly clips: readonly ElementClipGeometry[];
  /** Index of the page paint entry that owns this occurrence. */
  readonly paintOrderIndex: number;
  /** Stable tie-breaker for inline drawings inside the same paint entry. */
  readonly sourceOrder: number;
}

/** Inline image/chart occurrence projected into physical page points. */
export interface InlineResourceGeometry {
  readonly placement: ResourcePlacement & Readonly<{ resourceKind: 'image' | 'chart' }>;
  readonly source: SourceRef;
  readonly pointToPage: Matrix2DData;
  /** Ancestor paint clips, in their retained local coordinate spaces. */
  readonly clips: readonly ElementClipGeometry[];
  /** Index of the page paint entry that owns the paragraph occurrence. */
  readonly paintOrderIndex: number;
  /** Stable tie-breaker within the owning paint entry. */
  readonly sourceOrder: number;
}

export type ElementGeometry = DrawingGeometry | InlineResourceGeometry;

interface ProjectionContext {
  readonly collectTextRuns: boolean;
  readonly collectTextRunSources: boolean;
  readonly collectCompletedParagraphSources: boolean;
  readonly collectDrawings: boolean;
  readonly collectRasterPaintOccurrences: boolean;
  readonly drawingEntries: ReadonlyMap<string, PagePaintDrawingEntry>;
  readonly rootPointToPage: ReadonlyMap<string, Matrix2DData>;
  readonly rootPaintOrder: ReadonlyMap<string, number>;
  readonly drawingPaintOrder: ReadonlyMap<string, number>;
  readonly emittedTextBoxes: Set<string>;
  readonly emittedDrawings: Set<string>;
  readonly runs: TextRunGeometry[];
  readonly sourceRuns: Map<string, Set<number>>;
  readonly completedParagraphSources: Set<string>;
  readonly drawings: DrawingGeometry[];
  readonly inlineResources: InlineResourceGeometry[];
  readonly rasterPaintOccurrences: RasterPaintOccurrence[];
  readonly emittedRasterDrawings: Set<string>;
  drawingSourceOrder: number;
}

interface NodeProjection {
  readonly pointToPage: Matrix2DData;
  readonly layoutTranslationPt: PointPt;
  readonly rootNodeId: string;
  readonly paintOrderIndex: number;
  readonly clips: readonly ElementClipGeometry[];
  readonly textBoxVerticalMode?: NonNullable<TextBoxLayout['verticalMode']>;
  /** Projected inside a line's retained frame; nested stories do not apply
   * their own (paint/types.ts CanvasPaintContext.lineFramed). */
  readonly lineFramed?: true;
  /** WORD_CELL_CONTENT_CLIP: the clip of the cell owning this paragraph. It
   * applies to the paragraph's text and inline content, not to its VML
   * pictures, anchored drawings or nested tables, as paint applies it. */
  readonly cellContentClip?: ElementClipGeometry;
}

const IDENTITY_AFFINE = Object.freeze({
  a: 1, b: 0, c: 0, d: 1, e: 0, f: 0,
}) satisfies Matrix2DData;

const EMPTY_CLIPS = Object.freeze([]) satisfies readonly ElementClipGeometry[];

const CLOCKWISE_QUARTER_TURN = quarterTurnAffine(1);
const COUNTER_CLOCKWISE_QUARTER_TURN = quarterTurnAffine(-1);
const TRANSPOSE: Matrix2DData = Object.freeze({ a: 0, b: 1, c: 1, d: 0, e: 0, f: 0 });

function appendRasterPaintOccurrence(
  context: ProjectionContext,
  resourceKey: string,
  resourceKind: 'image' | 'chart' | 'picture-bullet',
  bounds: LayoutRect,
  pointToPage: Matrix2DData,
  options: Readonly<{
    orientation?: UprightResourceOrientation;
    textBoxVerticalMode?: NonNullable<TextBoxLayout['verticalMode']>;
  }> = {},
): void {
  if (!context.collectRasterPaintOccurrences) return;
  let matrix = pointToPage;
  let localWidthPt = bounds.widthPt;
  let localHeightPt = bounds.heightPt;
  if (options.textBoxVerticalMode) {
    matrix = composeAffine(
      matrix,
      options.textBoxVerticalMode === 'vert270'
        ? CLOCKWISE_QUARTER_TURN
        : COUNTER_CLOCKWISE_QUARTER_TURN,
    );
    [localWidthPt, localHeightPt] = [localHeightPt, localWidthPt];
  }
  if (options.orientation !== undefined) {
    // The retained local turn inverts the owner section's quarter turn.
    matrix = composeAffine(
      matrix,
      options.orientation === 'upright-physical'
        ? COUNTER_CLOCKWISE_QUARTER_TURN
        : options.orientation === 'upright-physical-transposed'
          ? TRANSPOSE
          : CLOCKWISE_QUARTER_TURN,
    );
    [localWidthPt, localHeightPt] = [localHeightPt, localWidthPt];
  }
  const widthPt = localWidthPt * Math.hypot(matrix.a, matrix.b);
  const heightPt = localHeightPt * Math.hypot(matrix.c, matrix.d);
  // The painters do no visible work for a non-positive destination. Invalid
  // retained numbers are caught by the document-layout invariant boundary.
  if (!Number.isFinite(widthPt) || !Number.isFinite(heightPt)
    || !(widthPt > 0) || !(heightPt > 0)) return;
  context.rasterPaintOccurrences.push(Object.freeze({
    resourceKey,
    resourceKind,
    widthPt,
    heightPt,
  }));
}

function collectDrawingRasterPaintOccurrences(
  drawing: DrawingLayout,
  projection: NodeProjection,
  context: ProjectionContext,
): void {
  if (!context.collectRasterPaintOccurrences
    || context.emittedRasterDrawings.has(drawing.id)) return;
  context.emittedRasterDrawings.add(drawing.id);
  for (const command of drawing.commands) {
    if (command.kind === 'resource' && command.resourceKind !== 'math') {
      appendRasterPaintOccurrence(
        context,
        command.resourceKey,
        command.resourceKind,
        command.rect,
        projection.pointToPage,
        { orientation: command.orientation },
      );
      continue;
    }
    if (command.kind === 'watermark-text' && command.imageResourceKey) {
      appendRasterPaintOccurrence(context, command.imageResourceKey, 'image', command.maskBounds ?? command.rect, projection.pointToPage);
    }
    if (command.kind !== 'drawingml-image-fill') continue;
    // Part 4 §§19.1.2.12/.21: a physical textured pen has ink even on
    // a zero-height/width line. Decode against its acquired mask extent;
    // allocation-only demand would silently classify this visible pen as empty.
    if (command.strokeMask && command.maskBounds) {
      appendRasterPaintOccurrence(context, command.resourceKey, 'image', command.maskBounds, projection.pointToPage);
      continue;
    }
    const { x, y, w, h } = command.plan.rect;
    const fillRect = command.fillRect ?? { l: 0, t: 0, r: 0, b: 0 };
    appendRasterPaintOccurrence(context, command.resourceKey, 'image', {
      xPt: x + fillRect.l * w,
      yPt: y + fillRect.t * h,
      widthPt: w * (1 - fillRect.l - fillRect.r),
      heightPt: h * (1 - fillRect.t - fillRect.b),
    }, projection.pointToPage);
  }
}

function pageRegionsByDomain(
  page: LayoutPage,
): ReadonlyMap<string, LayoutPage['sectionRegions'][number]> {
  const byId = new Map(page.sectionRegions.map((region) => [region.id, region]));
  const byDomain = new Map<string, LayoutPage['sectionRegions'][number]>();
  for (const region of page.sectionRegions) {
    for (const flowDomainId of region.flowDomainIds) {
      byDomain.set(flowDomainId, region);
    }
  }
  for (const domain of page.flowDomains) {
    if (domain.kind !== 'footnote' && domain.kind !== 'endnote'
      && domain.sectionRegionId === undefined) continue;
    const storyRegion = domain.sectionRegionId
      ? byId.get(domain.sectionRegionId)
      : page.sectionRegions[0];
    if (!storyRegion) {
      throw new Error(
        `${domain.id} references missing page story region ${domain.sectionRegionId ?? '<default>'}`,
      );
    }
    byDomain.set(domain.id, storyRegion);
  }
  return byDomain;
}

function pointToPageForRoot(
  regionByDomain: ReadonlyMap<string, LayoutPage['sectionRegions'][number]>,
  root: Pick<PageLayerRoot, 'coordinateSpace' | 'node'>,
): Matrix2DData {
  if (root.coordinateSpace === 'upright-physical') return IDENTITY_AFFINE;
  const matrix = regionByDomain.get(root.node.flowDomainId)
    ?.coordinateSpace.logicalToPhysical;
  return matrix ?? IDENTITY_AFFINE;
}

function projectionForEntry(
  context: ProjectionContext,
  entry: PagePaintDrawingEntry,
): NodeProjection {
  const rootPointToPage = context.rootPointToPage.get(entry.rootNodeId);
  if (!rootPointToPage) {
    throw new Error(`Drawing entry ${entry.node.id} references missing root ${entry.rootNodeId}`);
  }
  let pointToPage = rootPointToPage;
  const clips: ElementClipGeometry[] = [];
  for (const frame of entry.frames) {
    if (frame.kind === 'transform') {
      pointToPage = composeAffine(pointToPage, frame.transform);
    } else {
      clips.push(Object.freeze({ bounds: frame.clip, pointToPage }));
    }
  }
  return {
    pointToPage,
    layoutTranslationPt: entry.layoutTranslationPt,
    rootNodeId: entry.rootNodeId,
    paintOrderIndex: context.drawingPaintOrder.get(entry.node.id) ?? -1,
    clips: Object.freeze(clips),
  };
}

function withClip(
  projection: NodeProjection,
  bounds: LayoutRect | undefined,
): NodeProjection {
  if (!bounds) return projection;
  return {
    ...projection,
    clips: Object.freeze([
      ...projection.clips,
      Object.freeze({ bounds, pointToPage: projection.pointToPage }),
    ]),
  };
}

function placedChildProjection(
  child: ParagraphLayout | TableLayout,
  placement: Readonly<{ xPt: number; yPt: number }>,
  parent: NodeProjection,
): NodeProjection {
  const dxPt = placement.xPt - child.flowBounds.xPt;
  const dyPt = placement.yPt - child.flowBounds.yPt;
  return {
    ...parent,
    pointToPage: composeAffine(
      parent.pointToPage,
      translationAffine(dxPt, dyPt),
    ),
    layoutTranslationPt: {
      xPt: parent.layoutTranslationPt.xPt + dxPt,
      yPt: parent.layoutTranslationPt.yPt + dyPt,
    },
  };
}

function drawingTextBoxes(
  byId: ReadonlyMap<string, TextBoxLayout>,
  drawing: DrawingLayout,
): readonly TextBoxLayout[] {
  return (drawing.textBoxIds ?? []).flatMap((id) => {
    const textBox = byId.get(id);
    return textBox ? [textBox] : [];
  });
}

function visitTextBox(
  textBox: TextBoxLayout,
  projection: NodeProjection,
  context: ProjectionContext,
): void {
  if (context.emittedTextBoxes.has(textBox.id)) return;
  context.emittedTextBoxes.add(textBox.id);
  // A text box story is its own coordinate root: its resolved floating-table
  // placements and the page-owned axes of drawings anchored in it are stated
  // in story coordinates (story-page-frames.ts), which `pointToPage` already
  // maps. Only translations taken inside the story are undone for them, never
  // those that placed the box's own paragraph (a table cell's).
  const transformedProjection: NodeProjection = {
    ...projection,
    pointToPage: composeAffine(projection.pointToPage, textBox.transform),
    layoutTranslationPt: { xPt: 0, yPt: 0 },
    textBoxVerticalMode: textBox.verticalMode ?? projection.textBoxVerticalMode,
  };
  const textBoxProjection = withClip(transformedProjection, textBox.clipBounds);
  for (const block of textBox.story.blocks) {
    visitNode(block, textBoxProjection, context);
  }
}

function visitDrawing(
  textBoxesById: ReadonlyMap<string, TextBoxLayout>,
  drawing: DrawingLayout,
  projection: NodeProjection,
  context: ProjectionContext,
): void {
  const textBoxes = drawingTextBoxes(textBoxesById, drawing);
  const ownedProjection = drawingOwnedContentProjection(drawing, projection, context);
  collectDrawingRasterPaintOccurrences(drawing, ownedProjection, context);
  if (context.collectDrawings && !context.emittedDrawings.has(drawing.id)) {
    context.emittedDrawings.add(drawing.id);
    context.drawings.push(Object.freeze({
      drawing,
      textBoxes,
      pointToPage: ownedProjection.pointToPage,
      clips: ownedProjection.clips,
      paintOrderIndex: ownedProjection.paintOrderIndex,
      sourceOrder: context.drawingSourceOrder++,
    }));
  }
  for (const textBox of textBoxes) {
    visitTextBox(textBox, ownedProjection, context);
  }
}

function drawingOwnedContentProjection(
  drawing: DrawingLayout,
  projection: NodeProjection,
  context: ProjectionContext,
): NodeProjection {
  const retainedEntry = context.drawingEntries.get(drawing.id);
  let drawingProjection = projection;
  if (
    retainedEntry
    && retainedEntry.rootNodeId === projection.rootNodeId
  ) {
    drawingProjection = projectionForEntry(context, retainedEntry);
  }
  const translation = drawingProjection.layoutTranslationPt;
  const undoX = drawing.anchorLayer?.horizontalOwnership === 'page'
    ? -translation.xPt : 0;
  const undoY = drawing.anchorLayer?.verticalOwnership === 'page'
    ? -translation.yPt : 0;
  let ownedProjection = undoX === 0 && undoY === 0
    ? drawingProjection
    : {
        ...drawingProjection,
        pointToPage: composeAffine(
          drawingProjection.pointToPage,
          translationAffine(undoX, undoY),
        ),
      };
  if (drawing.orientation === 'upright-physical') {
    if (!drawing.transform) {
      throw new Error(`Upright physical drawing ${drawing.id} is missing its logical transform`);
    }
    ownedProjection = {
      ...ownedProjection,
      pointToPage: composeAffine(ownedProjection.pointToPage, drawing.transform),
    };
  }
  return ownedProjection;
}

/** A line's point-to-page matrix: its paragraph's, composed with the line's
 * own paint frame (line-frame.ts) as paint composes it. */
function lineProjectionPointToPage(
  paragraph: ParagraphLayout,
  line: ParagraphLayout['lines'][number],
  projection: NodeProjection,
): Matrix2DData {
  const lineFrame = projection.lineFramed ? null : lineFrameTransform(paragraph, line);
  return lineFrame ? composeAffine(projection.pointToPage, lineFrame) : projection.pointToPage;
}

const CLIP_ALIGNMENT_EPSILON_PT = 1e-6;

/** TextRunGeometry.visibleBounds of a run box under `clips`. */
function visibleBoundsEntry(
  bounds: LayoutRect,
  runPointToPage: Matrix2DData,
  clips: readonly ElementClipGeometry[],
): Readonly<{ visibleBounds?: LayoutRect }> {
  if (clips.length === 0) return {};
  let left = -Infinity;
  let top = -Infinity;
  let right = Infinity;
  let bottom = Infinity;
  for (const clip of clips) {
    const { xPt, yPt, widthPt, heightPt } = clip.bounds;
    const corners = [
      { xPt, yPt }, { xPt: xPt + widthPt, yPt }, { xPt: xPt + widthPt, yPt: yPt + heightPt }, { xPt, yPt: yPt + heightPt },
    ].map((corner) => inverseMapAffinePoint(runPointToPage, mapAffinePoint(clip.pointToPage, corner)));
    if (corners.some((corner) => corner === null)) return {};
    const local = corners as PointPt[];
    const xs = local.map((point) => point.xPt);
    const ys = local.map((point) => point.yPt);
    const aligned = local.every((point, index) => {
      const next = local[(index + 1) % 4]!;
      return Math.abs(point.xPt - next.xPt) < CLIP_ALIGNMENT_EPSILON_PT
        || Math.abs(point.yPt - next.yPt) < CLIP_ALIGNMENT_EPSILON_PT;
    });
    const region = aligned ? {
      left: Math.min(...xs), top: Math.min(...ys), right: Math.max(...xs), bottom: Math.max(...ys),
    } : clippedRunBoxExtent(bounds, local);
    left = Math.max(left, region.left);
    top = Math.max(top, region.top);
    right = Math.min(right, region.right);
    bottom = Math.min(bottom, region.bottom);
  }
  const xPt = Math.min(left, right);
  const yPt = Math.min(top, bottom);
  return { visibleBounds: Object.freeze({ xPt, yPt, widthPt: Math.max(0, right - xPt), heightPt: Math.max(0, bottom - yPt) }) };
}

/** The extent of `box` cut by the convex quadrilateral `polygon` (a clip
 * turned against the run); an empty extent when nothing remains. */
function clippedRunBoxExtent(
  box: LayoutRect,
  polygon: readonly PointPt[],
): Readonly<{ left: number; top: number; right: number; bottom: number }> {
  let points: PointPt[] = [
    { xPt: box.xPt, yPt: box.yPt }, { xPt: box.xPt + box.widthPt, yPt: box.yPt },
    { xPt: box.xPt + box.widthPt, yPt: box.yPt + box.heightPt }, { xPt: box.xPt, yPt: box.yPt + box.heightPt },
  ];
  // Orient the clip edges so that the inside is on the left of each edge.
  const area = polygon.reduce((sum, point, index) => {
    const next = polygon[(index + 1) % polygon.length]!;
    return sum + point.xPt * next.yPt - next.xPt * point.yPt;
  }, 0);
  const sign = area >= 0 ? 1 : -1;
  for (let index = 0; index < polygon.length && points.length > 0; index++) {
    const a = polygon[index]!;
    const b = polygon[(index + 1) % polygon.length]!;
    const side = (point: PointPt) => sign * ((b.xPt - a.xPt) * (point.yPt - a.yPt) - (b.yPt - a.yPt) * (point.xPt - a.xPt));
    const next: PointPt[] = [];
    for (let i = 0; i < points.length; i++) {
      const current = points[i]!;
      const following = points[(i + 1) % points.length]!;
      const sc = side(current);
      const sf = side(following);
      if (sc >= 0) next.push(current);
      if ((sc >= 0) !== (sf >= 0)) {
        const t = sc / (sc - sf);
        next.push({
          xPt: current.xPt + t * (following.xPt - current.xPt),
          yPt: current.yPt + t * (following.yPt - current.yPt),
        });
      }
    }
    points = next;
  }
  if (points.length === 0) return { left: box.xPt, top: box.yPt, right: box.xPt, bottom: box.yPt };
  const xs = points.map((point) => point.xPt);
  const ys = points.map((point) => point.yPt);
  return { left: Math.min(...xs), top: Math.min(...ys), right: Math.max(...xs), bottom: Math.max(...ys) };
}

function visitParagraph(
  paragraph: ParagraphLayout,
  projection: NodeProjection,
  context: ProjectionContext,
): void {
  const { cellContentClip, ...unclippedProjection } = withClip(projection, paragraph.clipBounds);
  const paragraphProjection: NodeProjection = unclippedProjection;
  // WORD_CELL_CONTENT_CLIP: the projection of content painted under the cell clip.
  const clippedProjection: NodeProjection = cellContentClip
    ? { ...paragraphProjection, clips: Object.freeze([...paragraphProjection.clips, cellContentClip]) }
    : paragraphProjection;
  const escapingDrawingIds = new Set(cellContentClip ? paragraph.lines.flatMap((line) => line.placements.flatMap(
    (placement) => placement.kind === 'drawing' && placement.escapesCellClip ? [placement.drawingId] : [],
  )) : []);
  // An inline drawing is line content: it projects through its owning
  // line's frame, as paint draws it (line-frame.ts).
  const inlineDrawingLines = new Map<string, ParagraphLayout['lines'][number]>();
  if (paragraph.lineFrame !== undefined) {
    for (const line of paragraph.lines) {
      for (const placement of line.placements) {
        if (placement.kind === 'drawing') inlineDrawingLines.set(placement.drawingId, line);
      }
    }
  }
  const inlineDrawingProjection = (drawing: DrawingLayout): NodeProjection => {
    const line = drawing.anchorLayer ? undefined : inlineDrawingLines.get(drawing.id);
    const base = drawing.anchorLayer || escapingDrawingIds.has(drawing.id)
      ? paragraphProjection : clippedProjection;
    return line && !base.lineFramed && paragraph.lineFrame !== undefined
      ? {
          ...base,
          lineFramed: true,
          pointToPage: lineProjectionPointToPage(paragraph, line, base),
        }
      : base;
  };
  const textBoxesById = new Map(
    paragraph.textBoxes.map((textBox) => [textBox.id, textBox]),
  );
  if (
    context.collectCompletedParagraphSources
    && paragraph.continuation?.continuesOnNext !== true
  ) {
    context.completedParagraphSources.add(sourceKey(paragraph.source));
  }
  if (context.collectTextRuns || context.collectTextRunSources) {
    const inlineTextDrawings = new Map(paragraph.drawings
      .filter((drawing) => !drawing.anchorLayer && (drawing.textBoxIds?.length ?? 0) > 0)
      .map((drawing) => [drawing.id, drawing]));
    for (const line of paragraph.lines) {
      const linePointToPage = lineProjectionPointToPage(paragraph, line, projection);
      for (const retained of line.placements) {
        // ECMA-376 §20.4.2.8: an inline drawing occupies its run's place in
        // paragraph flow. Library selection policy includes nested text there
        // for copy; floating stories keep their separate reading-order policy.
        if (retained.kind === 'drawing') {
          const drawing = inlineTextDrawings.get(retained.drawingId);
          if (drawing) visitDrawing(textBoxesById, drawing, inlineDrawingProjection(drawing), context);
        }
        if (retained.kind === 'text') for (const placement of sourceOwnedTextPlacements(retained)) {
          if (context.collectTextRuns) {
            context.runs.push(Object.freeze({
              placement,
              pointToPage: linePointToPage,
              ...visibleBoundsEntry(placement.bounds, linePointToPage, clippedProjection.clips),
              source: paragraph.source,
              ...(paragraph.paragraphId !== undefined
                ? { paragraphId: paragraph.paragraphId }
                : {}),
            }));
          }
          if (
            context.collectTextRunSources
            && placement.sourceRunIndex !== undefined
            && placement.text.length > 0
          ) {
            const key = sourceKey(paragraph.source);
            const indices = context.sourceRuns.get(key) ?? new Set<number>();
            if (!context.sourceRuns.has(key)) context.sourceRuns.set(key, indices);
            indices.add(placement.sourceRunIndex);
          }
        }
      }
    }
  }
  if (context.collectRasterPaintOccurrences) {
    for (const line of paragraph.lines) {
      const linePointToPage = lineProjectionPointToPage(paragraph, line, paragraphProjection);
      for (const placement of line.placements) {
        if (placement.kind !== 'resource' || placement.resourceKind === 'math') continue;
        appendRasterPaintOccurrence(
          context,
          placement.resourceKey,
          placement.resourceKind,
          placement.bounds,
          linePointToPage,
          {
            orientation: placement.orientation,
            textBoxVerticalMode: paragraphProjection.textBoxVerticalMode,
          },
        );
      }
    }
  }

  const ownedTextBoxIds = new Set<string>();
  const drawingsInSourceOrder = paragraph.drawings
    .map((drawing, index) => {
      const runIndex = drawing.source.path.at(-1);
      if (runIndex === undefined || !Number.isSafeInteger(runIndex) || runIndex < 0) {
        throw new Error(`Drawing ${drawing.id} has no retained paragraph run index`);
      }
      return { drawing, index, runIndex };
    })
    .sort((left, right) => left.runIndex - right.runIndex || left.index - right.index);
  const inlineResourcesInSourceOrder = context.collectDrawings
    ? paragraph.lines.flatMap((line) =>
        line.placements.flatMap((placement, index) => {
          if (placement.kind !== 'resource' ||
            (placement.resourceKind !== 'image' && placement.resourceKind !== 'chart') ||
            placement.sourceRunIndex === undefined) return [];
          return [{
            placement: placement as ResourcePlacement & Readonly<{ resourceKind: 'image' | 'chart' }>,
            index,
            runIndex: placement.sourceRunIndex,
            pointToPage: lineProjectionPointToPage(paragraph, line, paragraphProjection),
          }];
        }))
    : [];
  // Both anchored and inline drawings retain their paragraph run index as the
  // terminal SourceRef path component. That is the one comparable source-order
  // domain; anchor stacking ordinals and drawings-array indexes are not mixed.
  for (const { drawing } of drawingsInSourceOrder) {
    for (const id of drawing.textBoxIds ?? []) ownedTextBoxIds.add(id);
  }
  const elementsInSourceOrder = [
    ...drawingsInSourceOrder.map((entry) => ({ kind: 'drawing' as const, ...entry })),
    ...inlineResourcesInSourceOrder.map((entry) => ({ kind: 'resource' as const, ...entry })),
  ].sort((left, right) => left.runIndex - right.runIndex || left.index - right.index);
  for (const entry of elementsInSourceOrder) {
    if (entry.kind === 'drawing') {
      visitDrawing(textBoxesById, entry.drawing, inlineDrawingProjection(entry.drawing), context);
      continue;
    }
    if (!context.collectDrawings) continue;
    context.inlineResources.push(Object.freeze({
      placement: entry.placement,
      source: Object.freeze({
        ...paragraph.source,
        path: Object.freeze([...paragraph.source.path, entry.runIndex]),
      }),
      pointToPage: entry.pointToPage,
      clips: entry.placement.escapesCellClip ? paragraphProjection.clips : clippedProjection.clips,
      paintOrderIndex: paragraphProjection.paintOrderIndex,
      sourceOrder: context.drawingSourceOrder++,
    }));
  }
  for (const textBox of paragraph.textBoxes) {
    if (!ownedTextBoxIds.has(textBox.id)) {
      visitTextBox(textBox, clippedProjection, context);
    }
  }
}

function visitTable(
  table: TableLayout,
  projection: NodeProjection,
  context: ProjectionContext,
): void {
  const tableProjection = withClip(projection, table.clipBounds);
  for (const row of table.rows) {
    for (const cell of row.cells) {
      const ownsContinuationPaint = cell.visualMergeOwnership === 'continuation';
      if (cell.verticalMerge === 'continue' && !ownsContinuationPaint) continue;
      const cellProjection = withClip(tableProjection, cell.clipBounds);
      if (cell.verticalText) {
        // ECMA-376 §17.4.72: map the rotated local frame into the table.
        const rotatedProjection: NodeProjection = {
          ...cellProjection,
          pointToPage: composeAffine(cellProjection.pointToPage, cell.verticalText.transform),
          textBoxVerticalMode: cell.verticalText.mode,
        };
        for (const block of cell.blocks) {
          visitNode(block.layout, placedChildProjection(block.layout, {
            xPt: cell.contentBounds.xPt,
            yPt: block.offsetPt,
          }, rotatedProjection), context);
        }
        continue;
      }
      // WORD_CELL_CONTENT_CLIP: the cell's paragraphs, not its nested tables.
      const paragraphProjection: NodeProjection = cell.contentClipBounds
        ? {
            ...cellProjection,
            cellContentClip: Object.freeze({
              bounds: cell.contentClipBounds,
              pointToPage: cellProjection.pointToPage,
            }),
          }
        : cellProjection;
      for (const block of cell.blocks) {
        const child = block.layout;
        visitNode(child, placedChildProjection(child, {
          xPt: cell.contentBounds.xPt
            + (child.kind === 'table' ? child.flowBounds.xPt : 0),
          yPt: cell.flowBounds.yPt + block.offsetPt
            + (child.kind === 'table' ? child.flowBounds.yPt : 0),
        }, child.kind === 'paragraph' ? paragraphProjection : cellProjection), context);
      }
    }
  }
  for (const placement of table.resolvedFloatingTables ?? []) {
    visitNode(placement.child, placedChildProjection(placement.child, {
      xPt: placement.xPt - projection.layoutTranslationPt.xPt,
      yPt: placement.yPt - projection.layoutTranslationPt.yPt,
    }, tableProjection), context);
  }
}

function visitNode(
  node: PaintNode,
  projection: NodeProjection,
  context: ProjectionContext,
): void {
  switch (node.kind) {
    case 'paragraph':
      visitParagraph(node, projection, context);
      return;
    case 'table':
      visitTable(node, projection, context);
      return;
    case 'note': {
      // Separator/notice paragraphs carry only zero-advance participants, so
      // they add source ownership but no selectable text.
      if (node.leading?.paragraph) visitNode(node.leading.paragraph, projection, context);
      for (const block of node.story.blocks) {
        visitNode(block, withClip(projection, node.story.clipBounds), context);
      }
      if (node.trailing?.paragraph) visitNode(node.trailing.paragraph, projection, context);
      return;
    }
    case 'textbox':
      visitTextBox(node, projection, context);
      return;
    case 'drawing': {
      const entry = context.drawingEntries.get(node.id);
      visitDrawing(
        new Map((entry?.textBoxes ?? []).map((textBox) => [textBox.id, textBox])),
        node,
        projection,
        context,
      );
      return;
    }
    default: {
      const exhaustive: never = node;
      throw new Error(`Unknown text-index node: ${String(exhaustive)}`);
    }
  }
}

/**
 * Indexes retained text placements in physical page points. Sequence follows
 * semantic reading order; paint order contributes only already-materialized
 * anchor frame geometry.
 */
function pageGeometryIndex(
  layout: DocumentLayout,
  pageIndex: number,
  options: Readonly<{
    collectTextRuns: boolean;
    collectTextRunSources: boolean;
    collectCompletedParagraphSources?: boolean;
    collectDrawings: boolean;
    collectRasterPaintOccurrences?: boolean;
  }>,
): ProjectionContext {
  const page = layout.pages[pageIndex];
  if (!page) throw new RangeError(`Page index ${pageIndex} is out of range`);
  const roots = new Map(page.layers.roots.map((root) => [root.node.id, root]));
  const regionByDomain = pageRegionsByDomain(page);
  const rootPointToPage = new Map(page.layers.roots.map((root) => [
    root.node.id,
    pointToPageForRoot(regionByDomain, root),
  ]));
  const drawingEntries = new Map<string, PagePaintDrawingEntry>();
  const drawingPaintOrder = new Map<string, number>();
  const rootPaintOrder = new Map<string, number>();
  for (const [index, entry] of page.layers.paintOrder.entries()) {
    if (entry.kind === 'drawing') drawingEntries.set(entry.node.id, entry);
    if (entry.kind === 'drawing') drawingPaintOrder.set(entry.node.id, index);
    else rootPaintOrder.set(entry.node.id, index);
  }
  const context: ProjectionContext = {
    ...options,
    collectCompletedParagraphSources: options.collectCompletedParagraphSources === true,
    collectRasterPaintOccurrences: options.collectRasterPaintOccurrences === true,
    drawingEntries,
    rootPointToPage,
    rootPaintOrder,
    drawingPaintOrder,
    emittedTextBoxes: new Set(),
    emittedDrawings: new Set(),
    runs: [],
    sourceRuns: new Map(),
    completedParagraphSources: new Set(),
    drawings: [],
    inlineResources: [],
    rasterPaintOccurrences: [],
    emittedRasterDrawings: new Set(),
    drawingSourceOrder: 0,
  };
  // ECMA-376 §17.2.1: the page background's picture brushes are painted in
  // physical page points before every page root, outside the story graph.
  if (context.collectRasterPaintOccurrences) {
    for (const command of page.pageBackground?.commands ?? []) {
      if (command.kind !== 'drawingml-image-fill') continue;
      const { x, y, w, h } = command.plan.rect;
      const fillRect = command.fillRect ?? { l: 0, t: 0, r: 0, b: 0 };
      appendRasterPaintOccurrence(context, command.resourceKey, 'image', {
        xPt: x + fillRect.l * w,
        yPt: y + fillRect.t * h,
        widthPt: w * (1 - fillRect.l - fillRect.r),
        heightPt: h * (1 - fillRect.t - fillRect.b),
      }, IDENTITY_AFFINE);
    }
  }
  const rootOrder = context.collectRasterPaintOccurrences
    ? page.layers.roots.map(({ node }) => node.id)
    : page.readingOrder;
  for (const nodeId of rootOrder) {
    const root = roots.get(nodeId);
    if (!root) throw new Error(`Reading-order node ${nodeId} is not a page root`);
    const pointToPage = rootPointToPage.get(nodeId);
    if (!pointToPage) throw new Error(`Reading-order node ${nodeId} has no page projection`);
    visitNode(root.node, {
      pointToPage,
      layoutTranslationPt: { xPt: 0, yPt: 0 },
      rootNodeId: root.node.id,
      paintOrderIndex: rootPaintOrder.get(root.node.id) ?? -1,
      clips: EMPTY_CLIPS,
    }, context);
  }
  return context;
}

export function textRunGeometryForPage(
  layout: DocumentLayout,
  pageIndex: number,
): readonly TextRunGeometry[] {
  return Object.freeze(pageGeometryIndex(layout, pageIndex, {
    collectTextRuns: true,
    collectTextRunSources: false,
    collectDrawings: false,
  }).runs);
}

/** Lightweight index of text placements that actually survived final-state
 * layout. It deliberately omits coordinates, font strings, and transforms;
 * comment-anchor fallback needs only canonical paragraph/run identity. */
export function textRunSourceIndexForDocument(
  layout: DocumentLayout,
): ReadonlyMap<string, ReadonlySet<number>> {
  const result = new Map<string, Set<number>>();
  for (let pageIndex = 0; pageIndex < layout.pages.length; pageIndex += 1) {
    const pageIndexResult = pageGeometryIndex(layout, pageIndex, {
      collectTextRuns: false,
      collectTextRunSources: true,
      collectDrawings: false,
    }).sourceRuns;
    for (const [key, pageIndices] of pageIndexResult) {
      const indices = result.get(key) ?? new Set<number>();
      if (!result.has(key)) result.set(key, indices);
      for (const index of pageIndices) indices.add(index);
    }
  }
  return result;
}

export interface ReviewProjectionIndex {
  readonly renderedRunIndex: ReadonlyMap<string, ReadonlySet<number>>;
  readonly completedSourceKeys: ReadonlySet<string>;
}

/** Build every index needed by comment/revision projection in one canonical
 * geometry traversal. Progressive publications call this only when review data
 * exists; combining the indexes avoids walking the growing prefix twice. */
export function reviewProjectionIndexForDocument(
  layout: DocumentLayout,
): ReviewProjectionIndex {
  const renderedRunIndex = new Map<string, Set<number>>();
  const completedSourceKeys = new Set<string>();
  for (let pageIndex = 0; pageIndex < layout.pages.length; pageIndex += 1) {
    const page = pageGeometryIndex(layout, pageIndex, {
      collectTextRuns: false,
      collectTextRunSources: true,
      collectCompletedParagraphSources: true,
      collectDrawings: false,
    });
    for (const [key, pageIndices] of page.sourceRuns) {
      const indices = renderedRunIndex.get(key) ?? new Set<number>();
      if (!renderedRunIndex.has(key)) renderedRunIndex.set(key, indices);
      for (const index of pageIndices) indices.add(index);
    }
    for (const key of page.completedParagraphSources) completedSourceKeys.add(key);
  }
  return Object.freeze({ renderedRunIndex, completedSourceKeys });
}

/** Index selectable drawings and inline image/chart placements in page paint order. */
export function elementGeometryForPage(
  layout: DocumentLayout,
  pageIndex: number,
): readonly ElementGeometry[] {
  const index = pageGeometryIndex(layout, pageIndex, {
    collectTextRuns: false,
    collectTextRunSources: false,
    collectDrawings: true,
  });
  const elements: ElementGeometry[] = [...index.drawings, ...index.inlineResources];
  elements.sort((left, right) => left.paintOrderIndex - right.paintOrderIndex ||
    left.sourceOrder - right.sourceOrder);
  return Object.freeze(elements);
}

/** Raster-backed resources that the completed page paint plan can reach. This
 * uses the same canonical projection as selection geometry, including page,
 * table, text-box, and anchored-drawing transforms, but intentionally ignores
 * clips so decode resolution remains conservative. */
export function rasterPaintOccurrencesForPage(
  layout: DocumentLayout,
  pageIndex: number,
): readonly RasterPaintOccurrence[] {
  return Object.freeze(pageGeometryIndex(layout, pageIndex, {
    collectTextRuns: false,
    collectTextRunSources: false,
    collectDrawings: false,
    collectRasterPaintOccurrences: true,
  }).rasterPaintOccurrences);
}

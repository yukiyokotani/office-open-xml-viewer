import { sourceOwnedTextPlacements, type SourceOwnedTextGeometry } from './text-source-ownership.js';
import {
  composeAffine,
  inverseMapAffinePoint,
  mapAffinePoint,
  quarterTurnAffine,
  translationAffine,
} from './affine.js';
import { sourceKey } from './source-key.js';
import type { ClusterBox } from './visual-cluster-geometry.js';
import { lineFrameTransform } from './line-frame.js';
import { resourcePlacementOutline } from './resource-outline.js';
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

/** Inline content of a line that is not text (a tab, an inline picture,
 * chart, equation or drawing): the box it occupies, addressed by its source
 * run like the text around it. */
export interface InlineContentGeometry {
  readonly inlineContent: 'tab' | 'object';
  readonly bounds: LayoutRect;
  readonly sourceRunIndex: number;
  readonly pointToPage: Matrix2DData;
  readonly clips: readonly ElementClipGeometry[];
  readonly source: ParagraphLayout['source'];
  readonly paragraphId?: string;
}

export interface TextRunGeometry {
  readonly placement: SourceOwnedTextGeometry;
  /** Where each cluster of the run is painted, from `placement.bounds.xPt`,
   * by offset in its text (visual-cluster-geometry.ts). */
  readonly clusterBoxes: readonly ClusterBox[];
  readonly pointToPage: Matrix2DData;
  /** Every ancestor paint clip of the run, outermost first, in its retained
   * local space (turned clips included): overlays clip with them. */
  readonly clips: readonly ElementClipGeometry[];
  /** Canonical structural source of the owning paragraph. */
  readonly source: ParagraphLayout['source'];
  /** Source `w14:paraId` of the owning paragraph, when authored. */
  readonly paragraphId?: string;
}

/** Retained rectangular Canvas clip in the coordinate space where it is applied. */
export interface ElementClipGeometry {
  readonly bounds: LayoutRect;
  readonly pointToPage: Matrix2DData;
  /** A clip that is a union of rectangles (a VML text box's textboxrect
   * regions), in the space of `bounds`, which is their bounding box. */
  readonly rects?: readonly LayoutRect[];
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
  readonly runs: (TextRunGeometry | InlineContentGeometry)[];
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
  rects?: readonly LayoutRect[],
): NodeProjection {
  if (rects) {
    // A union clip keeps only its rectangles with area; with none left it
    // clips everything, as the Canvas clip does.
    const regions = rects.filter((rect) => rect.widthPt > 0 && rect.heightPt > 0);
    if (regions.length === 0) {
      return {
        ...projection,
        clips: Object.freeze([
          ...projection.clips,
          Object.freeze({
            bounds: Object.freeze({ xPt: bounds?.xPt ?? 0, yPt: bounds?.yPt ?? 0, widthPt: 0, heightPt: 0 }),
            pointToPage: projection.pointToPage,
            rects: Object.freeze([]),
          }),
        ]),
      };
    }
    rects = regions;
    let left = Infinity;
    let top = Infinity;
    let right = -Infinity;
    let bottom = -Infinity;
    for (const rect of rects) {
      left = Math.min(left, rect.xPt);
      top = Math.min(top, rect.yPt);
      right = Math.max(right, rect.xPt + rect.widthPt);
      bottom = Math.max(bottom, rect.yPt + rect.heightPt);
    }
    return {
      ...projection,
      clips: Object.freeze([
        ...projection.clips,
        Object.freeze({
          bounds: Object.freeze({ xPt: left, yPt: top, widthPt: right - left, heightPt: bottom - top }),
          pointToPage: projection.pointToPage,
          rects: Object.freeze([...rects]),
        }),
      ]),
    };
  }
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
  // As paint clips the story: to the authored text rectangles when the box
  // has them, otherwise to its clip bounds.
  const textBoxProjection = withClip(transformedProjection, textBox.clipBounds, textBox.clipRects);
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

/** The box a tab or inline object is painted in on its line, in the line's
 * coordinates: its bounds, or for a picture or drawing the part its line clip
 * leaves. A turned picture is its visible outline cut by the clips it is
 * painted under (`clips`, with the line's `pointToPage`), then bounded: the
 * bounding box of the whole outline could reach into a clip the picture
 * itself never enters. An object of which nothing is left keeps its place
 * with a box without height: it still owns its source run, so a reader that
 * marks that run finds it, with nothing to mark. Null for a tab without a
 * box. */
function inlineContentBounds(
  placement: Extract<ParagraphLayout['lines'][number]['placements'][number], { kind: 'tab' | 'resource' | 'drawing' }>,
  pointToPage: Matrix2DData,
  clips: readonly ElementClipGeometry[],
): LayoutRect | null {
  if (placement.kind === 'tab') return placement.bounds ?? null;
  const hidden = { ...placement.bounds, heightPt: 0 };
  if (placement.kind === 'resource') {
    if (!placement.lineClip && !placement.rotationDeg) return placement.bounds;
    let pieces: PointPt[][] = [[...resourcePlacementOutline(placement)]];
    if (placement.rotationDeg) {
      for (const clip of clips) {
        pieces = pieces.flatMap((piece) => {
          // The piece in the clip's own space, cut by each of its rectangles.
          const local = piece.map((point) => inverseMapAffinePoint(clip.pointToPage, mapAffinePoint(pointToPage, point)));
          if (local.some((point) => point === null)) return [piece];
          return (clip.rects ?? [clip.bounds]).map((rect) => cutToRect(local as PointPt[], rect)
            .flatMap((point) => inverseMapAffinePoint(pointToPage, mapAffinePoint(clip.pointToPage, point)) ?? []));
        }).filter((piece) => piece.length >= 3);
      }
    }
    let left = Infinity;
    let top = Infinity;
    let right = -Infinity;
    let bottom = -Infinity;
    for (const point of pieces.flat()) {
      left = Math.min(left, point.xPt);
      top = Math.min(top, point.yPt);
      right = Math.max(right, point.xPt);
      bottom = Math.max(bottom, point.yPt);
    }
    return right - left > 1e-9 && bottom - top > 1e-9
      ? { xPt: left, yPt: top, widthPt: right - left, heightPt: bottom - top } : hidden;
  }
  const clip = placement.lineClip;
  if (!clip) return placement.bounds;
  const top = Math.max(placement.bounds.yPt, placement.bounds.yPt + clip.topPt);
  const bottom = Math.min(placement.bounds.yPt + placement.bounds.heightPt, placement.bounds.yPt + clip.bottomPt);
  return bottom > top ? { ...placement.bounds, yPt: top, heightPt: bottom - top } : hidden;
}

/** A convex polygon cut to an axis-aligned rectangle of its own space
 * (Sutherland-Hodgman); empty when nothing is inside. */
function cutToRect(polygon: readonly PointPt[], rect: LayoutRect): PointPt[] {
  const edges: readonly (readonly ['xPt' | 'yPt', number, 1 | -1])[] = [
    ['xPt', rect.xPt, 1], ['xPt', rect.xPt + rect.widthPt, -1], ['yPt', rect.yPt, 1], ['yPt', rect.yPt + rect.heightPt, -1],
  ];
  let points = [...polygon];
  for (const [axis, edge, side] of edges) {
    const kept: PointPt[] = [];
    for (let index = 0; index < points.length; index += 1) {
      const current = points[index]!;
      const previous = points[(index + points.length - 1) % points.length]!;
      const currentIn = side * (current[axis] - edge) >= 0;
      const previousIn = side * (previous[axis] - edge) >= 0;
      if (currentIn !== previousIn) {
        const t = (edge - previous[axis]) / (current[axis] - previous[axis]);
        kept.push({ xPt: previous.xPt + t * (current.xPt - previous.xPt), yPt: previous.yPt + t * (current.yPt - previous.yPt) });
      }
      if (currentIn) kept.push(current);
    }
    points = kept;
    if (points.length === 0) break;
  }
  return points;
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
    const inlineDrawingRunIndexes = new Map(paragraph.drawings.flatMap((drawing) => {
      const runIndex = drawing.source.path.at(-1);
      return !drawing.anchorLayer && runIndex !== undefined && Number.isSafeInteger(runIndex) && runIndex >= 0
        ? [[drawing.id, runIndex] as const] : [];
    }));
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
        // A tab or an inline object occupies its run's place on the line:
        // readers that mark a range of source runs (comments) cover it, and
        // a paragraph that holds only such content still has geometry.
        if (retained.kind === 'tab' || retained.kind === 'resource' || retained.kind === 'drawing') {
          const runIndex = retained.kind === 'drawing'
            ? inlineDrawingRunIndexes.get(retained.drawingId) : retained.sourceRunIndex;
          const contentClips = retained.kind !== 'tab' && retained.escapesCellClip ? paragraphProjection.clips : clippedProjection.clips;
          const bounds = inlineContentBounds(retained, linePointToPage, contentClips);
          if (runIndex !== undefined && bounds && (retained.kind !== 'resource' || retained.resourceKind !== 'picture-bullet')) {
            if (context.collectTextRuns) {
              context.runs.push(Object.freeze({
                inlineContent: retained.kind === 'tab' ? 'tab' as const : 'object' as const,
                bounds,
                sourceRunIndex: runIndex,
                pointToPage: linePointToPage,
                clips: contentClips,
                source: paragraph.source,
                ...(paragraph.paragraphId !== undefined ? { paragraphId: paragraph.paragraphId } : {}),
              }));
            }
            if (context.collectTextRunSources) {
              const key = sourceKey(paragraph.source);
              const indices = context.sourceRuns.get(key) ?? new Set<number>();
              if (!context.sourceRuns.has(key)) context.sourceRuns.set(key, indices);
              indices.add(runIndex);
            }
          }
        }
        if (retained.kind === 'text') for (const placement of sourceOwnedTextPlacements(retained)) {
          if (context.collectTextRuns) {
            context.runs.push(Object.freeze({
              placement,
              pointToPage: linePointToPage,
              clips: clippedProjection.clips,
              clusterBoxes: placement.clusterBoxes,
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

/** The text runs of a page; with `'line-content'`, also the tabs and inline
 * objects between them, in the order their lines hold them. */
export function textRunGeometryForPage(layout: DocumentLayout, pageIndex: number): readonly TextRunGeometry[];
export function textRunGeometryForPage(
  layout: DocumentLayout,
  pageIndex: number,
  content: 'line-content',
): readonly (TextRunGeometry | InlineContentGeometry)[];
export function textRunGeometryForPage(
  layout: DocumentLayout,
  pageIndex: number,
  content?: 'line-content',
): readonly (TextRunGeometry | InlineContentGeometry)[] {
  const runs = pageGeometryIndex(layout, pageIndex, {
    collectTextRuns: true,
    collectTextRunSources: false,
    collectDrawings: false,
  }).runs;
  return Object.freeze(content === 'line-content' ? runs : runs.filter((geometry) => 'placement' in geometry));
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
  /** The pages, in order, that hold laid-out runs of each paragraph source. */
  readonly sourcePages: ReadonlyMap<string, readonly number[]>;
  /** The runs that a paragraph crossing pages holds on more than one of
   * them, by paragraph source then run index. `continued` lists the pages
   * whose fragment of the run follows an earlier one (the run starts on an
   * earlier page), `continues` those whose fragment is followed by a later
   * one (the run ends on a later page). A paragraph repeated whole on
   * several pages (a header) splits nothing. */
  readonly splitRuns: SplitRunIndex;
}

export type SplitRunIndex = ReadonlyMap<string, ReadonlyMap<number, Readonly<{
  continued: readonly number[];
  continues: readonly number[];
}>>>;

/** Build every index needed by comment/revision projection in one canonical
 * geometry traversal. Progressive publications call this only when review data
 * exists; combining the indexes avoids walking the growing prefix twice. */
export function reviewProjectionIndexForDocument(
  layout: DocumentLayout,
): ReviewProjectionIndex {
  const renderedRunIndex = new Map<string, Set<number>>();
  const completedSourceKeys = new Set<string>();
  const sourcePages = new Map<string, number[]>();
  const splitRuns = new Map<string, Map<number, { continued: number[]; continues: number[] }>>();
  // Paragraphs whose last fragment so far continues: run index -> the last
  // page that holds a fragment of it.
  const continuing = new Map<string, Map<number, number>>();
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
      const pages = sourcePages.get(key);
      if (pages) pages.push(pageIndex);
      else sourcePages.set(key, [pageIndex]);
      const earlier = continuing.get(key);
      if (earlier) {
        for (const index of pageIndices) {
          const previousPage = earlier.get(index);
          if (previousPage === undefined) continue;
          const bySource = splitRuns.get(key) ?? new Map<number, { continued: number[]; continues: number[] }>();
          if (!splitRuns.has(key)) splitRuns.set(key, bySource);
          const split = bySource.get(index) ?? { continued: [], continues: [] };
          if (!bySource.has(index)) bySource.set(index, split);
          split.continued.push(pageIndex);
          split.continues.push(previousPage);
        }
      }
      if (!page.completedParagraphSources.has(key)) {
        const held = earlier ?? new Map<number, number>();
        if (!earlier) continuing.set(key, held);
        for (const index of pageIndices) held.set(index, pageIndex);
      }
    }
    for (const key of page.completedParagraphSources) {
      completedSourceKeys.add(key);
      continuing.delete(key);
    }
  }
  return Object.freeze({ renderedRunIndex, completedSourceKeys, sourcePages, splitRuns });
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

import {
  composeAffine,
  translationAffine,
} from './affine.js';
import { sourceKey } from './source-key.js';
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
  ResourcePlacement,
  SourceRef,
  TableLayout,
  TextBoxLayout,
  TextPlacement,
} from './types.js';

export interface TextRunGeometry {
  readonly placement: TextPlacement;
  readonly pointToPage: Matrix2DData;
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
  drawingSourceOrder: number;
}

interface NodeProjection {
  readonly pointToPage: Matrix2DData;
  readonly layoutTranslationPt: PointPt;
  readonly rootNodeId: string;
  readonly paintOrderIndex: number;
  readonly clips: readonly ElementClipGeometry[];
}

const IDENTITY_AFFINE = Object.freeze({
  a: 1, b: 0, c: 0, d: 1, e: 0, f: 0,
}) satisfies Matrix2DData;

const EMPTY_CLIPS = Object.freeze([]) satisfies readonly ElementClipGeometry[];

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
    if (domain.kind !== 'footnote' && domain.kind !== 'endnote') continue;
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
  const transformedProjection: NodeProjection = {
    ...projection,
    pointToPage: composeAffine(projection.pointToPage, textBox.transform),
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

function visitParagraph(
  paragraph: ParagraphLayout,
  projection: NodeProjection,
  context: ProjectionContext,
): void {
  const paragraphProjection = withClip(projection, paragraph.clipBounds);
  if (
    context.collectCompletedParagraphSources
    && paragraph.continuation?.continuesOnNext !== true
  ) {
    context.completedParagraphSources.add(sourceKey(paragraph.source));
  }
  if (context.collectTextRuns || context.collectTextRunSources) {
    for (const line of paragraph.lines) {
      for (const placement of line.placements) {
        if (placement.kind === 'text') {
          if (context.collectTextRuns) {
            context.runs.push(Object.freeze({
              placement,
              pointToPage: projection.pointToPage,
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

  const textBoxesById = new Map(
    paragraph.textBoxes.map((textBox) => [textBox.id, textBox]),
  );
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
      visitDrawing(textBoxesById, entry.drawing, paragraphProjection, context);
      continue;
    }
    if (!context.collectDrawings) continue;
    context.inlineResources.push(Object.freeze({
      placement: entry.placement,
      source: Object.freeze({
        ...paragraph.source,
        path: Object.freeze([...paragraph.source.path, entry.runIndex]),
      }),
      pointToPage: paragraphProjection.pointToPage,
      clips: paragraphProjection.clips,
      paintOrderIndex: paragraphProjection.paintOrderIndex,
      sourceOrder: context.drawingSourceOrder++,
    }));
  }
  for (const textBox of paragraph.textBoxes) {
    if (!ownedTextBoxIds.has(textBox.id)) {
      visitTextBox(textBox, paragraphProjection, context);
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
      const ownsContinuationPaint = 'visualMergeOwnership' in cell
        && cell.visualMergeOwnership === 'continuation';
      if (cell.verticalMerge === 'continue' && !ownsContinuationPaint) continue;
      const cellProjection = withClip(tableProjection, cell.clipBounds);
      for (const block of cell.blocks) {
        const child = block.layout;
        visitNode(child, placedChildProjection(child, {
          xPt: cell.contentBounds.xPt
            + (child.kind === 'table' ? child.flowBounds.xPt : 0),
          yPt: cell.flowBounds.yPt + block.offsetPt
            + (child.kind === 'table' ? child.flowBounds.yPt : 0),
        }, cellProjection), context);
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
    case 'note':
      for (const block of node.story.blocks) {
        visitNode(block, withClip(projection, node.story.clipBounds), context);
      }
      return;
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
    drawingSourceOrder: 0,
  };
  for (const nodeId of page.readingOrder) {
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

/** Canonical paragraph sources that have an occurrence in the supplied layout.
 * Unlike the rendered-run index this includes empty and final-state-hidden
 * paragraphs, so a provisional review projection can distinguish content that
 * is already paginated from content that belongs to a future prefix. */
export function completedParagraphSourceKeysForDocument(
  layout: DocumentLayout,
): ReadonlySet<string> {
  const result = new Set<string>();
  for (let pageIndex = 0; pageIndex < layout.pages.length; pageIndex += 1) {
    const pageSources = pageGeometryIndex(layout, pageIndex, {
      collectTextRuns: false,
      collectTextRunSources: false,
      collectCompletedParagraphSources: true,
      collectDrawings: false,
    }).completedParagraphSources;
    for (const key of pageSources) result.add(key);
  }
  return result;
}

/**
 * Index retained drawings in page paint order. The projection is shared with
 * text-box painting, including vertical-page and page-owned anchor transforms.
 */
export function drawingGeometryForPage(
  layout: DocumentLayout,
  pageIndex: number,
): readonly DrawingGeometry[] {
  const drawings = pageGeometryIndex(layout, pageIndex, {
    collectTextRuns: false,
    collectTextRunSources: false,
    collectDrawings: true,
  }).drawings;
  drawings.sort((left, right) => left.paintOrderIndex - right.paintOrderIndex ||
    left.sourceOrder - right.sourceOrder);
  return Object.freeze(drawings);
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

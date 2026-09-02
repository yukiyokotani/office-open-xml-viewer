import type {
  DrawingLayout,
  Matrix2DData,
  PageLayerId,
  PageLayerRoot,
  PageLayers,
  PagePaintDrawingEntry,
  PagePaintEntry,
  PagePaintFrame,
  PagePaintNodeEntry,
  PaintNode,
  ParagraphLayout,
  PointPt,
  ResolvedFloatingTablePlacementLayout,
  TableLayout,
  TextBoxLayout,
} from './types.js';

export const PAGE_LAYER_IDS = [
  'background',
  'behindText',
  'header',
  'body',
  'notes',
  'front',
  'footer',
] as const satisfies readonly PageLayerId[];

type MissingPageLayer = Exclude<PageLayerId, typeof PAGE_LAYER_IDS[number]>;
const pageLayersAreExhaustive: MissingPageLayer extends never ? true : never = true;
void pageLayersAreExhaustive;

export type PageLayerNode = Omit<PageLayerRoot, 'coordinateSpace'> & Readonly<{
  coordinateSpace?: PageLayerRoot['coordinateSpace'];
}>;

interface DrawingCandidate {
  readonly drawing: DrawingLayout;
  readonly owner?: ParagraphLayout;
  readonly textBoxes: readonly TextBoxLayout[];
  readonly frames: readonly PagePaintFrame[];
  readonly layoutTranslationPt: PointPt;
  readonly encounterOrder: number;
  readonly root: PageLayerRoot;
}

const identity = Object.freeze({
  a: 1, b: 0, c: 0, d: 1, e: 0, f: 0,
}) satisfies Matrix2DData;

function clipFrame(
  rect: Readonly<{ xPt: number; yPt: number; widthPt: number; heightPt: number }>,
): PagePaintFrame {
  return Object.freeze({
    kind: 'clip',
    clip: rect,
  });
}

function transformFrame(xPt: number, yPt: number): PagePaintFrame {
  return Object.freeze({
    kind: 'transform',
    transform: Object.freeze({ ...identity, e: xPt, f: yPt }),
  });
}

function textBoxesForDrawing(
  owner: ParagraphLayout,
  drawing: DrawingLayout,
): readonly TextBoxLayout[] {
  if (!drawing.textBoxIds?.length) return Object.freeze([]);
  const byId = new Map(owner.textBoxes.map((textBox) => [textBox.id, textBox]));
  return Object.freeze(drawing.textBoxIds.flatMap((id) => {
    const textBox = byId.get(id);
    return textBox ? [textBox] : [];
  }));
}

function visitAnchoredDrawings(
  node: PaintNode,
  root: PageLayerRoot,
  frames: readonly PagePaintFrame[],
  layoutTranslationPt: PointPt,
  candidates: DrawingCandidate[],
): void {
  if (node.kind === 'drawing') {
    if (!node.anchorLayer) return;
    candidates.push(Object.freeze({
      drawing: node,
      textBoxes: Object.freeze([]),
      frames: Object.freeze([...frames]),
      layoutTranslationPt: Object.freeze({ ...layoutTranslationPt }),
      encounterOrder: candidates.length,
      root,
    }));
    return;
  }
  if (node.kind === 'textbox') {
    // A text-box story is an atomic descendant of its owner drawing. Its own
    // anchors remain in that local stacking context and are never page-queued.
    return;
  }
  if (node.kind === 'note') {
    const noteFrames = node.story.clipBounds
      ? Object.freeze([...frames, clipFrame(node.story.clipBounds)])
      : frames;
    for (const block of node.story.blocks) {
      visitAnchoredDrawings(
        block,
        root,
        noteFrames,
        layoutTranslationPt,
        candidates,
      );
    }
    return;
  }
  if (node.kind === 'paragraph') {
    const paragraphFrames = node.clipBounds
      ? Object.freeze([...frames, clipFrame(node.clipBounds)])
      : frames;
    for (const drawing of node.drawings) {
      if (!drawing.anchorLayer) continue;
      candidates.push(Object.freeze({
        drawing,
        owner: node,
        textBoxes: textBoxesForDrawing(node, drawing),
        frames: Object.freeze([...paragraphFrames]),
        layoutTranslationPt: Object.freeze({ ...layoutTranslationPt }),
        encounterOrder: candidates.length,
        root,
      }));
    }
    return;
  }
  visitTableAnchoredDrawings(
    node,
    root,
    frames,
    layoutTranslationPt,
    candidates,
  );
}

function visitPlacedTableChild(
  node: PaintNode,
  placement: Readonly<{ xPt: number; yPt: number }>,
  root: PageLayerRoot,
  frames: readonly PagePaintFrame[],
  layoutTranslationPt: PointPt,
  candidates: DrawingCandidate[],
): void {
  const dxPt = placement.xPt - node.flowBounds.xPt;
  const dyPt = placement.yPt - node.flowBounds.yPt;
  visitAnchoredDrawings(
    node,
    root,
    Object.freeze([...frames, transformFrame(dxPt, dyPt)]),
    Object.freeze({
      xPt: layoutTranslationPt.xPt + dxPt,
      yPt: layoutTranslationPt.yPt + dyPt,
    }),
    candidates,
  );
}

function visitResolvedFloatingTable(
  placement: ResolvedFloatingTablePlacementLayout,
  root: PageLayerRoot,
  frames: readonly PagePaintFrame[],
  layoutTranslationPt: PointPt,
  candidates: DrawingCandidate[],
): void {
  visitPlacedTableChild(
    placement.child,
    {
      xPt: placement.xPt - layoutTranslationPt.xPt,
      yPt: placement.yPt - layoutTranslationPt.yPt,
    },
    root,
    frames,
    layoutTranslationPt,
    candidates,
  );
}

function visitTableAnchoredDrawings(
  table: TableLayout,
  root: PageLayerRoot,
  frames: readonly PagePaintFrame[],
  layoutTranslationPt: PointPt,
  candidates: DrawingCandidate[],
): void {
  const tableFrames = table.clipBounds
    ? Object.freeze([...frames, clipFrame(table.clipBounds)])
    : frames;
  for (const row of table.rows) {
    for (const cell of row.cells) {
      const ownsContinuationPaint = 'visualMergeOwnership' in cell
        && cell.visualMergeOwnership === 'continuation';
      if (cell.verticalMerge === 'continue' && !ownsContinuationPaint) continue;
      const cellFrames = cell.clipBounds
        ? Object.freeze([...tableFrames, clipFrame(cell.clipBounds)])
        : tableFrames;
      for (const block of cell.blocks) {
        visitPlacedTableChild(
          block.layout,
          {
            xPt: cell.contentBounds.xPt
              + (block.layout.kind === 'table' ? block.layout.flowBounds.xPt : 0),
            yPt: cell.flowBounds.yPt + block.offsetPt
              + (block.layout.kind === 'table' ? block.layout.flowBounds.yPt : 0),
          },
          root,
          cellFrames,
          layoutTranslationPt,
          candidates,
        );
      }
    }
  }
  for (const placement of table.resolvedFloatingTables ?? []) {
    visitResolvedFloatingTable(
      placement,
      root,
      tableFrames,
      layoutTranslationPt,
      candidates,
    );
  }
}

function drawingPaintEntry(candidate: DrawingCandidate): PagePaintDrawingEntry {
  const anchor = candidate.drawing.anchorLayer!;
  return Object.freeze({
    kind: 'drawing',
    layer: anchor.behindDoc ? 'behindText' : 'front',
    sourceLayer: candidate.root.layer,
    rootNodeId: candidate.root.node.id,
    coordinateSpace: candidate.root.coordinateSpace,
    flowDomainId: candidate.root.node.flowDomainId,
    node: candidate.drawing,
    ...(candidate.owner ? { ownerNodeId: candidate.owner.id } : {}),
    textBoxes: candidate.textBoxes,
    frames: candidate.frames,
    layoutTranslationPt: candidate.layoutTranslationPt,
  });
}

function nodePaintEntry(root: PageLayerRoot, omitAnchoredDrawings: boolean): PagePaintNodeEntry {
  return Object.freeze({
    kind: 'node',
    layer: root.layer,
    sourceLayer: root.layer,
    rootNodeId: root.node.id,
    coordinateSpace: root.coordinateSpace,
    flowDomainId: root.node.flowDomainId,
    node: root.node,
    omitAnchoredDrawings,
  });
}

function compareCandidates(left: DrawingCandidate, right: DrawingCandidate): number {
  return left.drawing.anchorLayer!.relativeHeight - right.drawing.anchorLayer!.relativeHeight
    || left.drawing.anchorLayer!.sourceOrder - right.drawing.anchorLayer!.sourceOrder
    || left.encounterOrder - right.encounterOrder;
}

function materializeStackingContext(roots: readonly PageLayerRoot[]): readonly PagePaintEntry[] {
  const candidates: DrawingCandidate[] = [];
  for (const root of roots) {
    visitAnchoredDrawings(
      root.node,
      root,
      Object.freeze([]),
      Object.freeze({ xPt: 0, yPt: 0 }),
      candidates,
    );
  }
  const behind = candidates
    .filter(({ drawing }) => drawing.anchorLayer!.behindDoc)
    .sort(compareCandidates)
    .map(drawingPaintEntry);
  const front = candidates
    .filter(({ drawing }) => !drawing.anchorLayer!.behindDoc)
    .sort(compareCandidates)
    .map(drawingPaintEntry);
  const candidateRoots = new Set(candidates.map(({ root }) => root.node));
  const nodes = roots.flatMap((root) => (
    root.node.kind === 'drawing' && root.node.anchorLayer
      ? []
      : [nodePaintEntry(root, candidateRoots.has(root.node))]
  ));
  return Object.freeze([...behind, ...nodes, ...front]);
}

function retainedNodeRequiresElementBackedVerticalGlyphPaint(
  node: PaintNode,
  seen: Set<PaintNode>,
): boolean {
  if (seen.has(node)) return false;
  seen.add(node);
  if (node.kind === 'drawing') return false;
  if (node.kind === 'paragraph') {
    const hasVerticalGlyph = node.lines.some((line) => line.placements.some((placement) => (
      placement.kind === 'text'
      && placement.paintOps?.some((operation) => operation.verticalFeature === true) === true
    )));
    return hasVerticalGlyph || node.textBoxes.some((textBox) => (
      retainedNodeRequiresElementBackedVerticalGlyphPaint(textBox, seen)
    ));
  }
  if (node.kind === 'textbox' || node.kind === 'note') {
    return node.story.blocks.some((block) => (
      retainedNodeRequiresElementBackedVerticalGlyphPaint(block, seen)
    ));
  }
  return node.rows.some((row) => row.cells.some((cell) => cell.blocks.some((block) => (
    retainedNodeRequiresElementBackedVerticalGlyphPaint(block.layout, seen)
  )))) || (node.resolvedFloatingTables ?? []).some((placement) => (
    retainedNodeRequiresElementBackedVerticalGlyphPaint(placement.child, seen)
  ));
}

function collectRetainedNodeResourceKeys(
  node: PaintNode,
  keys: Set<string>,
  seen: Set<PaintNode>,
): void {
  if (seen.has(node)) return;
  seen.add(node);
  if (node.kind === 'drawing') {
    for (const command of node.commands) {
      if (command.kind === 'resource' || command.kind === 'drawingml-image-fill') {
        keys.add(command.resourceKey);
      }
    }
    return;
  }
  if (node.kind === 'paragraph') {
    for (const resource of node.resources) keys.add(resource.resourceKey);
    for (const drawing of node.drawings) collectRetainedNodeResourceKeys(drawing, keys, seen);
    for (const textBox of node.textBoxes) collectRetainedNodeResourceKeys(textBox, keys, seen);
    return;
  }
  if (node.kind === 'textbox' || node.kind === 'note') {
    for (const block of node.story.blocks) collectRetainedNodeResourceKeys(block, keys, seen);
    return;
  }
  for (const row of node.rows) {
    for (const cell of row.cells) {
      for (const block of cell.blocks) {
        collectRetainedNodeResourceKeys(block.layout, keys, seen);
      }
    }
  }
  for (const placement of node.resolvedFloatingTables ?? []) {
    collectRetainedNodeResourceKeys(placement.child, keys, seen);
  }
}

/** Build the final immutable page paint plan. ECMA-376 §20.4.2.3 ordering is
 * applied only within an equivalent story stacking context; cross-story order
 * remains the explicit root order unless a separately registered compatibility
 * rule authorizes a change. */
export function buildPageLayers(entries: readonly PageLayerNode[]): PageLayers {
  const roots = Object.freeze(entries.map(({ layer, node, coordinateSpace }) => Object.freeze({
    layer,
    node,
    coordinateSpace: coordinateSpace ?? 'section-logical',
  })));
  const nodes = new Map(PAGE_LAYER_IDS.map((layer) => [layer, [] as PaintNode[]]));
  for (const entry of roots) nodes.get(entry.layer)!.push(entry.node);
  const paintOrder: PagePaintEntry[] = [];
  for (let index = 0; index < roots.length;) {
    const sourceLayer = roots[index]!.layer;
    let end = index + 1;
    while (roots[end]?.layer === sourceLayer) end += 1;
    const group = roots.slice(index, end);
    if (
      sourceLayer === 'header'
      || sourceLayer === 'body'
      || sourceLayer === 'notes'
      || sourceLayer === 'footer'
    ) {
      paintOrder.push(...materializeStackingContext(group));
    } else {
      paintOrder.push(...group.map((root) => nodePaintEntry(root, false)));
    }
    index = end;
  }
  const seen = new Set<PaintNode>();
  const resourceKeys = new Set<string>();
  const resourceSeen = new Set<PaintNode>();
  for (const { node } of roots) {
    collectRetainedNodeResourceKeys(node, resourceKeys, resourceSeen);
  }
  return Object.freeze({
    roots,
    paintOrder: Object.freeze(paintOrder),
    capabilities: Object.freeze({
      requiresElementBackedVerticalGlyphPaint: roots.some(({ node }) => (
        retainedNodeRequiresElementBackedVerticalGlyphPaint(node, seen)
      )),
      resourceKeys: Object.freeze([...resourceKeys]),
    }),
    background: Object.freeze(nodes.get('background')!),
    behindText: Object.freeze(nodes.get('behindText')!),
    header: Object.freeze(nodes.get('header')!),
    body: Object.freeze(nodes.get('body')!),
    notes: Object.freeze(nodes.get('notes')!),
    front: Object.freeze(nodes.get('front')!),
    footer: Object.freeze(nodes.get('footer')!),
  });
}

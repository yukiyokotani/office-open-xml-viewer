import type {
  LayoutPage,
  LogicalBlockFootprint,
  PageOccurrenceCoordinateSpace,
  PageLayerId,
  PaintNode,
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

export type PageLayerNode = Readonly<{
  layer: PageLayerId;
  node: PaintNode;
  coordinateSpace: PageOccurrenceCoordinateSpace;
  logicalBlock?: LogicalBlockFootprint;
}>;

export class PageGraphError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'PageGraphError';
  }
}

export function pageLayerNodes(page: LayoutPage): readonly PageLayerNode[] {
  return orderedPagePaintEntries(page);
}

export function orderedPagePaintEntries(page: LayoutPage): readonly PageLayerNode[] {
  const nodes = new Map<string, Readonly<{ layer: PageLayerId; node: PaintNode }>>();
  for (const layer of PAGE_LAYER_IDS) for (const node of page.layers[layer]) {
    const entry = { layer, node };
    if (nodes.has(entry.node.id)) throw new PageGraphError(`Duplicate paint node ${entry.node.id}`);
    nodes.set(entry.node.id, entry);
  }

  const painted = new Set<string>();
  const ordered: PageLayerNode[] = [];
  for (const entry of page.layers.paintOrder) {
    const target = nodes.get(entry.nodeId);
    if (!target) throw new PageGraphError(`Missing paint node ${entry.nodeId}`);
    if (target.layer !== entry.layer) {
      throw new PageGraphError(`Paint node ${entry.nodeId} belongs to ${target.layer}, not ${entry.layer}`);
    }
    if (painted.has(entry.nodeId)) throw new PageGraphError(`Duplicate paint reference ${entry.nodeId}`);
    painted.add(entry.nodeId);
    ordered.push({
      ...target,
      coordinateSpace: entry.coordinateSpace,
      ...(entry.logicalBlock ? { logicalBlock: entry.logicalBlock } : {}),
    });
  }
  if (painted.size !== nodes.size) {
    const missing = [...nodes.keys()].find((id) => !painted.has(id));
    throw new PageGraphError(`Missing paint-order reference for ${missing ?? '<unknown>'}`);
  }
  return ordered;
}

export function orderedPagePaintNodes(page: LayoutPage): readonly PaintNode[] {
  return orderedPagePaintEntries(page).map((entry) => entry.node);
}

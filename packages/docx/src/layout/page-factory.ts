import type { SectionLayoutContext } from '../layout-context.js';
import { canonicalLogicalToPhysical, mapAffineRect } from './affine.js';
import { PAGE_LAYER_IDS, type PageLayerNode } from './page-graph.js';
import type {
  DeepReadonly,
  FlowDomain,
  LayoutPage,
  PageBookmarkStart,
  PageLayers,
  PageNumberMetadata,
  PagePaintNode,
  PageSectionRegion,
  PaintNode,
  ParagraphLayout,
  TableLayout,
  TextBoxLayout,
  WritingMode,
} from './types.js';

export interface PhysicalPageInput {
  readonly widthPt: number;
  readonly heightPt: number;
  /** Effective main-story edges after §17.6.11 header/footer interaction. */
  readonly contentTopPt: number;
  readonly contentBottomPt: number;
}

export interface LogicalColumnInput {
  readonly inlineStartPt: number;
  readonly inlineExtentPt: number;
}

export interface PageSectionRegionInput {
  readonly id: string;
  readonly sectionOccurrenceId: string;
  readonly section: DeepReadonly<SectionLayoutContext>;
  readonly writingMode: WritingMode;
  readonly blockStartPt: number;
  readonly blockEndPt: number;
  readonly columns: readonly LogicalColumnInput[];
}

export interface LayoutPageAccumulatorInput {
  readonly pageIndex: number;
  readonly physicalPage: PhysicalPageInput;
  readonly sectionOccurrenceId: string;
  readonly section: DeepReadonly<SectionLayoutContext>;
}

export interface LayoutPageAccumulator extends LayoutPageAccumulatorInput {
  readonly sectionRegions: readonly PageSectionRegionInput[];
  readonly paint: readonly PageLayerNode[];
  readonly readingOrder: readonly PaintNode[];
}

export interface LayoutPageFactoryInput extends LayoutPageAccumulatorInput {
  readonly sectionRegions: readonly PageSectionRegionInput[];
  readonly paint: readonly PageLayerNode[];
  readonly readingOrder: readonly PaintNode[];
  readonly pageNumber: PageNumberMetadata;
}

export interface ParityBlankLayoutPageInput {
  readonly pageIndex: number;
  readonly physicalPage: PhysicalPageInput;
  readonly sectionOccurrenceId: string;
  readonly section: DeepReadonly<SectionLayoutContext>;
  readonly pageNumber: PageNumberMetadata;
}

export function bodyFlowDomainId(
  pageIndex: number,
  regionId: string,
  columnIndex: number,
): string {
  return `page:${pageIndex}:region:${encodeURIComponent(regionId)}:column:${columnIndex}`;
}

function pageGeometry(page: PhysicalPageInput): LayoutPage['geometry'] {
  requireEffectivePageEdges(page);
  return {
    xPt: 0,
    yPt: 0,
    widthPt: page.widthPt,
    heightPt: page.heightPt,
    // §17.6.11 makes positive top/bottom edges depend on header/footer extent,
    // while negative values ignore that extent. The page owner resolves those
    // facts; this factory only retains the resulting effective coordinates.
    contentTopPt: page.contentTopPt,
    contentBottomPt: page.contentBottomPt,
  };
}

function requireEffectivePageEdges(page: PhysicalPageInput): void {
  if (
    !Number.isFinite(page.heightPt)
    || !Number.isFinite(page.contentTopPt)
    || !Number.isFinite(page.contentBottomPt)
    || page.contentTopPt < 0
    || page.contentTopPt > page.contentBottomPt
    || page.contentBottomPt > page.heightPt
  ) {
    throw new RangeError(
      'Effective page edges must satisfy 0 <= contentTopPt <= contentBottomPt <= heightPt',
    );
  }
  // Equal edges are valid and represent an empty main-story interval.
}

function requirePageIndex(pageIndex: number): void {
  if (!Number.isInteger(pageIndex) || pageIndex < 0) {
    throw new RangeError('Layout page index must be a non-negative integer');
  }
}

function buildRegions(
  pageIndex: number,
  physicalPage: PhysicalPageInput,
  inputs: readonly PageSectionRegionInput[],
): Readonly<{
  regions: readonly PageSectionRegion[];
  domains: readonly FlowDomain[];
  sectionByDomain: ReadonlyMap<string, string>;
}> {
  const regions: PageSectionRegion[] = [];
  const domains: FlowDomain[] = [];
  const sectionByDomain = new Map<string, string>();

  for (const input of inputs) {
    const logicalToPhysical = canonicalLogicalToPhysical(
      input.writingMode,
      physicalPage.widthPt,
    );
    const flowDomainIds = input.columns.map((column, columnIndex) => {
      const id = bodyFlowDomainId(pageIndex, input.id, columnIndex);
      domains.push({
        id,
        kind: 'body',
        bounds: mapAffineRect(logicalToPhysical, {
          xPt: column.inlineStartPt,
          yPt: input.blockStartPt,
          widthPt: column.inlineExtentPt,
          heightPt: input.blockEndPt - input.blockStartPt,
        }),
      });
      sectionByDomain.set(id, input.sectionOccurrenceId);
      return id;
    });
    regions.push({
      id: input.id,
      sectionOccurrenceId: input.sectionOccurrenceId,
      coordinateSpace: { writingMode: input.writingMode, logicalToPhysical },
      blockStartPt: input.blockStartPt,
      blockEndPt: input.blockEndPt,
      flowDomainIds,
      section: input.section,
    });
  }

  return { regions, domains, sectionByDomain };
}

function buildLayers(entries: readonly PageLayerNode[]): PageLayers {
  const nodes = new Map(PAGE_LAYER_IDS.map((layer) => [layer, [] as PagePaintNode[]]));
  for (const entry of entries) nodes.get(entry.layer)!.push(entry.node);
  return {
    paintOrder: entries.map(({ layer, node, coordinateSpace, logicalBlock }) => ({
      layer,
      nodeId: node.id,
      coordinateSpace,
      ...(logicalBlock ? { logicalBlock } : {}),
    })),
    background: nodes.get('background')!,
    behindText: nodes.get('behindText')!,
    header: nodes.get('header')!,
    body: nodes.get('body')!,
    notes: nodes.get('notes')!,
    front: nodes.get('front')!,
    footer: nodes.get('footer')!,
  };
}

function visitBookmarkParagraphs(
  node: PaintNode,
  visit: (paragraph: ParagraphLayout) => void,
): void {
  if (node.kind === 'paragraph') {
    visit(node);
    node.drawings.forEach((drawing) => visitBookmarkParagraphs(drawing, visit));
    node.textBoxes.forEach((textBox) => visitBookmarkParagraphs(textBox, visit));
    return;
  }
  if (node.kind === 'table') {
    visitTableBookmarks(node, visit);
    return;
  }
  if (node.kind === 'textbox') {
    visitTextBoxBookmarks(node, visit);
  }
}

function visitTableBookmarks(
  table: TableLayout,
  visit: (paragraph: ParagraphLayout) => void,
): void {
  for (const row of table.rows) {
    for (const cell of row.cells) {
      for (const block of cell.blocks) visitBookmarkParagraphs(block.layout, visit);
    }
  }
}

function visitTextBoxBookmarks(
  textBox: TextBoxLayout,
  visit: (paragraph: ParagraphLayout) => void,
): void {
  textBox.paragraphs.forEach((paragraph) => visitBookmarkParagraphs(paragraph, visit));
}

function bookmarkStarts(
  paint: readonly PageLayerNode[],
  defaultSectionOccurrenceId: string,
  sectionByDomain: ReadonlyMap<string, string>,
): readonly PageBookmarkStart[] {
  const starts: PageBookmarkStart[] = [];
  const seen = new Set<string>();
  for (const { node } of paint) {
    const sectionOccurrenceId = sectionByDomain.get(node.flowDomainId)
      ?? defaultSectionOccurrenceId;
    visitBookmarkParagraphs(node, (paragraph) => {
      for (const line of paragraph.lines) {
        for (const placement of line.placements) {
          if (placement.kind !== 'text' || !placement.bookmark || seen.has(placement.bookmark)) {
            continue;
          }
          seen.add(placement.bookmark);
          starts.push({
            name: placement.bookmark,
            nodeId: paragraph.id,
            sectionOccurrenceId,
          });
        }
      }
    });
  }
  return starts;
}

export function createLayoutPage(input: LayoutPageFactoryInput): LayoutPage {
  requirePageIndex(input.pageIndex);
  const { regions, domains, sectionByDomain } = buildRegions(
    input.pageIndex,
    input.physicalPage,
    input.sectionRegions,
  );
  return {
    pageIndex: input.pageIndex,
    geometry: pageGeometry(input.physicalPage),
    flowDomains: domains,
    section: input.section,
    sectionOccurrenceId: input.sectionOccurrenceId,
    parityBlank: false,
    bookmarkStarts: bookmarkStarts(
      input.paint,
      input.sectionOccurrenceId,
      sectionByDomain,
    ),
    pageNumber: input.pageNumber,
    sectionRegions: regions,
    layers: buildLayers(input.paint),
    readingOrder: input.readingOrder.map((node) => node.id),
  };
}

export function createLayoutPageAccumulator(
  input: LayoutPageAccumulatorInput,
): LayoutPageAccumulator {
  requirePageIndex(input.pageIndex);
  requireEffectivePageEdges(input.physicalPage);
  return Object.freeze({
    ...input,
    sectionRegions: Object.freeze([]),
    paint: Object.freeze([]),
    readingOrder: Object.freeze([]),
  });
}

export function accumulatePageSectionRegion(
  accumulator: LayoutPageAccumulator,
  sectionRegion: PageSectionRegionInput,
): LayoutPageAccumulator {
  return Object.freeze({
    ...accumulator,
    sectionRegions: Object.freeze([...accumulator.sectionRegions, sectionRegion]),
  });
}

export function accumulatePagePaintNode(
  accumulator: LayoutPageAccumulator,
  entry: PageLayerNode,
  inReadingOrder: boolean,
): LayoutPageAccumulator {
  return Object.freeze({
    ...accumulator,
    paint: Object.freeze([...accumulator.paint, entry]),
    readingOrder: inReadingOrder
      ? Object.freeze([...accumulator.readingOrder, entry.node])
      : accumulator.readingOrder,
  });
}

export function finalizeLayoutPage(
  accumulator: LayoutPageAccumulator,
  pageNumber: PageNumberMetadata,
): LayoutPage {
  return createLayoutPage({ ...accumulator, pageNumber });
}

export function createParityBlankLayoutPage(
  input: ParityBlankLayoutPageInput,
): LayoutPage {
  requirePageIndex(input.pageIndex);
  return {
    pageIndex: input.pageIndex,
    geometry: pageGeometry(input.physicalPage),
    flowDomains: [],
    section: input.section,
    sectionOccurrenceId: input.sectionOccurrenceId,
    parityBlank: true,
    bookmarkStarts: [],
    pageNumber: input.pageNumber,
    sectionRegions: [],
    layers: buildLayers([]),
    readingOrder: [],
  };
}

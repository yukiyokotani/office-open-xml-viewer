import type { SectionLayoutContext } from '../layout-context.js';
import {
  createSectionRegionCoordinateSpace,
  logicalPageExtent,
  sectionWritingMode,
  transformRect,
  uprightPhysicalExtent,
  type PhysicalPageExtent,
} from './coordinate-space.js';
import { columnSeparatorSegments } from './column-separators.js';
import { materializePageBorderLayout } from './page-border.js';
import {
  createPageLayers,
  orderedPagePaintNodes,
  type PageLayerNode,
} from './page-graph.js';
import type { BodyOccurrenceDestination } from './occurrence-projection.js';
import type {
  DeepReadonly,
  FlowDomain,
  LayoutPage,
  LayoutRect,
  PageBookmarkStart,
  PageNumberMetadata,
  PageSectionRegion,
  PaintNode,
  ParagraphLayout,
  TableLayout,
  TextBoxLayout,
  WritingMode,
} from './types.js';

export interface PhysicalPageInput extends PhysicalPageExtent {
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
  readonly pageBorders?: DeepReadonly<import('../types.js').PageBorders> | null;
  readonly writingMode: WritingMode;
  readonly blockStartPt: number;
  readonly blockEndPt: number;
  readonly columnFlowDirection?: 'ltr' | 'rtl';
  /** Defaults to every authored section column for ordinary full-width regions. */
  readonly columnIndexes?: readonly number[];
  readonly columns: readonly LogicalColumnInput[];
}

export interface LayoutPageAccumulatorInput {
  readonly pageIndex: number;
  readonly physicalPage: PhysicalPageInput;
  readonly sectionOccurrenceId: string;
  readonly section: DeepReadonly<SectionLayoutContext>;
  readonly pageBorders?: DeepReadonly<import('../types.js').PageBorders> | null;
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
  /** Required whenever page borders are authored; document finalization owns
   * this occurrence-sensitive fact. */
  readonly firstSectionOwnedPage?: boolean;
}

export interface ParityBlankLayoutPageInput {
  readonly pageIndex: number;
  readonly physicalPage: PhysicalPageInput;
  readonly sectionOccurrenceId: string;
  readonly section: DeepReadonly<SectionLayoutContext>;
  readonly pageNumber: PageNumberMetadata;
  readonly pageBorders?: DeepReadonly<import('../types.js').PageBorders> | null;
  /** Required whenever page borders are authored; parity blanks participate
   * in section-owned page identity. */
  readonly firstSectionOwnedPage?: boolean;
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
    !Number.isFinite(page.widthPt)
    || !Number.isFinite(page.heightPt)
    || !Number.isFinite(page.contentTopPt)
    || !Number.isFinite(page.contentBottomPt)
    || page.widthPt <= 0
    || page.heightPt <= 0
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

function requireIdentity(value: string, name: string): void {
  if (value.length === 0) throw new RangeError(`${name} must not be empty`);
}

function pageBorderFirstPageFlag(
  pageBorders: DeepReadonly<import('../types.js').PageBorders> | null | undefined,
  firstSectionOwnedPage: boolean | undefined,
): boolean {
  if (pageBorders && firstSectionOwnedPage === undefined) {
    throw new RangeError(
      'Page-border finalization requires explicit section-owned page identity',
    );
  }
  return firstSectionOwnedPage ?? false;
}

function equalColumns(
  left: SectionLayoutContext['columns'],
  right: SectionLayoutContext['columns'],
): boolean {
  return left.length === right.length && left.every((column, index) => {
    const other = right[index];
    return other !== undefined && column.xPt === other.xPt && column.wPt === other.wPt;
  });
}

function equalLineNumbering(
  left: SectionLayoutContext['lineNumbering'],
  right: SectionLayoutContext['lineNumbering'],
): boolean {
  return left === right || (left !== undefined && right !== undefined
    && left.start === right.start
    && left.countBy === right.countBy
    && left.distance === right.distance
    && left.restart === right.restart);
}

export function sectionLayoutContextsEqual(
  left: DeepReadonly<SectionLayoutContext>,
  right: DeepReadonly<SectionLayoutContext>,
): boolean {
  return left.geometry.pageWidth === right.geometry.pageWidth
    && left.geometry.pageHeight === right.geometry.pageHeight
    && left.geometry.marginTop === right.geometry.marginTop
    && left.geometry.marginRight === right.geometry.marginRight
    && left.geometry.marginBottom === right.geometry.marginBottom
    && left.geometry.marginLeft === right.geometry.marginLeft
    && left.geometry.headerDistance === right.geometry.headerDistance
    && left.geometry.footerDistance === right.geometry.footerDistance
    && equalColumns(left.columns, right.columns)
    && left.columnSeparator === right.columnSeparator
    && left.textDirection === right.textDirection
    && (left.sectionFlow ?? null) === (right.sectionFlow ?? null)
    && (left.sectionBidi === true) === (right.sectionBidi === true)
    && left.grid.kind === right.grid.kind
    && left.grid.linePitchPt === right.grid.linePitchPt
    && left.grid.charSpacePt === right.grid.charSpacePt
    && left.verticalAlignment === right.verticalAlignment
    && equalLineNumbering(left.lineNumbering, right.lineNumbering);
}

function requireRegionSectionAgreement(input: PageSectionRegionInput): void {
  const writingMode = sectionWritingMode(input.section);
  if (writingMode !== input.writingMode) {
    throw new RangeError('Section region writing mode must agree with its section text direction');
  }
  const sectionColumnFlowDirection = input.section.sectionBidi === true ? 'rtl' : 'ltr';
  if (
    input.columnFlowDirection !== undefined
    && input.columnFlowDirection !== sectionColumnFlowDirection
  ) {
    throw new RangeError('Section region column flow direction must agree with sectPr bidi');
  }
  const columnIndexes = input.columnIndexes
    ?? input.section.columns.map((_, index) => index);
  if (input.columns.length !== columnIndexes.length
    || columnIndexes.some((columnIndex, index) => (
      !Number.isInteger(columnIndex)
      || columnIndex < 0
      || columnIndex >= input.section.columns.length
      || (index > 0 && columnIndex <= columnIndexes[index - 1]!)
    ))
    || input.columns.some((column, index) => {
      const sectionColumn = input.section.columns[columnIndexes[index]!];
      return sectionColumn === undefined
        || column.inlineStartPt !== sectionColumn.xPt
        || column.inlineExtentPt !== sectionColumn.wPt;
    })) {
    throw new RangeError('Section region columns must equal its normalized section columns');
  }
}

function requireRect(rect: LayoutRect, name: string): void {
  if (!Number.isFinite(rect.xPt) || !Number.isFinite(rect.yPt)
    || !Number.isFinite(rect.widthPt) || !Number.isFinite(rect.heightPt)
    || rect.widthPt < 0 || rect.heightPt < 0) {
    throw new RangeError(`${name} must be a finite rectangle with non-negative extents`);
  }
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
  const regionIds = new Set<string>();
  const occurrenceIds = new Set<string>();
  let pageWritingMode: WritingMode | undefined;
  const occupiedPhysicalDomains: LayoutRect[] = [];

  for (const input of inputs) {
    requireIdentity(input.id, 'Section region id');
    requireIdentity(input.sectionOccurrenceId, 'Section occurrence id');
    if (regionIds.has(input.id) || occurrenceIds.has(input.sectionOccurrenceId)) {
      throw new RangeError('Section region and occurrence identities must be unique');
    }
    regionIds.add(input.id);
    occurrenceIds.add(input.sectionOccurrenceId);
    if (pageWritingMode !== undefined && pageWritingMode !== input.writingMode) {
      throw new RangeError('One physical page cannot mix writing modes');
    }
    pageWritingMode = input.writingMode;
    requireRegionSectionAgreement(input);
    const expectedPhysicalExtent = uprightPhysicalExtent({
      widthPt: input.section.geometry.pageWidth,
      heightPt: input.section.geometry.pageHeight,
    }, input.writingMode);
    if (expectedPhysicalExtent.widthPt !== physicalPage.widthPt
      || expectedPhysicalExtent.heightPt !== physicalPage.heightPt) {
      throw new RangeError(
        `Section regions on one physical page must use the same page box: expected ${expectedPhysicalExtent.widthPt}x${expectedPhysicalExtent.heightPt}, got ${physicalPage.widthPt}x${physicalPage.heightPt}`,
      );
    }
    const logicalExtent = logicalPageExtent(physicalPage, input.writingMode);
    const logicalInlineExtent = logicalExtent.widthPt;
    const logicalBlockExtent = logicalExtent.heightPt;
    if (!Number.isFinite(input.blockStartPt) || !Number.isFinite(input.blockEndPt)
      || input.blockStartPt < 0 || input.blockEndPt < input.blockStartPt
      || input.blockEndPt > logicalBlockExtent) {
      throw new RangeError('Section regions must be inside the logical page');
    }
    if (input.columns.length === 0) throw new RangeError('Section region must contain a column');
    const columnIndexes = input.columnIndexes
      ?? input.section.columns.map((_, index) => index);
    let priorInlineEndPt = 0;
    const coordinateSpace = createSectionRegionCoordinateSpace(input.writingMode, physicalPage);
    const flowDomainIds = input.columns.map((column, columnPosition) => {
      const authoredColumnIndex = columnIndexes[columnPosition]!;
      if (!Number.isFinite(column.inlineStartPt) || !Number.isFinite(column.inlineExtentPt)
        || column.inlineStartPt < 0 || column.inlineExtentPt <= 0
        || column.inlineStartPt + column.inlineExtentPt > logicalInlineExtent
        || column.inlineStartPt < priorInlineEndPt) {
        throw new RangeError('Columns must be ordered, disjoint, and inside the logical page');
      }
      priorInlineEndPt = column.inlineStartPt + column.inlineExtentPt;
      const id = bodyFlowDomainId(pageIndex, input.id, authoredColumnIndex);
      if (sectionByDomain.has(id)) throw new RangeError(`Duplicate flow domain ${id}`);
      const logicalBounds = {
        xPt: column.inlineStartPt,
        yPt: input.blockStartPt,
        widthPt: column.inlineExtentPt,
        heightPt: input.blockEndPt - input.blockStartPt,
      };
      const physicalBounds = transformRect(coordinateSpace.logicalToPhysical, logicalBounds);
      if (occupiedPhysicalDomains.some((prior) => (
        physicalBounds.xPt < prior.xPt + prior.widthPt
        && prior.xPt < physicalBounds.xPt + physicalBounds.widthPt
        && physicalBounds.yPt < prior.yPt + prior.heightPt
        && prior.yPt < physicalBounds.yPt + physicalBounds.heightPt
      ))) {
        throw new RangeError('Section flow domains on one page must be physically disjoint');
      }
      occupiedPhysicalDomains.push(physicalBounds);
      domains.push({
        id,
        kind: 'body',
        logicalBounds,
        physicalBounds,
      });
      sectionByDomain.set(id, input.sectionOccurrenceId);
      return id;
    });
    regions.push({
      id: input.id,
      sectionOccurrenceId: input.sectionOccurrenceId,
      coordinateSpace,
      blockStartPt: input.blockStartPt,
      blockEndPt: input.blockEndPt,
      columnFlowDirection: input.columnFlowDirection
        ?? (input.section.sectionBidi === true ? 'rtl' : 'ltr'),
      columnIndexes: Object.freeze([...columnIndexes]),
      flowDomainIds,
      section: input.section,
    });
  }

  return { regions, domains, sectionByDomain };
}

export function bodyOccurrenceDestinationFor(
  pageIndex: number,
  region: PageSectionRegionInput,
  columnIndex: number,
  blockStartPt: number,
  retainedFlowBounds: LayoutRect,
): BodyOccurrenceDestination {
  requirePageIndex(pageIndex);
  requireIdentity(region.id, 'Section region id');
  if (!Number.isInteger(columnIndex) || columnIndex < 0) {
    throw new RangeError('Column index must identify a section region column');
  }
  if (!Number.isFinite(blockStartPt)) throw new RangeError('Block start must be finite');
  requireRect(retainedFlowBounds, 'Retained flow bounds');
  const columnIndexes = region.columnIndexes
    ?? region.section.columns.map((_, index) => index);
  const columnPosition = columnIndexes.indexOf(columnIndex);
  if (columnPosition < 0 || columnPosition >= region.columns.length) {
    throw new RangeError('Column index must identify a section region column');
  }
  const column = region.columns[columnPosition]!;
  if (!Number.isFinite(column.inlineStartPt)) throw new RangeError('Column inline start must be finite');
  return {
    coordinateSpace: 'logical-page-points',
    flowDomainId: bodyFlowDomainId(pageIndex, region.id, columnIndex),
    translation: {
      xPt: column.inlineStartPt - retainedFlowBounds.xPt,
      yPt: blockStartPt - retainedFlowBounds.yPt,
    },
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
  textBox.story.blocks.forEach((block) => visitBookmarkParagraphs(block, visit));
}

export function derivePageBookmarkStarts(
  paint: readonly PaintNode[],
  defaultSectionOccurrenceId: string,
  sectionByDomain: ReadonlyMap<string, string>,
): readonly PageBookmarkStart[] {
  const starts: PageBookmarkStart[] = [];
  const seen = new Set<string>();
  for (const node of paint) {
    const sectionOccurrenceId = sectionByDomain.get(node.flowDomainId)
      ?? defaultSectionOccurrenceId;
    visitBookmarkParagraphs(node, (paragraph) => {
      for (const bookmark of paragraph.bookmarkStarts ?? []) {
        if (!bookmark || seen.has(bookmark)) continue;
        seen.add(bookmark);
        starts.push({
          name: bookmark,
          nodeId: paragraph.id,
          sectionOccurrenceId,
        });
      }
    });
  }
  return starts;
}

function bookmarkSectionByDomain(page: LayoutPage): ReadonlyMap<string, string> {
  const regions = page.sectionRegions ?? [];
  const regionById = new Map(regions.map((region) => [region.id, region]));
  const sectionByDomain = new Map<string, string>();
  for (const region of regions) {
    for (const domainId of region.flowDomainIds) {
      sectionByDomain.set(domainId, region.sectionOccurrenceId);
    }
  }
  for (const domain of page.flowDomains) {
    if (domain.kind !== 'footnote' && domain.kind !== 'endnote'
      && domain.sectionRegionId === undefined) continue;
    const region = domain.sectionRegionId
      ? regionById.get(domain.sectionRegionId)
      : regions[0];
    if (region) sectionByDomain.set(domain.id, region.sectionOccurrenceId);
  }
  return sectionByDomain;
}

export function deriveRetainedPageBookmarkStarts(
  page: LayoutPage,
  sectionByDomain: ReadonlyMap<string, string> = bookmarkSectionByDomain(page),
): readonly PageBookmarkStart[] {
  return derivePageBookmarkStarts(
    orderedPagePaintNodes(page),
    page.sectionOccurrenceId ?? '',
    sectionByDomain,
  );
}

/** Bookmark navigation metadata is a projection of the complete retained page
 * graph. Page stories and notes join that graph after initial body pagination,
 * so the final document boundary derives the projection again from the graph
 * that paint and navigation will actually retain. ECMA-376 §17.13.6.2 permits
 * bookmark starts in paragraph content regardless of the owning story. */
export function finalizePageBookmarkStarts(page: LayoutPage): LayoutPage {
  if (page.parityBlank) return page;
  return Object.freeze({
    ...page,
    bookmarkStarts: Object.freeze([...deriveRetainedPageBookmarkStarts(page)]),
  });
}

export function createLayoutPage(input: LayoutPageFactoryInput): LayoutPage {
  requirePageIndex(input.pageIndex);
  requireIdentity(input.sectionOccurrenceId, 'Page-start section occurrence id');
  const { regions, domains, sectionByDomain } = buildRegions(
    input.pageIndex,
    input.physicalPage,
    input.sectionRegions,
  );
  const firstRegion = input.sectionRegions[0];
  const pageBorders = firstRegion?.pageBorders ?? input.pageBorders;
  if (firstRegion !== undefined && (
    input.sectionOccurrenceId !== firstRegion.sectionOccurrenceId
    || !sectionLayoutContextsEqual(input.section, firstRegion.section)
  )) {
    throw new RangeError('Page-start section context must equal the first section region');
  }
  return {
    pageIndex: input.pageIndex,
    geometry: pageGeometry(input.physicalPage),
    flowDomains: domains,
    section: input.section,
    sectionOccurrenceId: input.sectionOccurrenceId,
    parityBlank: false,
    bookmarkStarts: derivePageBookmarkStarts(
      input.paint.map(({ node }) => node),
      input.sectionOccurrenceId,
      sectionByDomain,
    ),
    pageNumber: input.pageNumber,
    sectionRegions: regions,
    columnSeparators: columnSeparatorSegments(regions),
    pageBorder: materializePageBorderLayout(
      pageBorders,
      input.section,
      input.physicalPage,
      pageBorderFirstPageFlag(pageBorders, input.firstSectionOwnedPage),
    ),
    layers: createPageLayers(input.paint),
    readingOrder: input.readingOrder.map((node) => node.id),
  };
}

export function createLayoutPageAccumulator(
  input: LayoutPageAccumulatorInput,
): LayoutPageAccumulator {
  requirePageIndex(input.pageIndex);
  requireIdentity(input.sectionOccurrenceId, 'Page-start section occurrence id');
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
  firstSectionOwnedPage?: boolean,
): LayoutPage {
  return createLayoutPage({ ...accumulator, pageNumber, firstSectionOwnedPage });
}

export function createParityBlankLayoutPage(
  input: ParityBlankLayoutPageInput,
): LayoutPage {
  requirePageIndex(input.pageIndex);
  requireIdentity(input.sectionOccurrenceId, 'Page-start section occurrence id');
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
    columnSeparators: [],
    pageBorder: materializePageBorderLayout(
      input.pageBorders,
      input.section,
      input.physicalPage,
      pageBorderFirstPageFlag(input.pageBorders, input.firstSectionOwnedPage),
    ),
    layers: createPageLayers([]),
    readingOrder: [],
  };
}

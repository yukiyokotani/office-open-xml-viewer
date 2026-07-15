import { snapshotPlainData } from './plain-data.js';
import {
  translateCompleteParagraphLayout,
  translatePoint,
  translateRect,
} from './retained-geometry-translation.js';
import type {
  TableCellFragmentLayout,
  TableFragmentLayout,
  TableRowFragmentLayout,
} from './table-pagination.js';
import type {
  BorderSegment,
  DrawingLayout,
  FloatingTablePlacementLayout,
  LayoutRect,
  ParagraphLayout,
  ParagraphPlacement,
  ResolvedFloatingTablePlacementLayout,
  TableCellLayout,
  TableLayout,
  TableRowLayout,
  TextBoxLayout,
} from './types.js';

export interface LogicalBodyOccurrenceDestination {
  readonly coordinateSpace: 'logical-body-points';
  readonly flowDomainId: string;
  readonly translation: Readonly<{ xPt: number; yPt: number }>;
}

export interface BodyOccurrenceProjectionOptions {
  readonly occurrenceId: string;
  readonly destination: LogicalBodyOccurrenceDestination;
}

interface ProjectionContext {
  readonly options: BodyOccurrenceProjectionOptions;
  readonly ids: Map<string, string>;
  readonly anchorIds: Map<string, string>;
  readonly occurrenceIds: Map<string, string>;
  readonly cellDomains: Map<string, string>;
}

type Translation = LogicalBodyOccurrenceDestination['translation'];
const NO_TRANSLATION: Translation = Object.freeze({ xPt: 0, yPt: 0 });

function translatedBorder(border: BorderSegment, delta: Translation): BorderSegment {
  return {
    ...border,
    from: translatePoint(border.from, delta),
    to: translatePoint(border.to, delta),
  };
}

function encoded(value: string): string {
  return encodeURIComponent(value);
}

function mapped(
  values: Map<string, string>,
  namespace: string,
  sourceId: string,
  occurrenceId: string,
): string {
  const retained = values.get(sourceId);
  if (retained) return retained;
  const value = `${occurrenceId}/${namespace}/${encoded(sourceId)}`;
  values.set(sourceId, value);
  return value;
}

function nodeId(context: ProjectionContext, sourceId: string): string {
  return mapped(context.ids, 'node', sourceId, context.options.occurrenceId);
}

function anchorId(context: ProjectionContext, sourceId: string): string {
  return mapped(context.anchorIds, 'anchor', sourceId, context.options.occurrenceId);
}

function graphOccurrenceId(context: ProjectionContext, sourceId: string): string {
  return mapped(context.occurrenceIds, 'occurrence', sourceId, context.options.occurrenceId);
}

function nestedDomain(
  context: ProjectionContext,
  kind: 'cell' | 'textbox',
  sourceId: string,
): string {
  // Source IDs can recur in split/repeated occurrences. Including occurrence
  // identity prevents two retained cell/textbox flows from aliasing each other.
  return [
    context.options.destination.flowDomainId,
    'occurrence',
    encoded(context.options.occurrenceId),
    kind,
    encoded(sourceId),
  ].join('/');
}

function rekeyPlacement(
  placement: ParagraphPlacement,
  context: ProjectionContext,
): ParagraphPlacement {
  if (placement.kind === 'drawing') {
    return { ...placement, drawingId: nodeId(context, placement.drawingId) };
  }
  if (placement.kind === 'anchor-host' && placement.anchorOccurrenceId) {
    return {
      ...placement,
      anchorOccurrenceId: anchorId(context, placement.anchorOccurrenceId),
    };
  }
  return { ...placement };
}

function rekeyDrawing(
  drawing: DrawingLayout,
  flowDomainId: string,
  context: ProjectionContext,
): DrawingLayout {
  return {
    ...drawing,
    id: nodeId(context, drawing.id),
    flowDomainId,
    ...(drawing.textBoxIds ? {
      textBoxIds: drawing.textBoxIds.map((id) => nodeId(context, id)),
    } : {}),
    ...(drawing.anchorLayer ? {
      anchorLayer: {
        ...drawing.anchorLayer,
        occurrenceId: anchorId(context, drawing.anchorLayer.occurrenceId),
      },
    } : {}),
  };
}

function rekeyTextBox(
  textBox: TextBoxLayout,
  context: ProjectionContext,
): TextBoxLayout {
  const flowDomainId = nestedDomain(context, 'textbox', textBox.id);
  return {
    ...textBox,
    id: nodeId(context, textBox.id),
    flowDomainId,
    paragraphs: textBox.paragraphs.map((paragraph) => (
      rekeyParagraph(paragraph, flowDomainId, context)
    )),
  };
}

function rekeyParagraph(
  paragraph: ParagraphLayout,
  flowDomainId: string,
  context: ProjectionContext,
): ParagraphLayout {
  return {
    ...paragraph,
    id: nodeId(context, paragraph.id),
    flowDomainId,
    lines: paragraph.lines.map((line) => ({
      ...line,
      placements: line.placements.map((placement) => rekeyPlacement(placement, context)),
    })),
    drawings: paragraph.drawings.map((drawing) => (
      rekeyDrawing(drawing, flowDomainId, context)
    )),
    textBoxes: paragraph.textBoxes.map((textBox) => rekeyTextBox(textBox, context)),
    exclusions: paragraph.exclusions.map((exclusion) => ({
      ...exclusion,
      id: nodeId(context, exclusion.id),
      ...(exclusion.anchorOccurrenceId ? {
        anchorOccurrenceId: anchorId(context, exclusion.anchorOccurrenceId),
      } : {}),
    })),
    ...(paragraph.anchorFrames ? {
      anchorFrames: paragraph.anchorFrames.map((frame) => ({
        ...frame,
        occurrenceId: anchorId(context, frame.occurrenceId),
      })),
    } : {}),
  };
}

function projectParagraph(
  paragraph: ParagraphLayout,
  flowDomainId: string,
  delta: Translation,
  context: ProjectionContext,
): ParagraphLayout {
  return rekeyParagraph(
    translateCompleteParagraphLayout(paragraph, delta),
    flowDomainId,
    context,
  );
}

function translatedBase<T extends { flowBounds: LayoutRect; inkBounds: LayoutRect; clipBounds?: LayoutRect }>(
  value: T,
  delta: Translation,
): T {
  return {
    ...value,
    flowBounds: translateRect(value.flowBounds, delta),
    inkBounds: translateRect(value.inkBounds, delta),
    ...(value.clipBounds ? { clipBounds: translateRect(value.clipBounds, delta) } : {}),
  };
}

function projectCell(
  cell: TableCellLayout,
  delta: Translation,
  context: ProjectionContext,
): TableCellLayout {
  const flowDomainId = nestedDomain(context, 'cell', cell.id);
  context.cellDomains.set(cell.id, flowDomainId);
  return {
    ...translatedBase(cell, delta),
    id: nodeId(context, cell.id),
    flowDomainId,
    contentBounds: translateRect(cell.contentBounds, delta),
    blocks: cell.blocks.map((block) => ({
      ...block,
      // Cell blocks remain in their cell-local coordinates. Canvas table paint
      // applies contentBounds/offsetPt; inheriting the outer delta here would
      // double-move nested-table alignment retained in child.flowBounds.
      layout: block.layout.kind === 'paragraph'
        ? projectParagraph(block.layout, flowDomainId, NO_TRANSLATION, context)
        : projectTable(block.layout, flowDomainId, NO_TRANSLATION, context),
    })),
  };
}

function projectRow(
  row: TableRowLayout,
  flowDomainId: string,
  delta: Translation,
  context: ProjectionContext,
): TableRowLayout {
  const projected = {
    ...translatedBase(row, delta),
    id: nodeId(context, row.id),
    flowDomainId,
    cells: row.cells.map((cell) => projectCell(cell, delta, context)),
  };
  if (!isFragmentRow(row)) return projected;
  const fragmentRow: TableRowFragmentLayout = {
    ...projected,
    logicalRowIndex: row.logicalRowIndex,
    fragmentIndex: row.fragmentIndex,
    ownership: row.ownership,
    occurrenceId: graphOccurrenceId(context, row.occurrenceId),
    physicalPageIndex: row.physicalPageIndex,
    displayPageNumber: row.displayPageNumber,
    cells: projected.cells.map((cell, index): TableCellFragmentLayout => ({
      ...cell,
      contentRanges: [...row.cells[index]!.contentRanges],
      ...(row.cells[index]!.visualMergeOwnership ? {
        visualMergeOwnership: row.cells[index]!.visualMergeOwnership,
      } : {}),
    })),
  };
  return fragmentRow;
}

function isFragmentRow(row: TableRowLayout): row is TableRowFragmentLayout {
  return 'occurrenceId' in row && typeof row.occurrenceId === 'string';
}

function projectFloating(
  placement: FloatingTablePlacementLayout,
  delta: Translation,
  context: ProjectionContext,
): FloatingTablePlacementLayout {
  const hostDomain = context.cellDomains.get(placement.hostCellId)
    ?? nestedDomain(context, 'cell', placement.hostCellId);
  return {
    ...placement,
    occurrenceId: graphOccurrenceId(context, placement.occurrenceId),
    hostCellId: nodeId(context, placement.hostCellId),
    tableId: nodeId(context, placement.tableId),
    ...(placement.columnBounds ? {
      columnBounds: translateRect(placement.columnBounds, delta),
    } : {}),
    anchorBounds: translateRect(placement.anchorBounds, delta),
    child: projectTable(placement.child, hostDomain, NO_TRANSLATION, context),
  };
}

function projectResolvedFloating(
  placement: ResolvedFloatingTablePlacementLayout,
  delta: Translation,
  context: ProjectionContext,
): ResolvedFloatingTablePlacementLayout {
  // The transaction has already resolved these boxes in page-local points.
  // Only graph identity changes; the child stays anchor-local for paintPlacedChild.
  const source = projectFloating(placement.source, NO_TRANSLATION, context);
  const hostDomain = context.cellDomains.get(placement.source.hostCellId)
    ?? nestedDomain(context, 'cell', placement.source.hostCellId);
  return {
    ...placement,
    occurrenceId: graphOccurrenceId(context, placement.occurrenceId),
    child: projectTable(placement.child, hostDomain, NO_TRANSLATION, context),
    source,
  };
}

function projectTable(
  table: TableLayout,
  flowDomainId: string,
  delta: Translation,
  context: ProjectionContext,
): TableLayout {
  const projected: TableLayout = {
    ...translatedBase(table, delta),
    id: nodeId(context, table.id),
    flowDomainId,
    borders: table.borders.map((border) => translatedBorder(border, delta)),
    rows: table.rows.map((row) => projectRow(row, flowDomainId, delta, context)),
  };
  if (!('floatingTables' in table) || !('resolvedFloatingTables' in table)) {
    return projected;
  }
  const fragment = table as TableFragmentLayout;
  return {
    ...projected,
    rows: projected.rows as readonly TableRowFragmentLayout[],
    floatingTables: fragment.floatingTables.map((placement) => (
      projectFloating(placement, delta, context)
    )),
    resolvedFloatingTables: fragment.resolvedFloatingTables.map((placement) => (
      projectResolvedFloating(placement, delta, context)
    )),
    ...(fragment.floatingTableCoordinateSpace ? {
      floatingTableCoordinateSpace: fragment.floatingTableCoordinateSpace,
    } : {}),
  } as TableFragmentLayout;
}

/** SourceRef remains authored identity while every layout ID is occurrence-local.
 * Projection stays in logical body points; region-to-physical vertical transforms
 * belong to the page boundary and must not be applied here. */
export function projectBodyOccurrence<T extends ParagraphLayout | TableLayout>(
  retained: T,
  options: BodyOccurrenceProjectionOptions,
): T {
  if (options.occurrenceId.length === 0 || options.destination.flowDomainId.length === 0) {
    throw new RangeError('Occurrence and destination flow-domain IDs must not be empty');
  }
  if (!Number.isFinite(options.destination.translation.xPt)
    || !Number.isFinite(options.destination.translation.yPt)) {
    throw new RangeError('Occurrence translation must contain finite logical point values');
  }
  const context: ProjectionContext = {
    options,
    ids: new Map(),
    anchorIds: new Map(),
    occurrenceIds: new Map(),
    cellDomains: new Map(),
  };
  const projected = retained.kind === 'paragraph'
    ? projectParagraph(
        retained,
        options.destination.flowDomainId,
        options.destination.translation,
        context,
      )
    : projectTable(
        retained,
        options.destination.flowDomainId,
        options.destination.translation,
        context,
      );
  return snapshotPlainData(projected, 'Body occurrence projection') as T;
}

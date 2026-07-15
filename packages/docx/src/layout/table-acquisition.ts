import type {
  BorderSpec,
  CellBorders,
  DocParagraph,
  DocTable,
  TableBorders,
} from '../types.js';
import {
  acquireTableCellBlocks,
  isStructuralTrailingParagraph,
} from './table-cell-blocks.js';
import type { ParagraphBorderEdges } from './paragraph-border-adjacency.js';
import { layoutTable } from './table.js';
import { tableCellHorizontalSpacingInsets } from './table-columns.js';
import { snapshotPlainData } from './plain-data.js';
import type {
  FloatingTablePositionInput,
  LayoutServices,
  LayoutNodeId,
  ParagraphLayout,
  TableBorderInput,
  TableEdgeInputs,
  TableFormatInput,
  TableLayout,
  TableLayoutInput,
} from './types.js';

export interface RetainedTableAcquisitionDependencies<State> {
  layoutServices(state: State): LayoutServices | undefined;
  tableFormat(table: Readonly<DocTable>): TableFormatInput;
  resolveColumns(table: DocTable, contentWidthPt: number, state: State): readonly number[];
  createCellState(state: State, contentWidthPt: number, cell: DocTable['rows'][number]['cells'][number]): State;
  acquireParagraph(
    state: State,
    paragraph: DocParagraph,
    contentWidthPt: number,
    sourcePath: readonly number[],
    flowDomainId: string,
    paragraphBorderEdges: ParagraphBorderEdges,
  ): ParagraphLayout;
  registerFloatingTable(
    state: State,
    request: Readonly<{
      child: TableLayout;
      positioning: FloatingTablePositionInput;
      overlap: 'never' | 'overlap';
    }>,
  ): Readonly<{ xPt: number; yPt: number }> | null;
  advanceState(state: State, advancePt: number): void;
}

/**
 * The finished geometry and the immutable semantic input that produced it travel
 * together. Pagination may derive page-local row and border geometry from the
 * input, while ordinary paint consumes the finished layout without measuring.
 */
export interface RetainedTableAcquisition {
  readonly input: TableLayoutInput;
  readonly layout: TableLayout;
  readonly nestedById: Readonly<Record<string, RetainedTableAcquisition>>;
  readonly floatingTables: readonly NestedFloatingTableOccurrence[];
}

/**
 * A cell-owned out-of-flow occurrence. The retained table is referenced by id
 * rather than embedded again so layout and paint cannot count the same child as
 * both an ordinary block and a floating placement.
 */
export interface NestedFloatingTableOccurrence {
  readonly hostCellId: LayoutNodeId;
  readonly sourceBlockIndex: number;
  readonly anchorBlockIndex: number;
  readonly tableId: LayoutNodeId;
  readonly overlap: 'never' | 'overlap';
  readonly positioning: FloatingTablePositionInput;
  readonly acquiredTextOffsetPt?: Readonly<{ xPt: number; yPt: number }>;
}

function nextRegularParagraphIndex(
  content: DocTable['rows'][number]['cells'][number]['content'],
  afterIndex: number,
): number {
  const anchorBlockIndex = content.findIndex((element, index) => (
    index > afterIndex
      && element.type === 'paragraph'
      && element.framePr == null
  ));
  if (anchorBlockIndex < 0) {
    throw new Error('A nested floating table requires a following regular paragraph anchor');
  }
  return anchorBlockIndex;
}

function retainedBorder(border: BorderSpec | null): TableBorderInput | null {
  if (!border) return null;
  const authored = border.color ?? '000000';
  return Object.freeze({
    widthPt: border.width,
    color: authored.startsWith('#') ? authored : `#${authored}`,
    authoredStyle: border.style,
  });
}

function retainedEdges(edges: CellBorders | TableBorders): TableEdgeInputs {
  return Object.freeze({
    top: retainedBorder(edges.top),
    right: retainedBorder(edges.right),
    bottom: retainedBorder(edges.bottom),
    left: retainedBorder(edges.left),
    insideH: retainedBorder(edges.insideH),
    insideV: retainedBorder(edges.insideV),
  });
}

function physicalAlignment(
  value: string | null | undefined,
  bidiVisual: boolean,
): TableLayoutInput['alignment'] {
  if (value === 'center') return 'center';
  const trailing = value === 'right' || value === 'end';
  return (bidiVisual ? !trailing : trailing) ? 'right' : 'left';
}

function paragraphHasPageDependency(layout: ParagraphLayout): boolean {
  return layout.lines.some((line) => line.placements.some((placement) => (
    placement.kind === 'text' && placement.dependency === 'page'
  )));
}

/**
 * Acquire an ordinary or nested table from final-width retained children.
 * Parser-private authored-presence and lexical facts arrive only through the
 * immutable TableFormatInput; this fold owns recursive table geometry and never
 * reaches back into parser/model metadata.
 */
export function acquireRetainedTable<State>(
  table: DocTable,
  columnWidthsPt: readonly number[],
  contentWidthPt: number,
  outerState: State,
  sourcePath: readonly number[],
  dependencies: RetainedTableAcquisitionDependencies<State>,
): RetainedTableAcquisition {
  const services = dependencies.layoutServices(outerState);
  if (!services) throw new Error('Retained table acquisition requires layout services');
  const flowDomainId = `table:${sourcePath.join('.')}`;
  const format = dependencies.tableFormat(table);
  const bidiVisual = table.bidiVisual === true;
  const firstRowException = format.firstRowException;
  const tableIndentPt = firstRowException?.indentAuthored
    ? (firstRowException.indentPt ?? 0)
    : (table.tblInd ?? 0);
  const nestedById: Record<string, RetainedTableAcquisition> = {};
  const floatingTables: NestedFloatingTableOccurrence[] = [];
  const rows: TableLayoutInput['rows'] = table.rows.map((row, rowIndex) => {
    const rowFormat = format.rows[rowIndex];
    let columnStart = Math.max(0, Math.min(columnWidthsPt.length, row.gridBefore ?? 0));
    const cells = row.cells.map((cell, cellIndex) => {
      const formatMargins = rowFormat?.cells[cellIndex]?.marginsPt ?? {
        top: cell.marginTop ?? table.cellMarginTop,
        right: cell.marginRight ?? table.cellMarginRight,
        bottom: cell.marginBottom ?? table.cellMarginBottom,
        left: cell.marginLeft ?? table.cellMarginLeft,
      };
      const currentColumnStart = columnStart;
      const columnSpan = Math.min(
        Math.max(1, cell.colSpan),
        Math.max(0, columnWidthsPt.length - currentColumnStart),
      );
      columnStart += columnSpan;
      const cellTotalWidthPt = columnWidthsPt
        .slice(currentColumnStart, currentColumnStart + columnSpan)
        .reduce((sum, width) => sum + width, 0);
      const spacingInsets = tableCellHorizontalSpacingInsets(
        rowFormat?.cellSpacingPt ?? 0,
        currentColumnStart,
        columnSpan,
        columnWidthsPt.length,
      );
      const cellPath = [...sourcePath, rowIndex, cellIndex];
      const cellId = `${flowDomainId}:cell:${rowIndex}.${cellIndex}`;
      const acquired = cell.vMerge === false
        ? []
        : acquireTableCellBlocks({
            cell,
            table,
            cellTotalWidthPt,
            outerState,
            sourcePath: cellPath,
          }, {
            resolveContentWidthPt: (_cell, _table, totalWidthPt) => Math.max(
              0,
              totalWidthPt
                - spacingInsets.startPt
                - spacingInsets.endPt
                - formatMargins.left
                - formatMargins.right,
            ),
            createCellState: dependencies.createCellState,
            acquireParagraph: (
              cellState,
              paragraph,
              paragraphWidthPt,
              paragraphPath,
              paragraphBorderEdges,
            ) => dependencies.acquireParagraph(
              cellState,
              paragraph,
              paragraphWidthPt,
              paragraphPath,
              `${flowDomainId}:cell:${rowIndex}.${cellIndex}`,
              paragraphBorderEdges,
            ),
            acquireNestedTable: (cellState, nestedTable, nestedContentWidthPt, nestedPath) => {
              const nestedColumns = dependencies.resolveColumns(
                nestedTable,
                nestedContentWidthPt,
                cellState,
              );
              const nested = acquireRetainedTable(
                nestedTable,
                nestedColumns,
                nestedContentWidthPt,
                cellState,
                nestedPath,
                dependencies,
              );
              nestedById[nested.layout.id] = nested;
              const nestedFormat = dependencies.tableFormat(nestedTable);
              const effectivePositioning = nestedFormat.positioning;
              if (effectivePositioning) {
                const sourceBlockIndex = nestedPath[nestedPath.length - 1]!;
                const positioning = effectivePositioning;
                const overlap = nestedTable.overlap === 'never' ? 'never' : 'overlap';
                const acquiredTextOffsetPt = dependencies.registerFloatingTable(cellState, {
                  child: nested.layout,
                  positioning,
                  overlap,
                });
                const occurrence = {
                  hostCellId: cellId,
                  sourceBlockIndex,
                  anchorBlockIndex: nextRegularParagraphIndex(cell.content, sourceBlockIndex),
                  tableId: nested.layout.id,
                  overlap,
                  positioning,
                  ...(acquiredTextOffsetPt == null ? {} : {
                    acquiredTextOffsetPt: Object.freeze({ ...acquiredTextOffsetPt }),
                  }),
                } as const;
                floatingTables.push(occurrence);
              }
              return nested.layout;
            },
            advanceState: dependencies.advanceState,
          });
      return {
        id: cellId,
        source: { story: 'body' as const, storyInstance: 'body', path: cellPath },
        columnStart: currentColumnStart,
        columnSpan,
        verticalMerge: cell.vMerge === true
          ? 'restart' as const
          : cell.vMerge === false ? 'continue' as const : 'none' as const,
        margins: {
          topPt: formatMargins.top,
          rightPt: formatMargins.right,
          bottomPt: formatMargins.bottom,
          leftPt: formatMargins.left,
        },
        vAlign: cell.vAlign,
        ...(cell.background ? {
          background: {
            color: cell.background.startsWith('#') ? cell.background : `#${cell.background}`,
          },
        } : {}),
        borders: retainedEdges(cell.borders),
        blocks: acquired.flatMap((layout, sourceBlockIndex) => {
          const sourceElement = cell.content[sourceBlockIndex];
          // ECMA-376 §17.4.57 keeps tblpPr tables at their logical source
          // position only for anchoring; they do not participate in cell flow.
          if (
            sourceElement?.type === 'table'
            && dependencies.tableFormat(sourceElement).ordinaryFlow === false
          ) return [];
          return [{
            layout,
            sourceBlockIndex,
            ...((layout.kind === 'paragraph' && paragraphHasPageDependency(layout))
              ? { pageDependent: true }
              : {}),
            ...(isStructuralTrailingParagraph(cell.content, sourceBlockIndex)
              ? { structuralTrailing: true }
              : {}),
          }];
        }),
      };
    });
    const heightRule = rowFormat?.height?.rule ?? 'auto';
    return {
      id: `${flowDomainId}:row:${rowIndex}`,
      source: { story: 'body' as const, storyInstance: 'body', path: [...sourcePath, rowIndex] },
      logicalRowIndex: rowIndex,
      cantSplit: rowFormat?.cantSplit ?? row.cantSplit === true,
      heightPt: rowFormat?.height?.valuePt ?? null,
      heightRule,
      cellSpacingPt: rowFormat?.cellSpacingPt ?? 0,
      exceptionBorders: rowFormat?.exception?.borders
        ? retainedEdges(rowFormat.exception.borders)
        : null,
      alignment: physicalAlignment(rowFormat?.justification ?? table.jc, bidiVisual),
      indentPt: tableIndentPt,
      cells,
      repeatedHeader: rowFormat?.repeatedHeader ?? row.isHeader === true,
    };
  });
  const input = snapshotPlainData<TableLayoutInput>({
    kind: 'table',
    id: flowDomainId,
    source: { story: 'body', storyInstance: 'body', path: [...sourcePath] },
    flowDomainId,
    ordinaryFlow: format.ordinaryFlow,
    alignment: physicalAlignment(table.jc, bidiVisual),
    indentPt: tableIndentPt,
    bidiVisual,
    columnWidthsPt,
    borders: retainedEdges(table.borders),
    rows,
  }, 'RetainedTableAcquisition.input') as TableLayoutInput;
  const bounds = {
    xPt: 0,
    yPt: 0,
    widthPt: contentWidthPt,
    heightPt: 1,
  };
  const layout = layoutTable(input, {
    container: { id: flowDomainId, kind: 'tableCell', bounds },
    cursor: { xPt: 0, yPt: 0 },
    availableBounds: bounds,
  }, services).layout;
  return Object.freeze({
    input,
    layout,
    nestedById: Object.freeze(nestedById),
    floatingTables: snapshotPlainData(
      floatingTables,
      'RetainedTableAcquisition.floatingTables',
    ) as readonly NestedFloatingTableOccurrence[],
  });
}

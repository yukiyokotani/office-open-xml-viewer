import type {
  BorderSpec,
  CellBorders,
  TableBorders,
} from '../types.js';
import type { ParagraphLayoutSource } from './text.js';
import type { TableLayoutSource } from './table-source-acquisition.js';
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
  DrawingMLCollisionEntryPt,
  LayoutServices,
  LayoutNodeId,
  PaintNode,
  ParagraphLayout,
  TableBorderInput,
  TableEdgeInputs,
  TableFormatInput,
  TableLayout,
  TableLayoutInput,
  SourceRef,
  WrapExclusion,
} from './types.js';

export interface RetainedTableAcquisitionDependencies<State> {
  layoutServices(state: State): LayoutServices | undefined;
  tableFormat(table: TableLayoutSource): TableFormatInput;
  resolveColumns(table: TableLayoutSource, contentWidthPt: number, state: State): readonly number[];
  createCellState(state: State, contentWidthPt: number, cell: TableLayoutSource['rows'][number]['cells'][number]): State;
  acquireParagraph(
    state: State,
    paragraph: ParagraphLayoutSource,
    contentWidthPt: number,
    sourcePath: readonly number[],
    flowDomainId: string,
    paragraphBorderEdges?: ParagraphBorderEdges,
    inheritedAuthority?: Readonly<{
      exclusions: readonly WrapExclusion[];
      collisions: readonly DrawingMLCollisionEntryPt[];
    }>,
    source?: SourceRef,
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

function retainedNodeIsReusableAcrossPages(
  node: PaintNode,
  visited: Set<PaintNode>,
): boolean {
  if (visited.has(node)) return true;
  visited.add(node);
  if (node.kind === 'drawing') return node.anchorLayer === undefined;
  if (node.kind === 'paragraph') {
    return node.lines.every((line) => line.placements.every((placement) => (
      placement.kind !== 'text' || placement.dependency !== 'page'
    )))
      && node.drawings.every((drawing) => (
        retainedNodeIsReusableAcrossPages(drawing, visited)
      ))
      && node.textBoxes.every((textBox) => (
        retainedNodeIsReusableAcrossPages(textBox, visited)
      ));
  }
  if (node.kind === 'textbox' || node.kind === 'note') {
    return node.story.blocks.every((block) => (
      retainedNodeIsReusableAcrossPages(block, visited)
    ));
  }
  return node.rows.every((row) => row.cells.every((cell) => (
    cell.blocks.every((block) => (
      retainedNodeIsReusableAcrossPages(block.layout, visited)
    ))
  )))
    && (node.floatingTables ?? []).every((placement) => (
      retainedNodeIsReusableAcrossPages(placement.child, visited)
    ))
    && (node.resolvedFloatingTables ?? []).every((placement) => (
      retainedNodeIsReusableAcrossPages(placement.child, visited)
    ));
}

/**
 * Whether a retained acquisition can serve every page of a layout session
 * unchanged at the same inline extent. Two classes of baked geometry vary by
 * destination page:
 *
 * - PAGE-field (ECMA-376 §17.16.5.44) text. Blocks carrying it are flagged
 *   `pageDependent` and re-acquired per destination page during pagination
 *   (TableFragmentContext.reacquirePageDependentBlock), but only on paths that
 *   provide that hook, so a reusable acquisition must not contain them.
 * - Anchored drawings, whose reference frames (including page parity for
 *   inside/outside alignment) are resolved against the acquisition-time page.
 *
 * The remaining folded inputs (note numbers, numbering markers, current date)
 * are constant within one body layout session. Plain retained geometry stays
 * table/cell-relative; the graph walk below rejects page/section-sensitive
 * fields and anchors even when they are nested in a text-box story.
 */
function retainedTableAcquisitionGraphIsReusableAcrossPages(
  acquisition: RetainedTableAcquisition,
  visited: Set<PaintNode>,
): boolean {
  const rowsAreReusable = acquisition.input.rows.every((row) => (
    row.cells.every((cell) => cell.blocks.every((block) => (
      block.pageDependent !== true
      && retainedNodeIsReusableAcrossPages(block.layout, visited)
    )))
  ));
  return rowsAreReusable
    && Object.values(acquisition.nestedById).every(
      (nested) => retainedTableAcquisitionGraphIsReusableAcrossPages(nested, visited),
    );
}

export function retainedTableAcquisitionIsReusableAcrossPages(
  acquisition: RetainedTableAcquisition,
): boolean {
  return retainedTableAcquisitionGraphIsReusableAcrossPages(
    acquisition,
    new Set<PaintNode>(),
  );
}

function nextRegularParagraphIndex(
  content: TableLayoutSource['rows'][number]['cells'][number]['content'],
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
  table: TableLayoutSource,
  columnWidthsPt: readonly number[],
  contentWidthPt: number,
  outerState: State,
  source: SourceRef | readonly number[],
  dependencies: RetainedTableAcquisitionDependencies<State>,
): RetainedTableAcquisition {
  const sourceRoot: SourceRef = Array.isArray(source)
    ? { story: 'body', storyInstance: 'body', path: source }
    : source as SourceRef;
  const sourcePath = sourceRoot.path;
  const sourceAt = (path: readonly number[]): SourceRef => ({
    story: sourceRoot.story,
    storyInstance: sourceRoot.storyInstance,
    path,
  });
  const services = dependencies.layoutServices(outerState);
  if (!services) throw new Error('Retained table acquisition requires layout services');
  const flowDomainId = sourceRoot.story === 'body' && sourceRoot.storyInstance === 'body'
    ? `table:${sourcePath.join('.')}`
    : `${sourceRoot.story}:${sourceRoot.storyInstance}:table:${sourcePath.join('.')}`;
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
              undefined,
              sourceAt(paragraphPath),
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
                sourceAt(nestedPath),
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
        source: sourceAt(cellPath),
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
      source: sourceAt([...sourcePath, rowIndex]),
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
    source: sourceAt([...sourcePath]),
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

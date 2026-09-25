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
import { layoutTable, measureTableCellBlockFlowHeightPt } from './table.js';
import { tableCellHorizontalSpacingInsets } from './table-columns.js';
import { snapshotPlainData } from './plain-data.js';
import { eastAsianUprightPaintOps } from './vertical-glyph-orientation.js';
import type {
  FloatingTablePositionInput,
  DrawingMLCollisionEntryPt,
  LayoutServices,
  LayoutNodeId,
  PaintNode,
  ParagraphLayout,
  TableBorderInput,
  TableCellVerticalMode,
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

/** The largest page dimension Word can author (MS-DOC 2.6.4 sprmSXaPage /
 * sprmSYaPage: at most 31680 twips = 1584pt). A rotated line can never be
 * longer than a page, so this bound serves only as the unconstrained line
 * length for measuring a rotated cell's natural line extent. */
const MAXIMUM_ROTATED_LINE_LENGTH_PT = 1584;
const ROTATED_LINE_LENGTH_EPSILON_PT = 0.01;

interface RotatedCellAcquisition {
  readonly rowIndex: number;
  readonly cellIndex: number;
  readonly margins: Readonly<{ top: number; bottom: number }>;
  reacquire(lineLengthPt: number): TableLayoutInput['rows'][number]['cells'][number]['blocks'];
}

/**
 * ECMA-376 §17.4.72 cell text direction projected onto the rotated-frame
 * modes this renderer paints. tbRl and btLr rotate the whole text frame a
 * quarter turn; tbRlV additionally keeps East Asian glyphs upright (the
 * DrawingML eaVert projection). lrTbV (horizontal lines with rotated East
 * Asian glyphs) and tbLrV (vertical lines advancing left to right) are not
 * rendered rotated yet, and a cell containing a nested table keeps
 * horizontal layout because a nested table cannot be re-acquired along the
 * rotated line axis. Those cells lay out horizontally.
 */
function verticalCellMode(
  cell: TableLayoutSource['rows'][number]['cells'][number],
): TableCellVerticalMode | undefined {
  if (cell.content.some((element) => element.type === 'table')) return undefined;
  switch (cell.textDirection) {
    case 'tbRl': return 'vert';
    case 'btLr': return 'vert270';
    case 'tbRlV': return 'eaVert';
    default: return undefined;
  }
}

/**
 * Line extent of rotated cell content at an acquired line width: each line's
 * placement span (independent of its alignment on that
 * line) plus the paragraph's side indents and, on the first line, a positive
 * first-line indent. The line box advance is also a minimum along the rotated
 * row axis, even when a glyph is narrower; see WORD_ROTATED_CELL_AUTO_ROW_WRAP.
 * The unconstrained extent is an upper bound on the row height needed to fit
 * the rotated content; it is not the auto-row minimum.
 */
function naturalLineExtentPt(
  layouts: readonly (ParagraphLayout | TableLayout)[],
  content: TableLayoutSource['rows'][number]['cells'][number]['content'],
): number {
  let extentPt = 0;
  layouts.forEach((layout, index) => {
    if (layout.kind !== 'paragraph') {
      extentPt = Math.max(extentPt, layout.flowBounds.widthPt);
      return;
    }
    const source = content[index];
    const paragraph = source?.type === 'paragraph' ? source : undefined;
    const sideIndentsPt = Math.max(0, paragraph?.indentLeft ?? 0)
      + Math.max(0, paragraph?.indentRight ?? 0);
    layout.lines.forEach((line, lineIndex) => {
      extentPt = Math.max(extentPt, line.advancePt);
      let startPt = Number.POSITIVE_INFINITY;
      let endPt = Number.NEGATIVE_INFINITY;
      for (const placement of line.placements) {
        if (!('bounds' in placement) || !placement.bounds) continue;
        startPt = Math.min(startPt, placement.bounds.xPt);
        endPt = Math.max(endPt, placement.bounds.xPt + placement.bounds.widthPt);
      }
      if (endPt < startPt) return;
      const firstLinePt = lineIndex === 0 ? Math.max(0, paragraph?.indentFirst ?? 0) : 0;
      extentPt = Math.max(extentPt, endPt - startPt + sideIndentsPt + firstLinePt);
    });
  });
  return Math.min(MAXIMUM_ROTATED_LINE_LENGTH_PT, Math.ceil(extentPt * 100) / 100);
}

/** Find the shortest rotated line axis that fits all resulting columns within
 * the physical cell width. ECMA-376 §17.4.72 rotates the text frame and
 * §17.4.80 lets auto/atLeast rows grow to fit their content. The
 * compatibility observation and its tested bounds are registered as
 * WORD_ROTATED_CELL_AUTO_ROW_WRAP in table-compatibility.ts.
 *
 * Most cells need only the narrowest and unconstrained acquisitions. Search
 * only when the narrowest columns overflow the cell width; bisection is
 * bounded by the existing 0.01pt line-length resolution.
 */
function fittingRotatedLineLengthPt(
  minimumPt: number,
  contentWidthPt: number,
  acquire: (lineLengthPt: number) => TableLayoutInput['rows'][number]['cells'][number]['blocks'],
  maximumPt: () => number,
): number {
  const fits = (lengthPt: number) =>
    measureTableCellBlockFlowHeightPt(acquire(lengthPt)) <= contentWidthPt + ROTATED_LINE_LENGTH_EPSILON_PT;
  if (fits(minimumPt)) return minimumPt;
  const unconstrainedPt = Math.max(minimumPt, maximumPt());
  if (!fits(unconstrainedPt)) return unconstrainedPt;
  let lowerPt = minimumPt;
  let upperPt = unconstrainedPt;
  while (upperPt - lowerPt > ROTATED_LINE_LENGTH_EPSILON_PT) {
    const middlePt = Math.floor((lowerPt + upperPt) * 50) / 100;
    if (middlePt <= lowerPt || middlePt >= upperPt) break;
    if (fits(middlePt)) upperPt = middlePt;
    else lowerPt = middlePt;
  }
  return upperPt;
}

/** eaVert keeps East Asian clusters upright inside the rotated frame, as the
 * DrawingML vertical text boxes do (shared projection). */
function orientRotatedCellBlocks(
  layouts: readonly (ParagraphLayout | TableLayout)[],
  mode: TableCellVerticalMode,
): (ParagraphLayout | TableLayout)[] {
  if (mode !== 'eaVert') return [...layouts];
  return layouts.map((layout) => layout.kind !== 'paragraph' ? layout : {
    ...layout,
    lines: layout.lines.map((line) => ({
      ...line,
      placements: line.placements.map((placement) => placement.kind === 'text'
        ? { ...placement, paintOps: eastAsianUprightPaintOps(placement) }
        : placement),
    })),
  });
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
  const rotatedCells: RotatedCellAcquisition[] = [];
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
      // Same grouped insets as the horizontal content width below.
      const physicalContentWidthPt = Math.max(
        0,
        cellTotalWidthPt
          - (spacingInsets.startPt + spacingInsets.endPt)
          - (formatMargins.left + formatMargins.right),
      );
      const verticalMode = cell.vMerge === false ? undefined : verticalCellMode(cell);
      const acquireAt = (lineWidthPt: number | undefined) => cell.vMerge === false
        ? []
        : acquireTableCellBlocks({
            cell,
            table,
            cellTotalWidthPt,
            outerState,
            sourcePath: cellPath,
          }, {
            // Match the grouped insets added to the intrinsic AutoFit minimum
            // (intrinsic-width.ts and table-source-acquisition.ts). Reversing
            // those groups avoids rounding an exact measured-width boundary
            // below its own minimum. This preserves the measured boundary
            // without adding a width allowance. Margin ownership is
            // ECMA-376 §17.4.41/.42. A rotated cell is re-acquired along its
            // rotated line axis instead (see verticalCellMode).
            resolveContentWidthPt: (_cell, _table, totalWidthPt) => lineWidthPt ?? Math.max(
              0,
              totalWidthPt
                - (spacingInsets.startPt + spacingInsets.endPt)
                - (formatMargins.left + formatMargins.right),
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
      let acquired: ReturnType<typeof acquireAt>;
      let verticalText: TableLayoutInput['rows'][number]['cells'][number]['verticalText'];
      if (verticalMode) {
        const rowRule = rowFormat?.height?.rule ?? 'auto';
        const rowHeightPt = rowFormat?.height?.valuePt ?? 0;
        const authoredLengthPt = Math.max(0, rowHeightPt - formatMargins.top - formatMargins.bottom);
        const minimumLengthPt = rowRule === 'exact'
          ? authoredLengthPt
          : Math.max(
              authoredLengthPt,
              naturalLineExtentPt(acquireAt(0), cell.content),
            );
        const lineLengthPt = rowRule === 'exact'
          ? authoredLengthPt
          : fittingRotatedLineLengthPt(
              minimumLengthPt,
              physicalContentWidthPt,
              (lengthPt) => cellBlocks(acquireAt(lengthPt)),
              () => naturalLineExtentPt(acquireAt(MAXIMUM_ROTATED_LINE_LENGTH_PT), cell.content),
            );
        acquired = orientRotatedCellBlocks(acquireAt(lineLengthPt), verticalMode);
        verticalText = { mode: verticalMode, lineLengthPt, requiredLineLengthPt: lineLengthPt };
        rotatedCells.push({
          rowIndex,
          cellIndex,
          margins: formatMargins,
          reacquire: (lengthPt) => cellBlocks(
            orientRotatedCellBlocks(acquireAt(lengthPt), verticalMode),
          ),
        });
      } else {
        acquired = acquireAt(undefined);
      }
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
        ...(cell.borders.tl2br || cell.borders.tr2bl ? {
          diagonalBorders: Object.freeze({
            tl2br: retainedBorder(cell.borders.tl2br ?? null),
            tr2bl: retainedBorder(cell.borders.tr2bl ?? null),
          }),
        } : {}),
        ...(verticalText ? { verticalText } : {}),
        blocks: cellBlocks(acquired),
      };
      function cellBlocks(layouts: typeof acquired) {
        return layouts.flatMap((layout, sourceBlockIndex) => {
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
            ...(sourceElement?.type === 'paragraph' ? {
              keepLines: sourceElement.keepLines === true,
              // §17.3.1.44: omission enables widow/orphan control.
              widowControl: sourceElement.widowControl !== false,
            } : {}),
            ...((layout.kind === 'paragraph' && paragraphHasPageDependency(layout))
              ? { pageDependent: true }
              : {}),
            ...(isStructuralTrailingParagraph(
              cell.content,
              sourceBlockIndex,
              cell.hideMark === true,
            )
              ? { structuralTrailing: true }
              : {}),
          }];
        });
      }
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
  const placement = {
    container: { id: flowDomainId, kind: 'tableCell' as const, bounds },
    cursor: { xPt: 0, yPt: 0 },
    availableBounds: bounds,
  };
  let finalInput = input;
  let layout = layoutTable(input, placement, services).layout;
  // A rotated cell's lines run along its final height. When other cells or a
  // merge make that height larger than the acquired line length, acquire the
  // lines again at the final length so paragraph alignment uses it; the row
  // requirement stays the natural length, so row heights cannot change.
  const grown = rotatedCells.flatMap((rotated) => {
    const laidOut = layout.rows[rotated.rowIndex]?.cells[rotated.cellIndex];
    const current = input.rows[rotated.rowIndex]?.cells[rotated.cellIndex]?.verticalText;
    if (!laidOut || !current) return [];
    const finalLengthPt = Math.max(
      0,
      laidOut.flowBounds.heightPt - rotated.margins.top - rotated.margins.bottom,
    );
    return finalLengthPt > current.lineLengthPt + ROTATED_LINE_LENGTH_EPSILON_PT
      ? [{ rotated, finalLengthPt, current }]
      : [];
  });
  if (grown.length > 0) {
    const replacements = new Map(grown.map(({ rotated, finalLengthPt, current }) => [
      `${rotated.rowIndex}:${rotated.cellIndex}`,
      {
        blocks: rotated.reacquire(finalLengthPt),
        verticalText: { ...current, lineLengthPt: finalLengthPt },
      },
    ]));
    finalInput = snapshotPlainData<TableLayoutInput>({
      ...input,
      rows: input.rows.map((row, rowIndex) => ({
        ...row,
        cells: row.cells.map((cell, cellIndex) => {
          const replacement = replacements.get(`${rowIndex}:${cellIndex}`);
          return replacement ? { ...cell, ...replacement } : cell;
        }),
      })),
    }, 'RetainedTableAcquisition.input') as TableLayoutInput;
    layout = layoutTable(finalInput, placement, services).layout;
  }
  return Object.freeze({
    input: finalInput,
    layout,
    nestedById: Object.freeze(nestedById),
    floatingTables: snapshotPlainData(
      floatingTables,
      'RetainedTableAcquisition.floatingTables',
    ) as readonly NestedFloatingTableOccurrence[],
  });
}

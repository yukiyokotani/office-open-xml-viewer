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
import { layoutTable, measureTableCellBlockFlowHeightPt, tableRowRuleInsetsPt } from './table.js';
import { tableCellHorizontalSpacingInsets } from './table-columns.js';
import { unmeasuredTableMemberDecision, type LogicalTableDecision, type TableMemberDecision } from './table-layout-decision.js';
import { snapshotPlainData } from './plain-data.js';
import {
  leadingCellOwnerCarrier,
  tableRowsElectCarriers,
  CELL_OWNED_TABLE,
  type TableOwnerContext,
} from './table-owner-runs.js';
import { eastAsianUprightPaintOps } from './vertical-glyph-orientation.js';
import type {
  FloatingTablePositionInput,
  DrawingMLCollisionEntryPt,
  LayoutServices,
  LayoutNodeId,
  PaintNode,
  ParagraphLayout,
  TableBorderInput,
  TableCellBlockInput,
  TableCellVerticalMode,
  TableEdgeInputs,
  TableFormatInput,
  TableLayout,
  TableLayoutInput,
  TableRowLayoutInput,
  SourceRef,
  WrapExclusion,
} from './types.js';

export interface RetainedTableAcquisitionDependencies<State> {
  tableDecision?(table: TableLayoutSource, contentWidthPt: number, state: State): LogicalTableDecision;
  layoutServices(state: State): LayoutServices | undefined;
  tableFormat(table: TableLayoutSource): TableFormatInput;
  resolveColumns(table: TableLayoutSource, contentWidthPt: number, state: State, decision?: LogicalTableDecision): readonly number[];
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

const pagePlacedContentTables = new WeakMap<RetainedTableAcquisition, boolean>();

/**
 * Whether this table, laid out whole at a known page origin, has anything to
 * place on the page: a §17.4.57 positioned table of its own or below an
 * in-flow nested table at any depth, a positioned table's own positioned
 * children, or a cell paragraph's text box whose story holds such content.
 * Only a paginated table's direct positioned children are resolved by its
 * own fragment (table-pagination.ts finalFrameRow); every deeper one is
 * placed through the page origin of the table that holds it
 * (table-pagination.ts layoutNestedWhole).
 */
export function hasPagePlacedContent(source: RetainedTableAcquisition): boolean {
  const known = pagePlacedContentTables.get(source);
  if (known !== undefined) return known;
  const result = source.floatingTables.length > 0 || needsNestedPageOrigins(source);
  pagePlacedContentTables.set(source, result);
  return result;
}

/** Whether placing this table needs the page origin of a block inside it:
 * an in-flow nested table, a positioned child or a text box with page-placed
 * content. */
export function needsNestedPageOrigins(source: RetainedTableAcquisition): boolean {
  return source.input.rows.some((row) => rowNeedsPageOrigins(source, row))
    || source.floatingTables.some((occurrence) => {
      const positioned = source.nestedById[occurrence.tableId];
      return positioned !== undefined && hasPagePlacedContent(positioned);
    });
}

/** A cell paragraph whose text boxes hold page-placed content: their stories'
 * page frames need the paragraph's page position (paragraph.ts
 * hostFlowPageTranslationPt), which only the table's pagination knows in a
 * cell. */
export function paragraphHasBandDependentTextBoxes(layout: ParagraphLayout): boolean {
  return layout.textBoxes.some((textBox) => textBox.story.bandDependent === true);
}

/** A cell block whose placement needs its page origin. */
export function blockNeedsPageOrigin(
  source: RetainedTableAcquisition,
  block: TableCellBlockInput,
): boolean {
  if (block.layout.kind === 'paragraph') return paragraphHasBandDependentTextBoxes(block.layout);
  if (block.layout.kind !== 'table') return false;
  const nested = source.nestedById[block.layout.id];
  return nested !== undefined && hasPagePlacedContent(nested);
}

export function rowNeedsPageOrigins(
  source: RetainedTableAcquisition,
  row: TableRowLayoutInput,
): boolean {
  return row.cells.some((cell) => cell.blocks.some((block) => blockNeedsPageOrigin(source, block)));
}

function retainedNodeIsReusableAcrossPages(
  node: PaintNode,
  visited: Set<PaintNode>,
): boolean {
  if (visited.has(node)) return true;
  visited.add(node);
  if (node.kind === 'drawing') return node.anchorLayer === undefined;
  if (node.kind === 'paragraph') {
    return node.lines.every((line) => !line.placements.some(placementDependsOnPage))
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
  return layout.lines.some((line) => line.placements.some(placementDependsOnPage));
}

/** A PAGE field result is its own placement, or (same formatting as the
 * text before it) one owner range of a shared glyph sequence. */
function placementDependsOnPage(placement: ParagraphLayout['lines'][number]['placements'][number]): boolean {
  return placement.kind === 'text' && (
    placement.dependency === 'page'
    || placement.sourceRuns?.some((owner) => owner.dependency === 'page') === true
  );
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

type TableCellLayoutInput = TableLayoutInput['rows'][number]['cells'][number];

/** Each retained table's rotated cell acquisitions by cell id, so a
 * page-local fragment or owner segment of the table can acquire their lines
 * again at its own length (layoutTableFragment). Keyed by the acquisition's
 * retained input; a projection with its own input (an owner segment) shares
 * them through shareRotatedCellReacquisition. Cell ids survive the
 * per-occurrence cell copies pagination makes. Same-realm only: the closures
 * re-enter this acquisition's paragraph services. */
const rotatedCellReacquisitions = new WeakMap<TableLayoutInput, ReadonlyMap<string, RotatedCellAcquisition>>();

/** Let a projection of retained tables (their rows re-sliced, or an adjacent
 * group's members combined, into a new input) re-acquire their rotated cells. */
export function shareRotatedCellReacquisition(
  from: TableLayoutInput | readonly TableLayoutInput[],
  to: TableLayoutInput,
): void {
  const sources = (Array.isArray(from) ? from : [from]) as readonly TableLayoutInput[];
  const shared = new Map<string, RotatedCellAcquisition>();
  for (const source of sources) {
    rotatedCellReacquisitions.get(source)?.forEach((rotated, cellId) => shared.set(cellId, rotated));
  }
  if (shared.size > 0) rotatedCellReacquisitions.set(to, shared);
}

/**
 * A rotated cell's lines run along its final height. When other cells or a
 * merge make that height larger than the acquired line length, acquire the
 * lines again at the final length so paragraph alignment uses it; the row
 * requirement stays the natural length, so row heights cannot change. The
 * final length excludes the cell's collapsed rules
 * (WORD_COLLAPSED_BORDER_ROW_TRACK_FOOTPRINT), which only an exact row's
 * authored line length did not already leave out: that fixed row is
 * re-acquired shorter as well. Those rules depend on the rows laid out
 * together, so a page-local fragment resolves them again.
 */
function relayoutRotatedCellLines(
  input: TableLayoutInput,
  laidOut: ReturnType<typeof layoutTable>,
  placement: Parameters<typeof layoutTable>[1],
  services: Parameters<typeof layoutTable>[2],
  rotatedOf: (cell: TableCellLayoutInput) => RotatedCellAcquisition | undefined,
  label: string,
): { input: TableLayoutInput; laidOut: ReturnType<typeof layoutTable> } {
  const { layout } = laidOut;
  let ruleInsets: ReturnType<typeof tableRowRuleInsetsPt> | null = null;
  const tableBottomPt = layout.flowBounds.yPt + layout.advancePt;
  const replacements = new Map<string, Pick<TableCellLayoutInput, 'blocks' | 'verticalText'>>();
  input.rows.forEach((row, rowIndex) => row.cells.forEach((cell, cellIndex) => {
    const rotated = rotatedOf(cell);
    const laidOut = layout.rows[rowIndex]?.cells[cellIndex];
    const current = cell.verticalText;
    if (!rotated || !laidOut || !current) return;
    ruleInsets ??= tableRowRuleInsetsPt(input);
    const reachesTableBottom =
      laidOut.flowBounds.yPt + laidOut.flowBounds.heightPt >= tableBottomPt - ROTATED_LINE_LENGTH_EPSILON_PT;
    const finalLengthPt = Math.max(
      0,
      laidOut.flowBounds.heightPt - rotated.margins.top - rotated.margins.bottom
        - (ruleInsets.top[rowIndex] ?? 0)
        - (reachesTableBottom ? ruleInsets.bottom[input.rows.length - 1] ?? 0 : 0),
    );
    const exact = row.heightRule === 'exact';
    if (
      finalLengthPt > current.lineLengthPt + ROTATED_LINE_LENGTH_EPSILON_PT
      || (exact && finalLengthPt < current.lineLengthPt - ROTATED_LINE_LENGTH_EPSILON_PT)
    ) {
      replacements.set(`${rowIndex}:${cellIndex}`, {
        blocks: rotated.reacquire(finalLengthPt),
        verticalText: { ...current, lineLengthPt: finalLengthPt },
      });
    }
  }));
  if (replacements.size === 0) return { input, laidOut };
  const finalInput = snapshotPlainData<TableLayoutInput>({
    ...input,
    rows: input.rows.map((row, rowIndex) => ({
      ...row,
      cells: row.cells.map((cell, cellIndex) => {
        const replacement = replacements.get(`${rowIndex}:${cellIndex}`);
        return replacement ? { ...cell, ...replacement } : cell;
      }),
    })),
  }, label) as TableLayoutInput;
  return { input: finalInput, laidOut: layoutTable(finalInput, placement, services) };
}

/**
 * Lay out a page-local fragment of a retained table. A fragment's first and
 * last rows take the table's outer rules, so an exact row's rotated cell can
 * have a different line length than it had in the whole table; its lines are
 * acquired again at the fragment's length.
 */
export function layoutTableFragment(
  source: RetainedTableAcquisition,
  input: TableLayoutInput,
  placement: Parameters<typeof layoutTable>[1],
  services: Parameters<typeof layoutTable>[2],
): ReturnType<typeof layoutTable> {
  const laidOut = layoutTable(input, placement, services);
  const rotated = rotatedCellReacquisitions.get(source.input);
  if (!rotated) return laidOut;
  return relayoutRotatedCellLines(
    input,
    laidOut,
    placement,
    services,
    (cell) => rotated.get(cell.id),
    'TableFragment.input',
  ).laidOut;
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
  decision?: TableMemberDecision,
  owner: TableOwnerContext = CELL_OWNED_TABLE,
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
  const member = decision ?? dependencies.tableDecision?.(table, contentWidthPt, outerState).logical
    ?? unmeasuredTableMemberDecision(table, dependencies.tableFormat(table));
  const format = member.source.format;
  // Cell-owner rows (table-owner-runs.ts) elect only in a promoting owner
  // context (tableRowsElectCarriers, WORD_CELL_OWNER_ROW_CONTEXT: body,
  // header and footer roots) and only when the table's effective §17.4.57
  // positioning is null — lexical tblpPr presence does not decide
  // (`word-effective-floating-table-positioning`). The paginated body, its
  // vertical upright placement (body-table-measurement.ts) and header/footer
  // story roots (production-body-layout.ts) consume the run projection.
  // Nested tables (the default owner) and text box and note roots elect
  // nothing: their framePr stays an ordinary cell-paragraph fact.
  const electsOwnerRuns = format.positioning === null && tableRowsElectCarriers(owner);
  const bidiVisual = table.bidiVisual === true;
  const tableIndentPt = member.tableIndentPt;
  // ast-grep-ignore: no-plain-object-string-map -- keys are library-generated table flow ids.
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
              const nestedDecision = dependencies.tableDecision?.(nestedTable, nestedContentWidthPt, cellState);
              const nestedColumns = dependencies.resolveColumns(
                nestedTable,
                nestedContentWidthPt,
                cellState,
                nestedDecision,
              );
              const nested = acquireRetainedTable(
                nestedTable,
                nestedColumns,
                nestedContentWidthPt,
                cellState,
                sourceAt(nestedPath),
                dependencies,
                nestedDecision?.logical,
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
        // ECMA-376 §17.4.80 (trHeight): only exact and atLeast give @val a
        // meaning; an explicit auto row ignores it and sizes to its content
        // (as the row track does, see semanticRowFloor in table.ts), so it is
        // no floor on the rotated line length either. The rule arrives
        // normalized; see WORD_OMITTED_ROW_HEIGHT_RULE_AT_LEAST.
        const rowHeightPt = rowRule === 'auto' ? 0 : rowFormat?.height?.valuePt ?? 0;
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
    const ownerCarrier = electsOwnerRuns
      ? leadingCellOwnerCarrier(row, sourceAt([...sourcePath, rowIndex]))
      : null;
    return {
      ...(ownerCarrier ? { ownerCarrier } : {}),
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
      // Carry one signed leading-axis translation into the sole table layout
      // algorithm. End alignment uses each row's first margin; leading alignment
      // keeps the first-row anchor even when subsequent margins differ.
      indentPt: member.rowTranslationsPt[rowIndex] ?? tableIndentPt,
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
  const rotatedByCell = new Map(rotatedCells.map((rotated) => [
    input.rows[rotated.rowIndex]!.cells[rotated.cellIndex]!,
    rotated,
  ]));
  const { input: finalInput, laidOut: { layout } } = relayoutRotatedCellLines(
    input,
    layoutTable(input, placement, services),
    placement,
    services,
    (cell) => rotatedByCell.get(cell),
    'RetainedTableAcquisition.input',
  );
  if (rotatedCells.length > 0) {
    rotatedCellReacquisitions.set(finalInput, new Map(rotatedCells.map((rotated) => [
      input.rows[rotated.rowIndex]!.cells[rotated.cellIndex]!.id,
      rotated,
    ])));
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

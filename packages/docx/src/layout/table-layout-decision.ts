import type { TableFormatInput } from './types.js';
import { projectTableColumnLayoutInput, projectEffectiveTablePreferredWidthPt, absoluteTablePreferredWidthPt,
  type TableLayoutSource, type TableSourceAcquisitionInput } from './table-source-acquisition.js';
import { wordFixedOccupiedGridInput, wordMeasuredTableOriginMode, wordMeasuredTableOriginPlacementMode,
  wordTableEffectiveIndentPt, wordTableOriginTranslationPt } from './table-compatibility.js';
import { isVerticalTextDirection } from './section-orientation.js';

export interface TableDecisionEnvironment {
  readonly mode: number | undefined;
  readonly story: string | undefined;
  readonly topLevel: boolean;
  readonly singleColumn: boolean;
  readonly textDirection: string | null | undefined;
  readonly inTableCell: boolean;
  readonly contentX: number;
  readonly pageWidth: number;
}

/** WORD_TABLE_ORIGIN_COMPATIBILITY facts before source rows are acquired.
 * §17.4.37 groups share the first logical row's leading anchor and must test
 * the measured scope over every member, rather than restarting at each tbl. */
function tableOriginContext(
  table: TableLayoutSource,
  format: TableFormatInput,
  mode: number | undefined,
) {
  const firstRowException = format.firstRowException;
  const tableIndentPt = firstRowException?.indentAuthored
    ? (firstRowException.indentPt ?? 0)
    : (table.tblInd ?? 0);
  // WORD_TABLE_ORIGIN_COMPATIBILITY: preserve the prior contract for unresolved
  // margins/spacing, nested origins, positioned tables, and unmeasured modes.
  const measuredStructure = format.ordinaryFlow
    && table.widthPct == null
    && (!firstRowException?.preferredWidthAuthored || firstRowException.preferredWidth?.kind !== 'pct')
    && (firstRowException?.layout === 'fixed' || table.layout === 'fixed'
      || (firstRowException?.preferredWidthAuthored
        ? firstRowException.preferredWidth?.kind === 'dxa' && firstRowException.preferredWidth.value > 0
        : table.widthPt != null && table.widthPt > 0))
    && table.rows.every((row) => row.cells.every((cell) => cell.widthPt != null && cell.widthPt > 0 && cell.widthPct == null))
    && format.rows.every((row) => row.cellSpacingPt === 0);
  const measuredOrigin = wordMeasuredTableOriginPlacementMode(mode) && measuredStructure
    && format.rows.every((row) => row.cells[0]?.originLeftMarginPt != null);
  // WORD_TABLE_ORIGIN_COMPATIBILITY: the same structure with horizontal
  // margins omitted at every layer was measured only for end justification.
  const omittedMarginEndOrigin = wordMeasuredTableOriginMode(mode) && measuredStructure
    && format.rows.every((row) => row.cells.every((cell) => cell.horizontalMarginsOmitted === true));
  const indentAuthored = firstRowException?.indentAuthored || table.tblInd != null;
  // Provenance travels with the selected logical indent: a first-row tblPrEx
  // indent is authored, the table's own value may be style-inherited.
  const tableIndentInherited = !firstRowException?.indentAuthored && table.tblIndInherited === true;
  const firstLeftMarginPt = format.rows[0]?.cells[0]?.originLeftMarginPt ?? 0;
  const effectiveLayout: 'fixed' | 'autofit' = firstRowException?.layout === 'fixed' || table.layout === 'fixed' ? 'fixed' : 'autofit';
  return { measuredOrigin, omittedMarginEndOrigin, tableIndentPt, tableIndentInherited, indentAuthored,
    firstLeftMarginPt, effectiveLayout };
}


/** All compatibility selection lives at this boundary. A member receives facts
 * selected here; it must never re-evaluate the logical table's eligibility. */
function memberContext(
  table: TableLayoutSource,
  source: TableSourceAcquisitionInput,
  origin: ReturnType<typeof tableOriginContext>,
  measuredOrigin: boolean,
  contentWPt: number,
  state: TableDecisionEnvironment,
  logicalProperties = measuredOrigin,
  omittedMarginEndOrigin = false,
) {
  const format = source.format;
  const firstRowException = format.firstRowException;
  // MS-OI29500 §2.1.156/158: first-logical-row indent/layout authority
  // also applies when only the fixed-grid observation is enabled. The origin
  // observation controls margin hanging, not the scope of those properties.
  const tableIndentPt = logicalProperties ? origin.tableIndentPt
    : firstRowException?.indentAuthored ? (firstRowException.indentPt ?? 0) : (table.tblInd ?? 0);
  const baseIndentPt = logicalProperties ? origin.tableIndentPt
    : Number.isFinite(table.tblInd) ? (table.tblInd ?? 0) : 0;
  // ECMA-376 §17.4.50 judges tblInd by each row's resulting jc. A tblInd
  // inherited from the table style is dropped for a non-leading row in every
  // class; a direct (or tblPrEx) indent keeps its per-class contract below.
  // Each indent is used with the provenance of the value actually selected.
  const tableIndentInherited = logicalProperties ? origin.tableIndentInherited
    : !firstRowException?.indentAuthored && table.tblIndInherited === true;
  const baseIndentInherited = logicalProperties ? origin.tableIndentInherited
    : table.tblIndInherited === true;
  const leading = (justification: string | null | undefined) =>
    justification == null || justification === 'left' || justification === 'start' || justification === 'both';
  const rowIndent = (justification: string | null | undefined, indentPt: number, inherited: boolean) =>
    inherited && !leading(justification) ? 0 : indentPt;
  // A row is inside an origin observation through the explicit-margin class
  // or, for end justification only, the omitted-margin class.
  const measuredRow = (justification: string | null | undefined) => measuredOrigin
    || (omittedMarginEndOrigin && justification === 'end');
  const widthIndent = (justification: string | null | undefined, indentPt: number, inherited = baseIndentInherited) =>
    wordTableEffectiveIndentPt({ measured: measuredRow(justification), justification,
      indentPt: rowIndent(justification, indentPt, inherited) });
  const rowIndentPts = format.rows.map((row) => {
    const ownException = !logicalProperties && row.exception?.indentAuthored === true;
    return widthIndent(row.justification ?? table.jc,
      ownException ? (row.exception?.indentPt ?? 0) : baseIndentPt,
      !ownException && baseIndentInherited);
  });
  const rowTranslationsPt = format.rows.map((row) => wordTableOriginTranslationPt({
    mode: state.mode, measured: measuredRow(row.justification ?? table.jc), justification: row.justification ?? table.jc,
    indentPt: rowIndent(row.justification ?? table.jc, tableIndentPt, tableIndentInherited),
    indentAuthored: logicalProperties ? origin.indentAuthored : firstRowException?.indentAuthored || table.tblInd != null,
    firstLeftMarginPt: origin.firstLeftMarginPt,
    rowLeftMarginPt: row.cells[0]?.originLeftMarginPt ?? 0,
  }));
  const effectiveLayout: 'fixed' | 'autofit' = logicalProperties ? origin.effectiveLayout
    : firstRowException?.layout === 'fixed' || table.layout === 'fixed' ? 'fixed' : 'autofit';
  // §17.18.87 lets AutoFit override its preferred width up to the page width.
  // For a table with a preferred tblW, keep the ordinary text-band ceiling
  // unless §17.4.50 placement moves a top-level page-owned story table into
  // its semantic leading margin (auto-width tables use the leading-indent band
  // below instead). The promoted ceiling is the physical page distance left
  // after resolving jc + tblInd, rather than the page width in isolation: a
  // partial negative indent does not move the table origin all the way to the
  // page edge. Test the authored indent before bidiVisual reverses its physical translation;
  // table justification reverses under bidiVisual as well. Body tables must
  // also remain in a single-column page band; headers and footers are
  // page-owned stories and do not inherit the body's newspaper columns.
  const story = state.story;
  const isTopLevelPageOwnedStory = state.topLevel
    && (story === 'header'
      || story === 'footer'
      || (story === 'body' && state.singleColumn));
  const isLeadingMarginPageStoryTable = format.ordinaryFlow
    && isTopLevelPageOwnedStory
    && !isVerticalTextDirection(state.textDirection)
    // §17.4.50 is based on the resulting row jc. A leading table default
    // cannot promote the ceiling when every measured row is nonleading.
    // Keep the historical default-indent check for unmeasured classes.
    && (measuredOrigin && format.rows.length > 0
      ? rowIndentPts : [widthIndent(table.jc, baseIndentPt), ...rowIndentPts])
      .some((indentPt) => indentPt < 0);
  const rowPlacements = format.rows.length === 0
    ? [{ justification: table.jc, indentPt: widthIndent(table.jc, baseIndentPt) }]
    : format.rows.map((row, rowIndex) => ({
        justification: row.justification ?? table.jc,
        indentPt: rowIndentPts[rowIndex] ?? baseIndentPt,
      }));
  const bidiVisual = table.bidiVisual === true;
  const pageFitCeilingPt = Math.min(
    state.pageWidth,
    ...rowPlacements.map(({ justification, indentPt }) => {
      const trailing = justification === 'right' || justification === 'end';
      const alignment = justification === 'center'
        ? 'center'
        : (bidiVisual ? !trailing : trailing) ? 'right' : 'left';
      const signedIndentPt = bidiVisual ? -indentPt : indentPt;
      if (alignment === 'left') {
        const resolvedOriginPt = state.contentX + signedIndentPt;
        return state.pageWidth - resolvedOriginPt;
      }
      if (alignment === 'right') {
        const resolvedTrailingEdgePt = state.contentX + contentWPt + signedIndentPt;
        return resolvedTrailingEdgePt;
      }
      const resolvedCenterPt = state.contentX + contentWPt / 2 + signedIndentPt;
      return 2 * Math.min(resolvedCenterPt, state.pageWidth - resolvedCenterPt);
    }),
  );

  // WORD_AUTOFIT_LEADING_INDENT_BAND (table-compatibility.ts): for an
  // AutoFit table with no preferred tblW (§17.4.63 auto), a leading §17.4.50
  // tblInd moves only the leading edge while the trailing edge stays at the
  // text band. The text band available to the grid is contentW - tblInd for
  // either sign, and the physical page is not a ceiling. In mode 14 (or with
  // the mode omitted) outer cell margins hang outside that band, so the table
  // adds them in full; mode 15 removes that allowance. The additional
  // center/right grid-matched controls use that same mode-14 fitted width and
  // the whole mode-15 band. Their +5.4pt-indent geometry does not settle other
  // grid/indent combinations or absolute origin policy (see the rule's limits).
  // WORD_FIRST_ROW_TABLE_EXCEPTION_SCOPE makes first-row tblPrEx/tblW authoritative for
  // the whole table. Use the solver's resolver, including auto clearing dxa.
  const hasPreferredTableWidth = projectEffectiveTablePreferredWidthPt(source, contentWPt) !== null;
  const savedGridWidthPt = (table.colWidths ?? []).reduce(
    (sum, width) => sum + (Number.isFinite(width) ? Math.max(0, width) : 0), 0,
  );
  // Nonleading controls only establish the case where saved grid + indent
  // equals the band. There the two possible mode-14 ceilings coincide; keep
  // other nonleading geometries on their established width contract rather
  // than choosing between indistinguishable hypotheses. The epsilon covers
  // point arithmetic, not a visual fit tolerance.
  const hasMeasuredAlignmentGeometry = rowPlacements.every(({ justification, indentPt }) => {
    const nonleading = justification === 'center' || justification === 'right' || justification === 'end';
    return !nonleading || Math.abs(savedGridWidthPt + indentPt - contentWPt) <= 1e-9;
  });
  const compatibilityMode = state.mode;
  const hasMeasuredCompatibilityMode = compatibilityMode === undefined
    || compatibilityMode === 14 || compatibilityMode === 15;
  const usesLeadingIndentBand = effectiveLayout !== 'fixed'
    && format.ordinaryFlow
    && isTopLevelPageOwnedStory
    && !isVerticalTextDirection(state.textDirection)
    && !hasPreferredTableWidth
    && hasMeasuredCompatibilityMode
    && hasMeasuredAlignmentGeometry;
  const outerCellMarginsHangOutsideBand = compatibilityMode === undefined || compatibilityMode === 14;
  const textBandPt = usesLeadingIndentBand
    ? Math.max(0, Math.min(...rowPlacements.map(({ justification, indentPt }) => {
        const trailing = justification === 'right' || justification === 'end';
        const leading = justification !== 'center' && !trailing;
        // The center/right controls retain the leading band's width in mode
        // 14, but use the full text band in mode 15. This is a width rule;
        // origin translation is acquired separately under WORD_TABLE_ORIGIN_COMPATIBILITY.
        return leading || outerCellMarginsHangOutsideBand ? contentWPt - indentPt : contentWPt;
      })))
    : contentWPt;
  // WORD_AUTOFIT_OUTER_CELL_MARGIN_BAND (table-compatibility.ts): an AutoFit
  // grid can include outer §17.4.42 cell margins beyond the text band. For a
  // preferred-width table, only the margin overhang already represented by
  // §17.4.48 tblGrid is used (that class is not covered by the mode controls).
  // A row with skipped outer tracks cannot establish that margin ownership.
  // A nested table's saved overhang belongs to its containing cell; giving it
  // the page-table allowance enlarges that cell's contents beyond Word's grid.
  // Windows Word 16.0 #1881 control T9: an AutoFit table in a text box fits
  // its border box to the box's inner width; no outer margin overhangs.
  const possibleOuterCellMarginsPt = effectiveLayout === 'fixed'
    || !format.ordinaryFlow
    || state.story === 'textbox'
    || (isLeadingMarginPageStoryTable && !usesLeadingIndentBand)
    || isVerticalTextDirection(state.textDirection)
    || state.inTableCell
    || format.rows.length === 0
    ? 0
    : format.rows.reduce((minimumPt, row, rowIndex) => {
        const sourceRow = table.rows[rowIndex];
        if (!sourceRow || (sourceRow.gridBefore ?? 0) > 0 || (sourceRow.gridAfter ?? 0) > 0) return 0;
        const first = row.cells[0]?.marginsPt;
        const last = row.cells.at(-1)?.marginsPt;
        const leftPt = first?.left;
        const rightPt = last?.right;
        const marginPt = typeof leftPt === 'number' && Number.isFinite(leftPt)
          && typeof rightPt === 'number' && Number.isFinite(rightPt)
          ? Math.max(0, leftPt) + Math.max(0, rightPt)
          : 0;
        return Math.min(minimumPt, marginPt);
      }, Number.POSITIVE_INFINITY);
  // compatSetting compatibilityMode: WORD_AUTOFIT_LEADING_INDENT_BAND treats
  // an omitted setting like an explicit 14.
  // The two-cell forced-fit distribution always shares the actual outer
  // margins; only the ceiling differs by compatibility mode.
  const forcedFitOuterMarginsPt = usesLeadingIndentBand
    ? possibleOuterCellMarginsPt
    : Math.min(possibleOuterCellMarginsPt, Math.max(0, savedGridWidthPt - contentWPt));
  const outerCellMarginsPt = usesLeadingIndentBand
    ? (outerCellMarginsHangOutsideBand ? possibleOuterCellMarginsPt : 0)
    : Math.min(possibleOuterCellMarginsPt, Math.max(0, savedGridWidthPt - contentWPt));
  // A preferred-width table with a negative indent keeps the older physical
  // page ceiling (its authored width may reach the page edge); the auto-width
  // band above supersedes it for tables without a preferred width.
  const containerMaximumPt = (isLeadingMarginPageStoryTable && !usesLeadingIndentBand
    ? Math.max(contentWPt, pageFitCeilingPt)
    : textBandPt) + outerCellMarginsPt;
  // Windows Word 16.0 #1881 controls: an AutoFit table whose effective tblW is
  // an absolute length keeps its saved grid up to that width beyond its
  // container (body column, past the page edge, two-column section column);
  // a percentage or auto width fits the container. tblW never
  // widens the grid here (the WORD_TABLE_ORIGIN_COMPATIBILITY center/end
  // controls keep a 468pt grid under a 540pt tblW).
  // A nested AutoFit table keeps the established cell fit: the #1881 nested
  // control (grid kept past a 200pt cell) and a Word-exported resume whose
  // nested bars stay inside their cell disagree, and no measured input
  // separates them yet.
  const absolutePreferredPt = effectiveLayout !== 'fixed' && !state.inTableCell
    ? absoluteTablePreferredWidthPt(source) : null;
  const maximumTableWidthPt = absolutePreferredPt === null ? containerMaximumPt
    : Math.max(containerMaximumPt, Math.min(absolutePreferredPt, savedGridWidthPt));
  const isFixedNestedTable = effectiveLayout === 'fixed'
    && state.inTableCell;

  return Object.freeze({ source, tableIndentPt, effectiveLayout,
    rowTranslationsPt: Object.freeze(rowTranslationsPt), maximumTableWidthPt,
    isFixedNestedTable, usesLeadingIndentBand, forcedFitOuterMarginsPt });
}

export type TableMemberDecision = ReturnType<typeof memberContext>;
export interface LogicalTableDecision {
  readonly measuredOrigin: boolean;
  readonly usesLogicalProperties: boolean;
  readonly dropsUnusedLeadingGrid: boolean;
  readonly logical: TableMemberDecision;
  readonly members: readonly TableMemberDecision[];
}

/** ECMA-376 §17.4.37 concatenates same-style adjacent table rows. Evaluate
 * WORD_TABLE_ORIGIN_COMPATIBILITY and WORD_FIXED_UNUSED_LEADING_GRID once
 * over that complete sequence. MS-OI29500 §§2.1.156/158/167 give its first row
 * authority over width/layout/indent exceptions. Differing authored frames or
 * grids retain the established union policy. Disabled observations apply to
 * every member, including a member that would qualify by itself. Disabled
 * groups preserve the pre-existing member width/origin contract, including its
 * split/single inconsistency; this rule does not repair unmeasured classes.
 * Construction is linear in the group's rows/cells/tracks, with no measurement. */
export function decideLogicalTable(
  table: TableLayoutSource,
  source: TableSourceAcquisitionInput,
  members: readonly Readonly<{ table: TableLayoutSource; source: TableSourceAcquisitionInput }>[],
  contentWidthPt: number,
  authoredEnvironment: TableDecisionEnvironment,
  commonFrame = true,
  commonGrid = true,
): LogicalTableDecision {
  // A producer-supplied outer-edge table (outerEdgeIndent) carries its own
  // physical outer edge and grid; the OOXML table compatibility modes
  // describe tblInd/tblGrid semantics it does not have, so its decisions
  // keep the mode-absent contract.
  const environment = source.format.outerEdgeIndent === true && authoredEnvironment.mode !== undefined
    ? { ...authoredEnvironment, mode: undefined }
    : authoredEnvironment;
  const origin = tableOriginContext(table, source.format, environment.mode);
  const topLevelBody = environment.topLevel && environment.story === 'body';
  // An omitted mode is inside the origin observation; the outer-edge table
  // keeps its unmeasured placement instead.
  const measuredOrigin = commonFrame && commonGrid && topLevelBody && origin.measuredOrigin
    && source.format.outerEdgeIndent !== true;
  const omittedMarginEndOrigin = commonFrame && commonGrid && topLevelBody && origin.omittedMarginEndOrigin
    && source.format.outerEdgeIndent !== true;
  const input = projectTableColumnLayoutInput(source, contentWidthPt,
    () => ({ minWidthPt: 0, maxWidthPt: 0 }), contentWidthPt);
  const gridScope = commonFrame && commonGrid && topLevelBody && source.format.ordinaryFlow
    && table.widthPct == null && source.format.firstRowException?.preferredWidth?.kind !== 'pct'
    && source.format.rows.every((row) => row.cellSpacingPt === 0);
  const dropsUnusedLeadingGrid = wordFixedOccupiedGridInput(input, environment.mode, gridScope) !== input;
  const usesLogicalProperties = measuredOrigin || dropsUnusedLeadingGrid;
  const logical = memberContext(table, source, origin, measuredOrigin, contentWidthPt, environment,
    usesLogicalProperties, omittedMarginEndOrigin);
  return Object.freeze({ measuredOrigin, usesLogicalProperties, dropsUnusedLeadingGrid, logical,
    members: Object.freeze(members.map((member) => memberContext(member.table, member.source,
      origin, measuredOrigin, contentWidthPt, environment, usesLogicalProperties, omittedMarginEndOrigin))),
  });
}

/** Public hand-built acquisition inputs have no parser lexical facts. Keep their
 * established unmeasured placement; production supplies the logical decision. */
export function unmeasuredTableMemberDecision(table: TableLayoutSource, format: TableFormatInput): TableMemberDecision {
  const source: TableSourceAcquisitionInput = {
    semantic: { colWidths: table.colWidths ?? [], layout: table.layout ?? null, widthPt: table.widthPt ?? null, widthPct: table.widthPct ?? null,
      rows: table.rows.map((row) => ({ gridBefore: row.gridBefore ?? 0, gridAfter: row.gridAfter ?? 0,
        cells: row.cells.map((cell) => ({ colSpan: cell.colSpan, widthPt: cell.widthPt, widthPct: cell.widthPct ?? null })) })) },
    lexical: { table: null, rows: table.rows.map((row) => ({ row: null, cells: row.cells.map(() => null) })) },
    format,
  };
  const origin: ReturnType<typeof tableOriginContext> = { measuredOrigin: false, omittedMarginEndOrigin: false,
    effectiveLayout: format.firstRowException?.layout === 'fixed' || table.layout === 'fixed' ? 'fixed' : 'autofit',
    tableIndentPt: format.firstRowException?.indentAuthored ? (format.firstRowException.indentPt ?? 0) : (table.tblInd ?? 0),
    tableIndentInherited: !format.firstRowException?.indentAuthored && table.tblIndInherited === true,
    indentAuthored: format.firstRowException?.indentAuthored || table.tblInd != null,
    firstLeftMarginPt: format.rows[0]?.cells[0]?.originLeftMarginPt ?? 0 };
  return memberContext(table, source, origin, false, 0, { mode: undefined, story: undefined, topLevel: false,
    singleColumn: false, textDirection: null, inTableCell: false, contentX: 0, pageWidth: 0 });
}

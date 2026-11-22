import type { FramePr } from '../types.js';
import { ownValue } from '@silurus/ooxml-core/internal/own-key';
import { defineCompatibilityRule } from './compatibility.js';
import type { DocxCompatProfile } from './compat-profile.js';
import { stableFingerprint } from './fingerprint.js';
import type { ParagraphLayoutSource, ParagraphAcquisitionInput } from './text.js';
import type { TableLayoutSource } from './table-source-acquisition.js';
import type { TableColumnLayoutInput, LayoutRect, SourceRef, FloatingTablePositionInput, TableCellLayoutInput, TableBorderInput } from './types.js';

export const WORD_CELL_OWNED_GRID_FRAME = defineCompatibilityRule({
  id: 'word-cell-owned-grid-frame',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'native-cell-grid-frame-position-controls',
    application: 'Microsoft Word',
    version: '16.113.3',
    platform: 'macOS 27.0',
  },
  description: 'Word controls (synthetic binary Word documents opened in Word) move six homogeneous fixed-layout horizontal LTR nested grids, including their borders and fills, together for left/center/right/absolute horizontal placement and nonzero text-relative vertical offsets. A producer supplies `cellFrame` only for a first host-cell block at depth two, with every physical cell paragraph agreeing with its first-cell frame, and no table-level positioning. Unpositioned anchors and no-wrap are admission counterexamples; mixed frames, later insertion, recursive child content, AutoFit, bidi and rotated cells retain residual diagnostics. The frame uses margin/text anchors, around wrap and automatic dimensions. The source format establishes row identity and frame coordinates, not a universal nested-grid placement or AUTO-width rule. In the observed equal host bands, different fixed grid widths have the same aligned displacement, consistent with host content width as the alignment frame. Retain fixed grid ink unchanged, including overhang. Reusing that band for exclusion is a bounded library policy: varied host widths and general wrapping are unproved. Original and vertical-offset controls additionally distinguish the initial, top-aligned empty carrier from a carrier containing authored whitespace: only the former keeps its paragraph-relative wrapNone drawing at its pre-grid-admission insertion reference. Table and text flow still clear the complete grid frame. Other references, collision-constrained drawings, continuations and nonempty carriers retain actual flow ownership. The following paragraph selects the page and wrap acquisition, but its before-spacing does not replace the cell insertion cursor as the vertical frame origin.',
});

export function cellOwnedGridFramePosition(frame: Readonly<FramePr>): FloatingTablePositionInput {
  return {
    textAnchor: 'cell-start', widthBasis: 'host-cell-content',
    leftFromTextPt: frame.hSpace, rightFromTextPt: frame.hSpace,
    topFromTextPt: frame.vSpace, bottomFromTextPt: frame.vSpace,
    horzAnchor: frame.hAnchor, horzSpecified: true, vertAnchor: frame.vAnchor,
    xPt: frame.x ?? 0, yPt: frame.y ?? 0,
    ...(frame.xAlign == null ? {} : { xAlign: frame.xAlign }),
    ...(frame.yAlign == null ? {} : { yAlign: frame.yAlign }),
  };
}

/** Bounded Word import observation: grid-wrap admission moves the table's
 * flow, but an empty, top-aligned carrier retains its paragraph insertion
 * reference. ECMA-376 §20.4.3.5 names the paragraph reference; it does not
 * prescribe this import behavior. Whitespace and other inline content own
 * real flow and are deliberate counterexamples. */
export function wordGridFramePreservesEmptyCarrierReference(cell: TableCellLayoutInput): boolean {
  const paragraph = cell.blocks[0]?.layout;
  if (cell.columnStart !== 0 || cell.vAlign !== 'top' || cell.verticalMerge !== 'none' || cell.verticalText
    || paragraph?.kind !== 'paragraph' || paragraph.clipBounds
    || paragraph.cellContainmentBounds || paragraph.lines.length !== 1
    || paragraph.resources.length !== 0 || paragraph.textBoxes.length !== 0
    || paragraph.events.length !== 0 || paragraph.exclusions.length !== 0
    || paragraph.drawings.length !== 1 || paragraph.anchorFrames?.length !== 1) return false;
  const frame = paragraph.anchorFrames[0]!;
  return frame.status === 'resolved' && frame.axes.vertical.relativeFrom === 'paragraph'
    && frame.geometry.wrap.kind === 'none'
    && paragraph.lines[0]!.placements.every((placement) => (
      (placement.kind === 'anchor-host' && placement.anchorOccurrenceId === frame.occurrenceId)
      || (placement.kind === 'drawing' && placement.advancePt === 0)
    ));
}

/** Source emptiness is independent of whether a whitespace/hidden run was
 * suppressed during line acquisition. Only the drawing's own metric host
 * and one parser-owned overlapping cell anchor may elect this reference. */
export function wordGridFrameCarrierSourceOwnsNoFlow(paragraph: Pick<ParagraphAcquisitionInput, 'runs' | 'numbering'>): boolean {
  const shapes = paragraph.runs.filter((run) => run.type === 'shape');
  const anchor = shapes.length === 1 ? shapes[0]?.anchorAcquisitionInput : undefined;
  return paragraph.numbering == null && anchor !== undefined
    && anchor.vertical.relativeFrom === 'paragraph' && anchor.wrap.kind === 'none'
    && anchor.behavior.allowOverlap === true && anchor.behavior.layoutInCell === true
    && paragraph.runs.every((run) => (
      run.type === 'shape' || (run.type === 'anchorHost' && run.anchorOccurrenceId === anchor.occurrenceId)
    ));
}

export const WORD_ROTATED_CELL_AUTO_ROW_WRAP = defineCompatibilityRule({
  id: 'word-rotated-cell-auto-row-wrap',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'rotated-cell-row-height-direction-rule-matrix',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'In a fixed-width table, btLr and tbRl cells with 1, 4, or 8 glyphs keep the same automatic row height as a one-line horizontal neighbor by wrapping into additional columns. Two paragraphs and an explicit line break also keep that height when their columns fit. A 20pt top/bottom margin sum adds 20pt, revealing a line-box minimum even when the glyph advance is narrower; a 60pt atLeast minimum and a five-line horizontal neighbor govern their rows; exact remains authored. A horizontal text-direction control keeps its ordinary line. These Word PDF observations cover short text fitting across a 225pt cell in compatibility mode 14; larger content and other compatibility modes are not Office-verified by this matrix.',
});

export const WORD_AUTOFIT_EMPTY_PARAGRAPH_CONTENT_WIDTH = defineCompatibilityRule({
  id: 'word-autofit-empty-paragraph-content-width',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'autofit-empty-paragraph-boundary-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'For table AutoFit content width, Word gives an empty unnumbered paragraph no intrinsic content width regardless of effective right, left, first-line, or hanging indentation. Cell margins still contribute, while whitespace, non-breaking space, visible text, and numbering remain content-bearing controls.',
});

export const WORD_AUTOFIT_OUTER_CELL_MARGIN_BAND = defineCompatibilityRule({
  id: 'word-autofit-outer-cell-margin-band',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'autofit-nowrap-identical-grid-matrix',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'For top-level ordinary AutoFit tables, Word permits a saved grid to extend beyond the text band by the resolved outer cell margins. Two-cell controls with 0, 2.7, and 5.4pt margins, dxa 250/300pt and auto cell widths, short/long text, and noWrap on/off establish that allowance. A Word-produced table with 5.4pt margins but a saved grid exactly at the text band retains that narrower grid on current Word PDF export. A nested table with saved grid overhang retains its containing-cell width rather than receiving the page-table allowance. Auto-width tables are now governed by WORD_AUTOFIT_LEADING_INDENT_BAND (compatibility-mode dependent); for a preferred-width table only a top-level outer-margin overhang already present in tblGrid increases the physical ceiling. Skipped outer grid tracks, leading-margin placement, vertical text, and floating tables remain outside this observation. Windows Word 16.0 build 16.0.20430 retest is consistent within the tested subset; see the officeObservations entry for this rule in scripts/docx-compatibility-microsoft-evidence.json for values and evidence limits.',
});

export const WORD_AUTOFIT_LEADING_INDENT_BAND = defineCompatibilityRule({
  id: 'word-autofit-leading-indent-band',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'autofit-nowrap-page-boundary-compat-mode-matrix',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'For top-level left-justified auto-width AutoFit tables, the measured ceiling is text band - tblInd + outer cell margins in mode 14 (and with the mode omitted), and text band - tblInd in mode 15. The physical page is not a ceiling. Other explicit compatibility modes retain the established solver contract because they are unmeasured. The original controls cover a 468pt page, 0/18pt side margins, 5.4pt outer cell margins, indents of +5.4/+36/-36pt, wide and band-minus-indent grids, and noWrap on/off. Eight additional wrapping controls reverse the content distribution, vary an omitted first tcW versus a 100.55pt dxa first tcW, and vary grid totals of 426.6/432pt on a 432pt band with +5.4pt indent. Neither that cell preference nor the grid-total equality changes the resulting column starts or line partitions within a mode. The second cell remains auto throughout. An all-dxa multirow mode-15 table also agrees with this ceiling: Word PDF cell clipping bounds constrain its width to approximately 482pt, and the fitted 481.9pt width reproduces both inspected two-line paragraphs with the same Calibri face. Substituting a different face can move those line boundaries; retaining the old wider table would hide that font dependency. Four centered/right controls at that same positive indent and band-minus-indent grid match a 437.4pt mode-14 ceiling and a full-band 432pt mode-15 ceiling. In mode 14 those four controls cannot distinguish band-minus-indent-plus-margins from saved-grid-plus-margins. Other nonleading alignment/grid/indent combinations remain unverified and retain their established solver contract; the new nonleading ceiling is gated to grid-plus-indent equaling the band, where both mode-14 hypotheses agree. These borderless PDFs establish column starts and line partitions. Their integer-point clipping rectangles bound cell/table extents but do not establish an exact trailing table edge. Absolute leading-origin differences remain separate from width fitting. Floating, nested, vertical-text and preferred-table-width tables are outside this rule. Windows Word 16.0 build 16.0.20430 retest is consistent within the tested subset; see the officeObservations entry for this rule in scripts/docx-compatibility-microsoft-evidence.json for values and evidence limits.',
});

export const WORD_AUTOFIT_CONTENT_COLUMN_GROWTH = defineCompatibilityRule({
  id: 'word-autofit-content-column-growth',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'autofit-nowrap-page-boundary-compat-mode-matrix',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'In the measured top-level auto-width AutoFit tables, a column without preferred cells grows beyond its initial tblGrid toward maximum content width, up to the occurrence ceiling. Two unpreferred tracks also release saved width beyond a short column maximum. Twelve simultaneous-growth controls in modes 14/15 start at 25/25pt and use breakable ii tokens in ratios 3:2, 1:1 and 5:1, with totals of 30 (nonbinding) and 90 (binding). Nonbinding columns reach their content maxima. Binding column widths match min_i + fraction * (max_i - min_i), where fraction = (ceiling - sum(min)) / sum(max - min), capped at 1. The unequal cases distinguish this from sharing deficits over saved widths, sharing absolute maxima, equal room and sequential growth. Intrinsic maxima must be unbounded by the text band. First-column differences are at most 0.228pt and both cells have identical Word line partitions. Only two nonspanning, wholly unpreferred tracks establish simultaneous sharing; larger grids, spans, and mixed preferred/unpreferred simultaneous growth retain the previous solver result. This is a measured content-interval rule, not an empirical scale factor or an observation of arbitrary multi-column growth. Windows Word 16.0.20527 controls (blk-0898 tables batch: one-row two-cell AutoFit tables with tblW auto, tcW auto and 108-twip margins, jc left/start/center/right/end, LTR and bidiVisual, compatibility mode omitted, 11, 12, 14 and 15) content-fit such a table for every justification and mode, not only for a leading table in modes 14/15. That nonbinding fit therefore applies to a single top-level content-sized table of two unspanned unpreferred tracks (no preferred table or cell width, omitted or dxa legacy left/right margins only, no cell spacing, no row exceptions, row justification or grid skips, not positioned, not a producer outer-edge table) in a single-column body, header or footer for those modes, when the column content maxima fit the established occurrence ceiling together. Where the leading-indent band is unmeasured (nonleading tables, modes 11/12) a binding fit is unmeasured and keeps the previous solve; the width ceilings of such a table keep their previous indent inputs, so only its placement and nonbinding fit change. Windows Word 16.0 build 16.0.20430 retest is consistent within the tested subset; see the officeObservations entry for this rule in scripts/docx-compatibility-microsoft-evidence.json for values and evidence limits.',
});

export const WORD_AUTOFIT_NOWRAP_AUTO_FORCED_FIT = defineCompatibilityRule({
  id: 'word-autofit-nowrap-auto-forced-fit',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'autofit-nowrap-identical-grid-matrix',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'When a two-cell AutoFit table has one auto-width noWrap cell whose unbroken minimum alone exceeds the text band, Word scales the content-only widths to that band. It assigns the extra outer-margin width to each cell in proportion to the other cell content width. The observed matrix varies the ordinary cell text width from 5 to 85pt, outer margins from 0 to 5.4pt, and which side owns noWrap. Short noWrap text that fits the band is a counterexample and retains ordinary proportional fitting. More than two cells and spans are outside this compatibility claim. Windows Word 16.0 build 16.0.20430 retest is consistent within the tested subset; see the officeObservations entry for this rule in scripts/docx-compatibility-microsoft-evidence.json for values and evidence limits.',
});

export function wordAutofitEmptyParagraphHasNoIntrinsicContent(
  paragraph: Pick<ParagraphLayoutSource, 'runs' | 'numbering'>,
): boolean {
  return paragraph.runs.length === 0 && paragraph.numbering == null;
}

export const WORD_EXACT_ROW_HEIGHT_BOTTOM_PADDING = defineCompatibilityRule({
  id: 'word-exact-row-height-bottom-padding',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.180(d)',
  },
  description: 'Word adds the largest bottom cell margin to an exact trHeight instead of treating that margin as part of the authored height.',
});

export const WORD_TABLE_BORDER_LAYER_CASCADE = defineCompatibilityRule({
  id: 'word-table-border-layer-cascade',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.169',
  },
  description: 'During per-side border acquisition, none falls through to a lower-precedence layer while nil remains authored and blocks fallback only on that side.',
});

export const WORD_SPACED_CELL_INSIDE_BORDER_CONFLICT = defineCompatibilityRule({
  id: 'word-spaced-cell-inside-border-conflict',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §§2.1.136, 2.1.138',
  },
  description: 'With non-zero cell spacing, Word retains the narrow conditional tcBorders insideH/insideV conflict against the corresponding table inside border.',
});

export const WORD_TABLE_ORIGIN_COMPATIBILITY = defineCompatibilityRule({
  id: 'word-table-origin-compatibility',
  evidence: { kind: 'office-observation', syntheticFixtureId: 'table-origin-compatibility-and-occupied-grid',
    application: 'Microsoft Word', version: '16.113.2', platform: 'macOS 27.0' },
  description: '528 ordinary-table controls cover modes 11/12/14/15, LTR/RTL, leading/center/end, omitted/zero/positive/negative indentation, explicit zero/asymmetric cell margins, cell overrides, first-row exceptions, fixed/AutoFit, nested contents, borderless and full-band tables. For top-level body tables with positive dxa cell preferences, explicit legacy horizontal margins and zero spacing (fixed auto/dxa table width or AutoFit dxa table width), modes 11/12/14 hang the first cell left margin: leading placement only with authored indentation and anchored to the first row, end placement per row. Center and end ignore indentation in both placement and width fitting, as ECMA-376 §17.4.50 specifies (contrary to MS-OI29500 §2.1.155). Mode 15 removes the margin hang. Windows Word 16.0.20527 single-cell controls (5000-twip dxa table and cell, single borders, with no styles part, a styles part without a table style, or a default Normal Table style carrying tblInd 0 and 108-twip left/right margins; compatibility mode omitted, 11, 12, 14 and 15) extend this: an omitted compatibilityMode hangs exactly like mode 14 (style margins 108, tblCellMar 300/200, tcMar 400/100, tblInd 500, tblInd -300, tblInd 500 with tblCellMar left 200, right-aligned and bidiVisual tables all place text identically in both); a tblInd supplied by the table style hierarchy (the default table style\'s tblInd 0, or a basedOn style\'s tblInd 500) counts as authored indentation (modes 11 and 12 hang the style margin, tblCellMar 300 and tcMar 400 exactly like mode 14), while the same explicit margins without any tblInd in the hierarchy do not hang; a centered table never hangs; and a table nested in a cell keeps its edge at the containing cell\'s content edge in every mode (nested origins stay outside the hang). With horizontal margins omitted at every layer (cell, exception, table, style), the same top-level body structure in modes 11/12/14/15 was measured only on controls-1664 page 7 (fixed 108/180pt dxa tracks, 288pt dxa tblW, LTR and bidiVisual, justification omitted/start/end, tblInd 0 and +36pt): end justification places the table identically with tblInd 0 and +36pt in both directions, so ECMA-376 §17.4.50 ignores the indent there as well, while omitted/start justification applies it. Only that end class joins the observation; omitted-margin right/center rows, an omitted mode, mode 13, nonpositive or non-dxa cell widths and nested tables keep the previous origin behavior. Spaced tables remain unresolved and retain the previous origin behavior (the content inset of omitted margins is word-table-unspecified-horizontal-cell-margin); logical start/end margin spellings and other compatibility modes are outside the measured scope. Border rasterization is separate from nominal grid geometry. Windows Word 16.0 build 16.0.20430 retest is consistent within the tested subset; see the officeObservations entry for this rule in scripts/docx-compatibility-microsoft-evidence.json for values and evidence limits.',
});

export const WORD_FIXED_UNUSED_LEADING_GRID = defineCompatibilityRule({
  id: 'word-fixed-unused-leading-grid',
  evidence: { kind: 'office-observation', syntheticFixtureId: 'table-origin-compatibility-and-occupied-grid',
    application: 'Microsoft Word', version: '16.113.2', platform: 'macOS 27.0' },
  description: 'The original fixed gridBefore controls in modes 11/12/14/15 and 48 follow-ups in modes 14/15 show that a leading track skipped by every row has zero layout width. Fixed auto table width retains the occupied 108/180pt tracks; dxa 324pt scales them to 121.5/202.5pt. A preceding row occupying the 36pt track preserves it and ordinary skipped-row placement. wBefore of 0/18/36pt has no effect in either class. This refines ECMA-376 §§17.4.15 and 17.18.87 only for top-level body nonspaced fixed tables with auto/dxa table width and positive dxa cell preferences; nested and nonbody stories, percentage table width, AutoFit, trailing skips, spans, and nonpositive or missing cell preferences remain outside the observation. Windows Word 16.0 build 16.0.20430 retest is consistent within the tested subset; see the officeObservations entry for this rule in scripts/docx-compatibility-microsoft-evidence.json for values and evidence limits.',
});

/**
 * ECMA-376 §17.4.50: tblInd applies only to a leading (left/start/omitted)
 * table justification. WORD_TABLE_ORIGIN_COMPATIBILITY enables this for its
 * measured origin classes; this rule extends only the indent part (not the
 * legacy margin hang) to the further measured table structures below.
 */
export const WORD_TABLE_NONLEADING_INDENT_SCOPE = defineCompatibilityRule({
  id: 'word-table-nonleading-indent-scope',
  evidence: { kind: 'office-observation', syntheticFixtureId: 'table-nonleading-indent-classes',
    application: 'Microsoft Word', version: '16.0.20527', platform: 'Windows' },
  description: 'ECMA-376 §17.4.50 ignores tblInd unless the table justification is leading. Windows Word 16.0.20527 controls (one-row two-cell 4000-twip tables, LTR and bidiVisual, jc left/start/center/right/end, tblInd omitted, +720, -360 and a table-style tblInd 720; compatibility mode omitted, 11, 12, 14 and 15) place and size every center, right and end table identically for all four indents, while left/start tables move by the indent (mirrored under bidiVisual). This holds for omitted horizontal margins (fixed and AutoFit dxa tables), a right-only tblCellMar, explicit 108-twip margins with percentage, auto, nil, missing and zero-dxa cell widths, a percentage table width, an AutoFit auto-width table with dxa or auto cells, 60-twip cell spacing with explicit or omitted margins, tables nested once in a body table cell (explicit and omitted margins), and top-level tables in a header, a footer, a page-anchored text box and a footnote (explicit and omitted margins, LTR and bidiVisual). The indent rule alone extends to these classes; the legacy margin hang of word-table-origin-compatibility keeps its own measured scope. Horizontal margins must be omitted or authored as dxa legacy left/right elements at every layer (cell, exception, table, style), the kind resolved as §17.4.87 does (an omitted type is dxa, a percent measurement is pct); nil/auto/pct widths and logical start/end spellings were not measured. An adjacent-table (§17.4.37) group is one table and qualifies with a common frame and grid. An AutoFit table without any effective preferred table or cell width (tblW auto, tcW auto; zero, nil and auto widths count as absent, as in the column solver) also ignores the indent in these controls. It joins this rule within the class word-autofit-content-column-growth content-fits (two unspanned unpreferred tracks in a single top-level page-owned table), including a binding fit that keeps its previous width, since every measured class ignores the indent as §17.4.50 specifies; outside that class the width defect remains and it keeps the previous behavior. Leading tables move by their indent, except that word-nested-negative-indent ignores a negative indent on a once-nested fixed table. Compatibility mode 13 is not a value Word opens; other modes, row-level justification or table exceptions, authored gridBefore/gridAfter (including offsets the parser clears against the grid), floating or vertical-text tables, tables in a quarter-turned (vertical bodyPr) text box, deeper nesting and endnotes were not measured and keep the previous behavior.',
});

/** WORD_TABLE_NONLEADING_INDENT_SCOPE: the measured structures where a
 * nonleading table ignores its tblInd outside the origin observation. */
export function wordNonleadingIndentScope(input: Readonly<{
  compat: Pick<DocxCompatProfile, 'tableOriginPlacementMode'>;
  story: string | undefined;
  topLevel: boolean;
  tableCellDepth: number;
  verticalText: boolean;
  commonFrame: boolean;
  marginClass: boolean;
  outerEdgeIndent: boolean;
  ordinaryFlow: boolean;
  positioned: boolean;
  rowOverrides: boolean;
  gridSkips: boolean;
  /** AutoFit with no preferred table or cell width outside the content-fitted
   * scope (excluded, see the rule). */
  unfittedContentSizedAutoFit: boolean;
}>): boolean {
  const story = input.story;
  const measuredStory = story === 'body' || story === 'header' || story === 'footer'
    || story === 'footnote' || story === 'textbox';
  const measuredNesting = input.topLevel || (story === 'body' && input.tableCellDepth === 1);
  return input.compat.tableOriginPlacementMode && measuredStory && measuredNesting
    && !input.verticalText && input.commonFrame && input.marginClass && !input.outerEdgeIndent && input.ordinaryFlow
    && !input.positioned && !input.rowOverrides && !input.gridSkips && !input.unfittedContentSizedAutoFit;
}

/**
 * WORD_TABLE_ORIGIN_COMPATIBILITY measured the legacy left-margin hang for
 * top-level body tables with positive dxa cell and table widths. This rule
 * extends the same hang to the further structures measured below.
 */
export const WORD_TABLE_MARGIN_HANG_SCOPE = defineCompatibilityRule({
  id: 'word-table-margin-hang-scope',
  evidence: { kind: 'office-observation', syntheticFixtureId: 'table-margin-hang-classes',
    application: 'Microsoft Word', version: '16.0.20527', platform: 'Windows' },
  description: 'Windows Word 16.0.20527 controls (one-row two-cell 4000-twip-grid tables with 108-twip dxa tblCellMar left/right, LTR and bidiVisual, jc left/start/center/right/end, tblInd omitted, +720, -360 and a table-style tblInd 720; compatibility mode omitted, 11, 12, 14 and 15) hang the first cell left margin exactly as word-table-origin-compatibility does for its measured class: in every mode except 15, a leading table with an indent (direct or from the table style) moves left by the row left margin, and a right or end table moves right by it; center tables and mode 15 do not hang. In the body this holds whatever the widths: tcW pct, auto, nil, omitted and dxa 0; tblW pct and auto; fixed and AutoFit. Top-level tables of the explicit-margin class hang identically in a header, a footer and a footnote. In a page-anchored text box only the leading hang applies; right and end tables keep their trailing edge at the inner band. A percentage tblW in the hanging modes is a share of the text band plus the first row outer margins (40% = 197.08pt = 0.4 x (481.9 + 10.8)pt, against 0.4 x 481.9pt in mode 15); this width part applies to page-owned stories (single-column body, header, footer). The rule applies to single tables inside word-table-nonleading-indent-scope; multiple rows anchor and hang per row as word-table-origin-compatibility measured, while adjacent-table (§17.4.37) groups keep the previous behavior. Cell spacing (tracked separately), omitted, logical or non-dxa margins, row-level justification or exceptions, authored gridBefore/gridAfter, nested, floating and vertical-text tables, quarter-turned text boxes, endnotes and other compatibility modes were not measured and keep the previous behavior.',
});

/** ECMA-376 §§17.4.43–45 establish separated edges and spacing precedence;
 * the following numeric geometry and compatibility-mode placement are observed
 * Word behavior, not a normative interpretation of the spacing measurement. */
export const WORD_TABLE_CELL_SPACING_GEOMETRY = defineCompatibilityRule({
  id: 'word-table-cell-spacing-geometry',
  evidence: { kind: 'office-observation', syntheticFixtureId: 'table-cell-spacing-geometry',
    application: 'Microsoft Word', version: '16.0.20527', platform: 'Windows' },
  description: 'Windows Word controls use three 2000-twip tracks, tblW 6000, spacing 0/60/200 twips, omitted/zero/108-twip margins, no borders or uniform single borders sz4/sz24, fixed layout or content-sized AutoFit, left/center/right justification, no indent or +720, LTR/bidiVisual and omitted/14/15 modes; the blk-0898 table controls add one-row two-column fixed tables (60-twip spacing, 108-twip or omitted margins, single sz4 borders without insideH) in modes 11 and 12, which place exactly like mode 14, and with tblInd -360, which moves a leading table like a positive indent. Positive spacing s reserves 2s at every outer and shared boundary. Each cell contains its own full border width t before its margins. Fixed cell widths carve (n+1)*2s from the resolved grid proportionally; a horizontal span retains its interior gaps. Vertical extents include the outer table top/bottom border as well as both cell borders. Mode 15 LTR boxes move inward by half the aligned outer border (center stays centered); bidi boxes shift a further half border left. Omitted/11/12/14 modes have no border shift: leading tables with authored indentation and trailing tables hang the full 2s+t+margin content inset in LTR, or s+t+margin in bidi, leaving exactly s there. Unindented leading and centered tables do not hang. AutoFit content-sized tables fit all unspanned unpreferred tracks to their content maxima only when those maxima and all gaps fit the occurrence ceiling and text band together; binding multitrack distribution is not established and retains the complete previous geometry, including content constraints, width ceiling, origin and row heights. Omitted margins retain the library zero instead of approximating the observed Word half-point quanta. Fixed proportional widths were measured within 0.1pt at s=3/10pt; these are ideal geometry rules, not fits to raster/quantized offsets. Scope is one ordinary top-level table in a horizontal single-column body, uniform positive table-level tblPr spacing (including style-inherited spacing) and uniform absent or single borders (an inside border a one-row or one-column table cannot show is ignored), legacy dxa or omitted margins, fixed dxa width or content-sized AutoFit, modes omitted/11/12/14/15. Row-level trPr spacing and tblPrEx spacing retain their previous geometry even when all effective row values match. Asymmetric borders, any cell border override, row exceptions/justification/skips, vertical merges, nested tables or contents, floating tables, other stories, vertical text, gaps exhausting the fixed width or AutoFit band, authored row heights and other modes retain their previous geometry. Horizontal spans join only the fixed width formula; AutoFit spans remain excluded. Pagination of the measured automatic rows uses the same page-local outer gaps and border footprints; Office continuation placement was not separately measured.',
});

/** An independent border box is established only for the measured uniform
 * table-border class. A cell override, even one matching the table, is outside
 * the evidence; never infer its ownership from an equal effective width. */
export function wordUniformSpacedTableBorderPt(table: TableLayoutSource): number | null {
  // An inside border a one-row (insideH) or one-column (insideV) table cannot
  // show has no geometry, so it does not break uniformity.
  const singleRow = table.rows.length <= 1;
  const singleColumn = table.rows.every((row) => row.cells.length <= 1);
  const widths = Object.entries(table.borders ?? {}).filter(([edge]) =>
    !(edge === 'insideH' && singleRow) && !(edge === 'insideV' && singleColumn)).map(([, border]) => {
    if (!border || border.style === 'nil' || border.style === 'none') return 0;
    return border.style === 'single' && Number.isFinite(border.width) ? border.width : null;
  });
  if (widths.some((width) => width === null || width !== widths[0])) return null;
  if (table.rows.some((row) => row.cells.some((cell) =>
    Object.values(cell.borders).some((border) => border != null)
    || cell.vMerge != null || (cell.textDirection != null && cell.textDirection !== 'lrTb')
    || cell.content.some((block) => block.type === 'table')))) return null;
  return widths[0] ?? 0;
}

/** Signed logical-axis displacement; alignedTableOrigin reverses it for bidi.
 * WORD_TABLE_CELL_SPACING_GEOMETRY keeps border placement and the legacy hang
 * separate, with no half-border shift in the legacy modes. */
export function wordSpacedTableTranslationPt(input: Readonly<{
  compat: Pick<DocxCompatProfile, 'tableMode15Geometry'>; justification: string | null | undefined; bidi: boolean;
  indentPt: number; indentAuthored: boolean; spacingPt: number; borderPt: number;
  leadingMarginPt: number; trailingMarginPt: number;
}>): number {
  const leading = input.justification == null || input.justification === 'left' || input.justification === 'start';
  const center = input.justification === 'center';
  const indent = leading ? input.indentPt : 0;
  if (input.compat.tableMode15Geometry) {
    const physicalBorderShift = center ? 0 : (leading !== input.bidi ? 1 : -1) * input.borderPt / 2;
    const physicalShift = physicalBorderShift - (input.bidi ? input.borderPt / 2 : 0);
    return indent + (input.bidi ? -physicalShift : physicalShift);
  }
  if (center || (leading && !input.indentAuthored)) return indent;
  const margin = leading ? input.leadingMarginPt : input.trailingMarginPt;
  const hang = (input.bidi ? input.spacingPt : 2 * input.spacingPt) + input.borderPt + margin;
  return indent + (leading ? -hang : hang);
}

/** ECMA-376 §17.4.50 lets tblInd move a leading table either way. */
export const WORD_NESTED_NEGATIVE_INDENT = defineCompatibilityRule({
  id: 'word-nested-negative-indent',
  evidence: { kind: 'office-observation', syntheticFixtureId: 'table-nonleading-indent-classes',
    application: 'Microsoft Word', version: '16.0.20527', platform: 'Windows' },
  description: 'Windows Word 16.0.20527 controls nest one-row two-cell 4000-twip fixed tables once in a 9000-twip fixed body table cell (explicit 108-twip and omitted horizontal margins, LTR and bidiVisual, compatibility mode omitted, 11, 12, 14 and 15). A left or start table with tblInd -360 is placed exactly like its twin without tblInd, at the containing cell content edge (mirrored under bidiVisual), while +720 and a table-style tblInd 720 move it by 36pt. The negative indent is therefore ignored, not clamped to the cell edge. Center, right and end tables ignore every indent under word-table-nonleading-indent-scope. Scope: fixed tables without cell spacing inside the structures of word-table-nonleading-indent-scope, nested once in a body table cell; AutoFit or spaced nested tables, deeper nesting, other stories and other modes were not measured and keep the previous behavior.',
});

export type WordTableMarginHangScope = 'none' | 'leading' | 'all';

/** WORD_TABLE_MARGIN_HANG_SCOPE: which rows of a table outside the origin
 * observation hang the legacy left margin. */
export function wordTableMarginHangScope(input: Readonly<{
  compat: Pick<DocxCompatProfile, 'tableOriginPlacementMode'>;
  story: string | undefined;
  topLevel: boolean;
  verticalText: boolean;
  commonFrame: boolean;
  /** Every row has an authored legacy left margin and only dxa margins. */
  marginClass: boolean;
  spacing: boolean;
  outerEdgeIndent: boolean;
  ordinaryFlow: boolean;
  positioned: boolean;
  rowOverrides: boolean;
  gridSkips: boolean;
}>): WordTableMarginHangScope {
  const story = input.story;
  if (!input.compat.tableOriginPlacementMode || !input.topLevel || input.verticalText
    || !input.commonFrame || !input.marginClass || input.spacing || input.outerEdgeIndent
    || !input.ordinaryFlow || input.positioned || input.rowOverrides || input.gridSkips) return 'none';
  if (story === 'body' || story === 'header' || story === 'footer' || story === 'footnote') return 'all';
  return story === 'textbox' ? 'leading' : 'none';
}

// The table-origin observation is deliberately restricted to the exported
// mode boundaries (DocxCompatProfile.tableOriginMeasuredMode). Its placement
// scope adds an omitted compatibilityMode, which places tables like mode 14
// (tableOriginPlacementMode); WORD_FIXED_UNUSED_LEADING_GRID keeps the
// explicit modes only.

/**
 * ECMA-376 §17.4.34 / §17.4.11 give tables whose style hierarchy never
 * specifies a leading/trailing cell margin 115 twips. Word lays out 0. The
 * DOCX parser applies this fallback (parser.rs, table cell margins), after
 * direct tblCellMar and the associated (or §17.7.4 default) table style.
 */
export const WORD_TABLE_UNSPECIFIED_HORIZONTAL_CELL_MARGIN = defineCompatibilityRule({
  id: 'word-table-unspecified-horizontal-cell-margin',
  evidence: { kind: 'office-observation', syntheticFixtureId: 'table-cell-margin-default-and-origin',
    application: 'Microsoft Word', version: '16.0.20527', platform: 'Windows' },
  description: 'A left or right cell margin that neither tcMar, tblCellMar nor the table style hierarchy specifies is 0, not the 115 twips of ECMA-376 §§17.4.34/17.4.11. Windows Word 16.0.20527 controls (5000-twip single-cell tables with no styles part, or a styles part without any table style, in compatibility modes omitted/11/12/14/15): text starts at the table edge and ends at its trailing edge for default margins, tblInd 500, centered, right-aligned, bidiVisual, tblInd -300 and nested 3000-twip tables, while explicit tblCellMar 0, 300/200 and tcMar 400/100 apply as authored; a table with only tblCellMar left 200 keeps a 0 right margin. The same 0 margin (default and tblInd 500 tables, modes omitted and 15) holds in a header, a footer, an inline text box and a footnote, identically to a body control in the same document. A default Normal Table style carrying 108-twip margins supplies 108 as usual. A table style that exists but omits tblCellMar was not measured separately; it reaches the same fallback.',
});

export const WORD_EXACT_ROW_VERTICAL_CLIP_ONLY = defineCompatibilityRule({
  id: 'word-exact-row-vertical-clip-only',
  evidence: {
    kind: 'library-policy',
    reference: 'packages/docx/src/layout/table.test.ts#clips an overflowing merged owner when every row in its span is exact',
  },
  description: 'Preserve the established exact-row overflow behavior that clips the owned vertical interval without clipping nested table ink horizontally to the cell box. Library policy, no Office measurement: vertical-only clipping of exact-height rows; [MS-OI29500] §2.1.180(d) covers only the exact height.',
});

export const WORD_OVER_PAGE_CANT_SPLIT_CLIP = defineCompatibilityRule({
  id: 'word-over-page-cant-split-clip',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.120',
  },
  description: 'Word starts an over-page cantSplit row on a fresh page and clips its overflow instead of synthesizing a row continuation. The note does not distinguish a nested row; as a library reading without an Office measurement, the page of a nested row is the fresh band of its table: the content band of the enclosing cell on a fresh page, not the page band, which no nested table inside a cell with margins, rules or repeated headers is ever given.',
});

/** Consumed by table-pagination.ts partialRow (the re-selection, and the
 * selection every enclosing row keeps: TableFragmentResult.wholeSource),
 * nestedTableFragmentContext (a nested table's own fresh band) and by the
 * containment limit of table.ts measureTableCellBlockFlowHeightPt. */
export const LIBRARY_OVER_PAGE_TABLE_CELL_UNIT_CLIP = defineCompatibilityRule({
  id: 'library-over-page-table-cell-unit-clip',
  evidence: {
    kind: 'library-policy',
    reference: 'packages/docx/src/layout/table-pagination.test.ts#clips an over-page first cell line on a fresh page and continues the cell on the next page',
  },
  description: 'Library policy, no Office measurement, applied in every compatibility mode: ECMA-376 §17.4.6 lets a splittable row continue across pages but defines nothing for a cell unit that no page can hold. When a row fragment on a fresh band (nothing but repeated header rows above it, so relocating cannot gain space) would otherwise select nothing, whether because no cell progresses or because word-parallel-paragraph-row-cut rejects a cut that leaves a blocked paragraph cell behind, the row is re-selected once: a layoutInCell containment interval (§20.4.2.3) whose own extent exceeds the fresh content band stops charging admission, so the text around its drawing keeps flowing, and a first line or block still taller than the band is retained alone. The recovered cells are materialized as admitted: a left-out containment interval neither grows the row nor offsets the text below a drawing that starts above it, and the cells are laid out from their top because a centered or bottom vAlign (§17.4.84) would align against the over-band row, so every admitted line is painted inside the band. The fragment is clipped at the band end and its remaining content continues on the next page. The same holds at every nesting depth: the fresh band of a nested table is the content band of its cell on a fresh page (the page band less every enclosing cell margin, row rule, cell spacing and repeated header row), the re-selection governs only the cell units of the row itself, so a table nested in a recovered cell recovers in its own band, and every enclosing row and table materializes a recovered or clipped fragment exactly as selected instead of its source layout, whether or not the row completes; floating-table anchors are resolved against the same placement. Content below the band end (part of an over-band unit, or of a drawing) is not painted. This is a recovery from an otherwise non-converging pagination, not a claim about Word or the specification; the warning channel is reserved for caller opt-outs, so it reports nothing.',
});

export const WORD_OVER_PAGE_CELL_BREAK_OCCUPANCY = defineCompatibilityRule({
  id: 'word-over-page-cell-break-occupancy',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'over-page-cell-followed-by-authored-break',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'A table-cell paragraph taller than the body band counts its invisible continuation page before a following authored page break. A fitting paragraph does not. This holds with and without cantSplit; without the authored break, the next paragraph starts at the top of the continuation page. Tested at 500pt and 800pt against a 648pt body band, so farther overflow remains an inferred geometric extension.',
});

export const WORD_AUTHORED_ROW_HEIGHT_PAGE_BOUNDARY = defineCompatibilityRule({
  id: 'word-authored-row-height-page-boundary',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'ordinary-table-row-height-boundary-matrix',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'Word relocates an ordinary splittable exact-height row, or an atLeast row whose authored minimum governs its complete height, when that height exceeds the remaining page band and fits a fresh page. A shorter atLeast minimum permits content fragmentation; an auto row may fragment. Tested with cantSplit on/off, fitting/overflow bands, and keepLines/widow controls. Repeated headers and atLeast rows expanded by content are outside this observation.',
});

export const WORD_CELL_OWNED_ANCHOR_PAGE_CUT = defineCompatibilityRule({
  id: 'word-cell-owned-anchor-page-cut',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'cell-owned-anchor-page-band-boundary',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'For an atLeast row without cantSplit containing a layoutInCell, allowOverlap wrapNone image, Word permits the image past its row border while the image stays in the page body band. Controlled preceding-spacing cases at 0pt and 40pt kept the row; at 80pt Word moved the complete row to the next page. A separate near-edge control confirmed that trailing empty cell paragraphs may still form an empty row continuation. Other anchor wrap/row-height combinations and over-page images are outside this observation.',
});

/** Compatibility choice governed by {@link WORD_CELL_OWNED_ANCHOR_PAGE_CUT}.
 * Geometry detection remains in the table paginator; this gate only chooses
 * the observed Word page cut when a fresh page offers more room. */
export function wordDefersCellOwnedAnchorPastPageBand(input: Readonly<{
  compatibility: 'word' | 'standard';
  availableHeightPt: number;
  freshPageHeightPt: number;
  epsilonPt: number;
}>): boolean {
  return input.compatibility === 'word'
    && input.availableHeightPt + input.epsilonPt < input.freshPageHeightPt;
}

export const WORD_PARALLEL_PARAGRAPH_ROW_CUT = defineCompatibilityRule({
  id: 'word-parallel-paragraph-row-cut',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'parallel-paragraph-row-cut-boundary-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'When a page cut crosses a row containing parallel paragraph content, Word emits no cell content unless every unfinished paragraph cell can reach at least its first legal line or block boundary in that page band. The observed rule does not cover nested-table child boundaries.',
});

export const WORD_POSITIONED_TABLE_ADJACENCY_EXCLUSION = defineCompatibilityRule({
  id: 'word-positioned-table-adjacency-exclusion',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.149(a)',
  },
  description: 'Word excludes effectively positioned tables from the logical adjacent-table sequence before retained layout consumes the parser-owned sequence identity.',
});

export const WORD_TABLE_BORDER_WEIGHT_PRECEDENCE = defineCompatibilityRule({
  id: 'word-table-border-weight-precedence',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.169',
  },
  description: 'Use the documented Word border numbers for shared-cell conflict weight and force dotted and dashed borders to a complete weight of one.',
});

export const WORD_OMITTED_ROW_HEIGHT_RULE_AT_LEAST = defineCompatibilityRule({
  id: 'word-omitted-row-height-rule-at-least',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.180',
  },
  description: 'Treat an omitted trHeight hRule as atLeast while retaining an explicitly authored auto rule as authored input.',
});

export const WORD_COLLAPSED_BORDER_ROW_TRACK_FOOTPRINT = defineCompatibilityRule({
  id: 'word-collapsed-border-row-track-footprint',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'issue-1904-table-rule-extent',
    application: 'Microsoft Word',
    version: '16.0.20430',
    platform: 'Windows',
  },
  description: 'With collapsed cell boundaries each horizontal rule counts once, at its full width, inside the table box: a row\'s top rule (the table top rule for the first row, the winning inside rule otherwise) lies at the top of that row and its cell content starts below it, and the table\'s bottom rule is appended below the last row with the content ending above it. An automatic or at-least row adds its top rule to its content or minimum height; an exact row\'s authored height already holds its top rule, and its content is clipped to that authored box. Page-local slices treat their first and last rows as the table\'s. Cell spacing keeps independent edges out of this footprint. Evidence: issue #1904 Windows controls (body and full-width page-anchored floating tables; two exact, at-least and automatic rows; no borders, all 0.5/3/6pt, outer-only 3pt and inside-only 3pt; modes 14 and 15): all 38 match within 0.3pt. Word 16.111.1 on macOS had been observed charging half of each adjacent rule to automatic and at-least rows and none to exact rows; the Windows behaviour is followed where the two differ.',
});

export const WORD_EFFECTIVE_FLOATING_TABLE_POSITIONING = defineCompatibilityRule({
  id: 'word-effective-floating-table-positioning',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.162',
  },
  description: 'Use parser-retained effective positioning status rather than lexical tblpPr presence to decide whether a table leaves ordinary flow.',
});

export const WORD_TABLE_CELL_SPACING_SCOPE_SHADOW = defineCompatibilityRule({
  id: 'word-table-cell-spacing-scope-shadow',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §§2.1.152, 2.1.153, 2.1.154',
  },
  description: 'At each table-cell-spacing precedence scope, pct, auto, and nil resolve to zero and shadow lower scopes instead of being treated as absent.',
});

export const WORD_TABLE_MARGIN_SCOPE_SHADOW = defineCompatibilityRule({
  id: 'word-table-margin-scope-shadow',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §§2.1.116, 2.1.125, 2.1.146, 2.1.177',
  },
  description: 'Preserve the documented scope-specific treatment of non-dxa table cell margins: leading/trailing defaults may resolve to zero while cell/exception and nil top/bottom values remain ignored.',
});

export const WORD_TABLE_CELL_ZERO_NIL_WIDTH_AUTO = defineCompatibilityRule({
  id: 'word-table-cell-zero-nil-width-auto',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.171(a-b)',
  },
  description: 'Interpret tcW with a zero w attribute, including dxa and pct, or nil type as auto. Intrinsic measurement and column constraints share the effective cell preference; noWrap uses an unbroken content minimum for these automatic cells instead of dxa width protection.',
});

export const WORD_FIRST_ROW_TABLE_EXCEPTION_SCOPE = defineCompatibilityRule({
  id: 'word-first-row-table-exception-scope',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §§2.1.156, 2.1.158, 2.1.167',
  },
  description: 'Apply the supported first-row table-property exception facts at table scope. An authored first-row tblPrEx/tblW shadows the body width, including auto/nil/zero clearing it; width ceiling selection and column fitting share this effective preference.',
});

export const WORD_HIDDEN_CELL_PARAGRAPH_JOINS_NEXT = defineCompatibilityRule({
  id: 'word-hidden-cell-paragraph-joins-next',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'hidden-cell-paragraph-controls',
    application: 'Microsoft Word',
    version: '16.0.20527',
    platform: 'Windows',
  },
  description: 'A table-cell paragraph whose paragraph mark is vanished (ECMA-376 §17.3.1.29 mark run properties, §17.3.2.41 vanish) and which holds no ink lays out nothing in the cell: no line, no stored or automatic spacing, and it is no neighbour of the blocks around it, as in body flow. The cell\'s last paragraph carries the end-of-cell mark and is kept as an ordinary line even when vanished. ECMA-376 does not describe the join; this is observed Word behaviour. Evidence: Windows Word 16.0.20527 PDFs of synthetic cells (Arial, exact 12pt lines, stored before/after 3/8pt, compatibility mode 15) with doNotUseHTMLParagraphAutoSpacing off and on, one variant per row between plain reference rows. Dropped: one and three leading vanished empty paragraphs before an automatic paragraph (it then opens the cell: its automatic before is dropped without the setting, its stored 3pt kept with it); a leading one whose own stored before/after is 5/11pt (its spacing is dropped too); one between two automatic paragraphs (gap 14pt without the setting, 8 + 3pt with it) and between two stored 3/8pt paragraphs (8pt, 11pt); one between an automatic paragraph and a nested table (the paragraph keeps its after before the table); and a leading one before a nested table (the table opens the cell). Kept as a 12pt line: a trailing vanished empty paragraph (the paragraph before it keeps its after) and the sole vanished paragraph of a cell; control 052 of the automatic-spacing matrix agrees. Not covered by this rule: a paragraph with visible text and a vanished mark, which Word joins onto the next paragraph\'s line; which paragraph\'s properties govern that joined line is not measured, so it keeps its own line here, as in body flow.',
});

export const WORD_TRAILING_STRUCTURAL_CELL_MARKER = defineCompatibilityRule({
  id: 'word-trailing-structural-cell-marker',
  evidence: {
    kind: 'library-policy',
    reference: 'packages/docx/src/layout/compatibility.test.ts#drops only an empty trailing paragraph after a non-paragraph cell block',
  },
  description: 'Exclude the required empty trailing cell paragraph from row-height and vertical-alignment measurements when it follows a visible non-paragraph block. Library policy, no Office measurement: structural trailing cell marker inventory (801aa7986); no evidence.',
});

export const WORD_CELL_VERTICAL_ALIGNMENT_INK_BLOCK = defineCompatibilityRule({
  id: 'word-cell-vertical-alignment-ink-block',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/cell-valign-leading-spacing.test.ts#inked block is vertically centred in the cell (midpoint = cell midpoint)',
  },
  description: 'Center or bottom-align the visible cell ink block without charging the first paragraph spaceBefore or final paragraph spaceAfter at the cell edges. Evidence (observed Office behaviour): Word PDF of a corpus document in PR #552 (midpoint placement); ECMA-376 §17.4.83 vAlign does not define the ink block.',
});

export const WORD_VERTICAL_MERGE_TERMINAL_BORDER = defineCompatibilityRule({
  id: 'word-vertical-merge-terminal-border',
  evidence: {
    kind: 'library-policy',
    reference: 'packages/docx/src/cell-border-conflict-render.test.ts#uses the final continuation cell border at the bottom of a vertical merge',
  },
  description: 'Resolve the bottom edge of a vertically merged cell from its terminal continuation cell before applying shared-edge conflict rules. Library policy, no Office measurement: terminal border of a vertical merge (PR #1032); the ECMA-376 §17.4.66 tcBorders note does not address vMerge.',
});

export const WORD_VERTICAL_MERGE_INTERIOR_BORDER_FOOTPRINT = defineCompatibilityRule({
  id: 'word-vertical-merge-interior-border-footprint',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'vmerge-interior-border-matrix',
    application: 'Microsoft Word',
    version: '16.0.20527',
    platform: 'Windows',
  },
  description: 'ECMA-376 §17.4.85 vMerge joins cells into one merged cell, so no horizontal rule is drawn between its rows, and §17.4.66 tcBorders does not say whether a merged constituent\'s own bottom border still takes space. Windows Word 16.0.20527 controls (4-row tables of 20pt atLeast rows, Times New Roman 11pt, vAlign bottom text; two- and three-row merges in the first column, ending on or before the last row; single 0.5pt and 3pt borders; on and off a 20pt line grid; with and without adjustLineHeightInTable; compatibility modes 14 and 15) show that the bottom border of every constituent except the last one is charged to the row below it like an ordinary rule between rows, though it is not drawn: a 0.5pt border on the first merged cell makes the next row 0.5pt taller and moves its text down 0.5pt (3pt: 3pt), a three-row merge charges both interior borders, a taller next row still grows by the border, and a border only on the last constituent charges nothing inside the merge. The table bottom keeps the terminal continuation border (WORD_VERTICAL_MERGE_TERMINAL_BORDER). The controls set the borders on the cells themselves; a table-level insideH reaching a merged cell was not measured and charges nothing.',
});

/** WORD_VERTICAL_MERGE_INTERIOR_BORDER_FOOTPRINT: the width a merged
 * constituent's own bottom border charges to the row below inside the merge. */
export function wordVerticalMergeInteriorBorderPt(bottom: TableBorderInput | null): number {
  void WORD_VERTICAL_MERGE_INTERIOR_BORDER_FOOTPRINT;
  if (!bottom || bottom.authoredStyle === 'nil' || bottom.authoredStyle === 'none') return 0;
  return Number.isFinite(bottom.widthPt) && bottom.widthPt > 0 ? bottom.widthPt : 0;
}

export const WORD_VERTICAL_SECTION_UPRIGHT_BLOCK_TABLE = defineCompatibilityRule({
  id: 'word-vertical-section-upright-block-table',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'vertical-section-upright-block-table',
    application: 'Microsoft Word',
    version: '16.0.20527',
    platform: 'Windows',
  },
  description: 'ECMA-376 §17.6.20 gives a section its text direction and says nothing about a block table in it. Word keeps a block table of a tbRl or tbLrV section upright (horizontal cells) and places it as a zero-advance block of the vertical flow. (1) Position: the outer edge of its near border rule is the current block position (the right edge in tbRl, the left edge in tbLrV), so the grid starts half that rule further on, and it adds no block advance: the next paragraph takes the next line strip. The table\'s box below runs between the outer edges of its near and far border rules. Along the inline axis it takes the first position from the column top at which its box clears the earlier tables it meets, an earlier table being met until the later one\'s box starts at least 14.28pt past the far edge of its box: a table meeting an earlier one stacks below it, a third table meeting only the stacked one returns to the column top, and a table that fits the band but not the space below the tables it meets is placed at the column top that distance past the nearest of them instead, still without advancing the flow. The table\'s indent (§17.4.50) and alignment (§17.4.28) do not move it. (2) Wrap: the table\'s box excludes text from its near edge inclusive to 7.14pt past its far edge. Every line strip whose block interval meets that range flows around it: beside a table at the column top the strip starts at the table\'s bottom edge; a strip meeting only a stacked table starts at the column top, uses the space above that table and continues below it. The strip laid out just before the table ends exactly where the table starts and wraps too; strips before that one do not. Text with no room beside the table resumes at the column top 7.14pt past the far edge of its box. (3) Fit: a table wider than the block extent left on its page starts the next page in a one-column section and stays, overflowing, in a two-column section. A table taller than the inline extent below its start starts the next page when content precedes it; on a fresh page whole rows fill the band and the remaining rows continue on the next page, where the following text flows beside the last rows. A row taller than the band takes a page of its own and is cut at the band end. A two-column section never splits or moves the table: one taller than its column stays whole where the flow reached it, and one that would pass the bottom page edge is raised to end there, no higher than the page top. Rows a table continues with on a later page meet and exclude like a table of their own. A keepNext paragraph before a table that has to start another page moves there first. Evidence: Windows Word 16.0.20527 PDFs of synthetic A4 tbRl and tbLrV sections (margins left 40, right 90, top 60, bottom 80pt, 18pt line grid, 10.5pt one-line paragraphs without spacing, compatibility mode 15) with single-bordered tables 180pt wide of 14.4pt rows. First set, both directions: table after one paragraph (table y 60–88.7pt, near edge at block 18pt; all three strips start at y 88.7); table first; table after six paragraphs (strips 1–5 at y 60, the sixth and the following ones at 88.7); paragraph, table, paragraph, table, paragraph (second table at y 88.7–117.5; the strips meeting it start at 117.5); an 80-character paragraph before the table (only its last strip starts at 88.7). Second set: the strip starting exactly at the table\'s far edge starts at 88.7 and the next at 60, with and without table borders; a second table clear of the first sits at y 60, one starting exactly at the first table\'s far edge at 88.7; a paragraph with 6pt space after moves the table 6pt along the block axis and its strip starts at 60; a 21pt first-line indent is added to the table\'s bottom edge, a 21pt left indent ending above it is not; a btLr section matches tbRl; a too-wide table after 20 paragraphs starts page 2 at the margin; a 60-row table after one paragraph starts page 2 with rows 1–49 (y 60–752.4) and ends on page 3 (y 60–215.8) with the following paragraphs starting at 215.8. Third set: with tables X at block 18–198 and Y at 108–288 (stacked), the strips at block 216–306 start at y 60, hold four Latin characters above Y and continue at 117.5; a table Z at block 216 is at y 60 and the strips meeting Y and Z start at 117.5; a 60-row table first on its page splits in place after row 49; two rows of at least 800pt each take a page each, drawn to y 762; in two columns a 2-row table lays out as in one column, a too-wide table after 20 paragraphs stays at block 360 across the page edge, and a 30-row table (y 60–484) stays whole; a keepNext paragraph before a too-wide table moves to page 2 with it (after 19 and after 24 paragraphs), and before a 60-row table moves to page 2 alone, the table starting page 3. Fourth set (tbRl): after a 49-row table that leaves no room beside it, the first strip starts 7.6–7.8pt past the table\'s far edge at y 60 on an 18, 15 or 20pt line grid, without a grid, with a 200pt-wide table and with 14pt text; a 20- or 25-row table that does not fit below a 40-row table sits at y 60 with its near edge 14.76pt past that table\'s far edge (14.64 without a grid) on the same grids and widths and whether one or three paragraphs lie between, while those paragraphs keep consecutive strips; in two columns a 60-row table (847pt) after one paragraph or first on its page is drawn whole from the page top (y 0) and the text resumes 7.7pt past it; after a 60-row table first on its page, its last rows (y 60–215.8 on page 2) put a following table at y 215.8–244.6 one strip in and the strips meeting both at 244.6, and after twelve paragraphs a table at block 216 sits at y 60. Fifth set: the distance from the first table\'s far grid edge to the second table\'s near grid edge is 14.19pt without borders, 14.28–14.55pt with a 0.5pt border on only one of the two facing edges, 14.76pt with 0.5pt, 15.75pt with 1.5pt and 17.31pt with 3pt borders, that is 14.28pt between the outer edges of the facing rules; a 10pt table indent leaves both the 7.7pt and the 14.76pt results and the table\'s own position unchanged; in tbLrV the two baseline shapes give 7.7pt and 14.76pt mirrored; in two columns a 2-row table in column 1 does not affect the column-2 strips of its block range, while a 30-row table reaching y 484 makes them start at 484, and tables of 51, 54 and 57 rows (to y 780.6, 823.0 and 842.0) keep y 60, keep y 60, and are raised to y 36.7 so as to end at the page edge. Sixth set (tbRl, a table ending within one character of the band end, then paragraphs): the first strip starts past the table\'s far grid edge by 7.14pt without borders, 7.38pt with a 0.5pt border on the far edge only and the same on the near edge only, 7.74pt with 0.5pt, 8.70pt with 1.5pt and 10.14pt with 3pt borders, and 8.94pt with 3pt on either edge and 0.5pt on the other: the near rule\'s outer edge is at the block position and the strip starts 7.14pt past the far rule\'s outer edge, to within 0.1pt. With 2pt cell spacing the strip starts 7.33pt past the outer far rule as drawn, the same as without spacing, and the second of two tables 14.28pt past the first, the same as with 0.5, 1.5 and 3pt borders. Table alignment center or right leaves the table, the strips and both clearances where they are without it. Earlier sets state the two distances from grid edges with 0.5pt borders (7.7 and 14.76pt). Seventh set: with a default header and footer (tbRl and tbLrV) and with footnotes, whose area lies at the block end of the page, the table, the strips beside it and the 7.14pt resume are where they are without them; in tbLrV two columns reproduce the tbRl results mirrored (column-2 strips unaffected by a 2-row table, a 30-row table kept whole, tables of 57 and 60 rows raised to y 36.7 and y 0), as do a 60-row table after one paragraph and the tables and strips after a table\'s continued rows. Cell spacing widens the table in Word (192pt for the 180pt grid); that is the table\'s own layout and outside this rule. Eighth set: with 195.3pt of block extent left, a table with a 195pt grid stays on the page without borders and starts the next page with 0.5pt and with 3pt borders, so the test for a table too wide uses its box; a 60-row borderless table whose first-row first cell alone has a 3pt near rule is offset by half that rule on both of its pages; with 2pt cell spacing, no table borders and 6pt cell rules the table\'s grid edge is at the block position and the rules are drawn inside it; in two columns a 57-row table that meets a 2-row table is placed 14.28pt past it and then raised to end at the page edge, two tables of one 800pt row each with nothing between them run on as one block, the second directly below the first and the pair (1601pt, taller than the page) raised to the page top, and a keepNext paragraph before a table of one 800pt row, or before a table wider than the block extent left, stays where it is with the table. Ninth set (two columns; an 800pt-row table Y, a 20pt or 300pt-row table X, the paragraph holding a column break leaving its empty last strip at the top of column 2): with X in column 1 and Y in column 2, directly after the break or after a further paragraph, Y is placed 14.28pt past X along the block axis and raised to end at the page edge, in tbRl and mirrored in tbLrV, whatever the height of X, so a table meets the tables of every column of its page; with both in column 2 (X, a paragraph, Y) the same holds and the paragraphs beside X start at its bottom edge; with Y in column 1, raised over both columns, and X in column 2, the empty strip of column 2 has no room beside Y and resumes 7.14pt past it, X sits at the top of column 2 one strip further on, clear of Y, and the following paragraph starts at X\'s bottom edge. Tenth set (tbRl and tbLrV mirrored). Text before a raised table, two columns: twelve paragraphs in column 1, a column break and an 800pt-row table in column 2 put the table one strip into the flow, raised over both columns, and every strip of column 1 resumes 7.14pt past it, as the text of column 2 does; with two paragraphs before the table in column 2 the first two strips of column 1 keep their place and the third, which ends where the table starts, and the later ones resume past it; with 27 paragraphs and no break the table lies two strips into column 2, where the flow reaches it when the 27 are laid out without it (25 in column 1), and column 1 then holds one strip before it and thirteen past it and column 2 the rest in the same way. So a table is placed from the flow laid out without it and all text of the page then flows around it, also text of an earlier column. Tables that directly follow one another with no paragraph between them are laid out as one table: a 180pt-wide and a 90pt-wide table, in either order, lie one directly below the other with their left edges aligned (the far edge in tbRl), every strip beside either starting below both; in two columns such a pair below a 30-row table of column 1 that only the wider one reaches is placed below that table as a whole; a 400pt-row narrow and a 500pt-row wide table (900pt, taller than the page) start at the page top together; in one column a paragraph, 40 rows and 25 rows put all 65 rows on the next page, cut between rows at the band end inside the second table and continued on a third page, where a following 2-row table lies directly below the last rows and the next paragraph starts at its bottom edge; first on a page, 45 and 25 rows fill it and continue on the next. Other objects: a table after a paragraph that anchors a square-wrapped 80x40pt drawing at a page position over the table\'s place is placed directly below the drawing (within 0.5pt), the strip before the table keeping its text above it; a drawing anchored after the table is itself moved below the table, the table staying (not implemented: the drawing keeps its authored position); a table after a page-anchored floating table lying over its place is placed directly below that table (not implemented: the exclusion a floating table registers in a vertical section is not stated in the section\'s frame); a floating table directly after a block table is laid out as rows of that table (not implemented). Footnotes and overflow: a table whose first cell holds a footnote reference sits and excludes like any other, also after a 66-character strip. In two columns, 24 paragraphs and a 32-character strip filling column 1 before a 2-row table leave the table in column 1, too wide and overflowing, at the position the flow reached, while the 32-character paragraph, which beside the table would need a second strip that column 1 does not have, moves to the top of column 2 whole and the text continues there; the same at the end of column 2 puts the table at the start of the next page with the strip before it not wrapped, so a table wider than the block extent left starts the next page from the last column of a page and stays, overflowing, in an earlier one. Library policy where no control exists: a run of directly following tables whose border rules differ is aligned by the tables\' boxes; rows continued on a later page and owner segments stay at the flow position and clear the tables they meet along the inline axis only; a table placed ahead of the flow that the flow then reaches on a later page (its exclusion moved the text before it off the page) is placed where the flow reaches it and only the text after it flows around it; vertical rules painted outside the grid widen the box by what they overhang. A counter-clockwise section flow, which no Transitional token selects, keeps the library policy of an atomic table at the column bottom that advances the flow by its width. The other vertical tokens reach this rule through the page frame they share (btLr through word-section-btlr-tbrl-page-frame); a strip that fits its column only at full length is word-vertical-section-upright-table-rewrapped-strip.',
});

export const WORD_UPRIGHT_TABLE_REWRAPPED_STRIP = defineCompatibilityRule({
  id: 'word-vertical-section-upright-table-rewrapped-strip',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'vertical-section-upright-block-table',
    application: 'Microsoft Word',
    version: '16.0.20527',
    platform: 'Windows',
  },
  description: 'Under word-vertical-section-upright-block-table the line strip before an upright block table wraps around the table. Wrapping can push its last characters into further strips; if the table then followed the flow past them, the first strip would no longer meet it, so the rule alone has no layout. Word keeps the table at the block position the flow reaches when that strip is laid out without the table: the strip wraps, its overflow fills the following strips beside the table like any later line, and the flow continues after it. Evidence: Windows Word 16.0.20527 PDFs of the controls of that rule: a 66-character paragraph (one 693pt line in the 701.9pt column) before a 2-row table, in tbRl and tbLrV, holds 64 characters from y 88.7 to the column end and two in the second strip, with the table\'s near edge one strip into the flow; a 64-character paragraph after it before a second table behaves the same beside both tables (61 characters from y 117.5, three in the next strip, the second table one strip after the paragraph\'s start); a 30-character paragraph before a 45-row table (63pt of room) fills five strips of six characters beside the table, whose near edge stays one strip into the flow. Layout: a table is placed once and never moves. A paragraph fragment that ends its paragraph just before such a table is laid out without the table, which gives the block position the flow reaches; the table, and the tables that directly follow it, are placed there against the tables already on the page; the fragment is then laid out once more around them and the tables take those positions when the flow reaches them (body-paginator.ts, body-table-measurement.ts planUprightRun). When the fragment laid out around the tables no longer ends its paragraph in the same column, the tables keep their place and the paragraph is paginated beside them like any other (measured in the tenth set of that rule: it moves to the next column whole). A table that a multi-column section raises over a column laid out before its own is pinned and the page is laid out once more from its start with the pin registered, so that column\'s text flows around it too; each table is pinned at most once, so a page is laid out a bounded number of times.',
});

/** `word-vertical-section-upright-block-table`: how far past the outer edge
 * of an upright block table's far border rule a line strip first clears it,
 * and how far past it the outer edge of a later upright table's near border
 * rule does (measured distances; see the rule). */
export const WORD_UPRIGHT_TABLE_TEXT_CLEARANCE_PT = 7.14;
export const WORD_UPRIGHT_TABLE_TABLE_CLEARANCE_PT = 14.28;

export function wordExactRowFloorPt(
  authoredHeightPt: number | null,
  bottomCellMarginsPt: readonly number[],
): number {
  return Math.max(0, authoredHeightPt ?? 0)
    + Math.max(0, ...bottomCellMarginsPt);
}

/** Compatibility projection governed by
 * {@link WORD_COLLAPSED_BORDER_ROW_TRACK_FOOTPRINT}. */
export function wordCollapsedBorderRowTrackFootprintPt(
  exact: boolean,
  topRuleWidthPt: number,
  tableBottomRuleWidthPt: number,
): number {
  // An exact row's authored height already holds its top rule.
  return (exact ? 0 : Math.max(0, topRuleWidthPt)) + Math.max(0, tableBottomRuleWidthPt);
}

export function wordAuthoredBorderParticipates(
  authoredStyle: string | null | undefined,
): boolean {
  // `word-table-border-layer-cascade` owns this pre-conflict distinction.
  return authoredStyle !== null
    && authoredStyle !== undefined
    && authoredStyle !== 'none';
}

export function wordAlignedTableOriginPt(
  alignedPt: number,
  indentPt: number,
  bidiVisual: boolean,
): number {
  return bidiVisual ? alignedPt - indentPt : alignedPt + indentPt;
}

export function wordSpacedCellInsideBorderOverridesTable(input: Readonly<{
  spacingPt: number;
  directStyle: string | null | undefined;
  conditionalInsideStyle: string | null | undefined;
}>): boolean {
  return input.spacingPt > 0
    && !wordAuthoredBorderParticipates(input.directStyle)
    && wordAuthoredBorderParticipates(input.conditionalInsideStyle);
}

export function wordExactRowVerticalClipBounds(
  cellFlowBounds: LayoutRect,
  containingFlowBounds: LayoutRect,
): LayoutRect {
  // The exact trHeight constraint is block-axis-only. Keep the containing
  // flow's inline spill area, but extend it to any signed §17.4.50 tblInd
  // placement owned by this cell so a table shifted into the leading margin
  // is not clipped back to the ordinary text band.
  const xPt = Math.min(cellFlowBounds.xPt, containingFlowBounds.xPt);
  const rightPt = Math.max(
    cellFlowBounds.xPt + cellFlowBounds.widthPt,
    containingFlowBounds.xPt + containingFlowBounds.widthPt,
  );
  return Object.freeze({
    xPt,
    yPt: cellFlowBounds.yPt,
    widthPt: rightPt - xPt,
    heightPt: cellFlowBounds.heightPt,
  });
}

export function wordClipsOverPageCantSplitRow(input: Readonly<{
  compatibility: 'word' | 'standard';
  availableHeightPt: number;
  freshPageHeightPt: number;
  epsilonPt: number;
}>): boolean {
  return input.compatibility === 'word'
    && input.availableHeightPt + input.epsilonPt >= input.freshPageHeightPt;
}

/** Word observation under {@link WORD_AUTHORED_ROW_HEIGHT_PAGE_BOUNDARY}.
 * ECMA-376 §§17.4.6, 17.4.80 define cantSplit and row-height constraints but
 * do not prescribe this page-cut choice. Limit relocation to a first fragment
 * whose complete row fits a fresh page. The atLeast observation covered rows
 * whose content stayed below the authored minimum; content-expanded atLeast
 * rows and over-page rows are outside the tested boundary. */
export function wordRelocatesAuthoredHeightRowAtPageBoundary(input: Readonly<{
  compatibility: 'word' | 'standard';
  heightRule: 'auto' | 'atLeast' | 'exact';
  repeatedHeader: boolean;
  authoredHeightPt: number | null;
  availableHeightPt: number;
  wholeHeightPt: number;
  freshAvailableHeightPt: number;
  epsilonPt: number;
}>): boolean {
  return input.compatibility === 'word'
    && !input.repeatedHeader
    && (input.heightRule === 'exact' || input.heightRule === 'atLeast')
    && input.authoredHeightPt !== null
    && (input.heightRule === 'exact'
      || input.wholeHeightPt <= input.authoredHeightPt + input.epsilonPt)
    && input.authoredHeightPt > input.availableHeightPt + input.epsilonPt
    && input.wholeHeightPt <= input.freshAvailableHeightPt + input.epsilonPt;
}

/** Compatibility projection governed by {@link WORD_PARALLEL_PARAGRAPH_ROW_CUT}. */
export function wordRelocatesParallelParagraphRowCut(input: Readonly<{
  compatibility: 'word' | 'standard';
  hasUnfinishedParagraphWithoutProgress: boolean;
}>): boolean {
  return input.compatibility === 'word' && input.hasUnfinishedParagraphWithoutProgress;
}

export const WORD_TABLE_BORDER_STYLE_PRECEDENCE = Object.freeze([
  'single', 'thick', 'double', 'dotted', 'dashed', 'dotDash', 'dotDotDash', 'triple',
  'thinThickSmallGap', 'thickThinSmallGap', 'thinThickThinSmallGap', 'thinThickMediumGap',
  'thickThinMediumGap', 'thinThickThinMediumGap', 'thinThickLargeGap', 'thickThinLargeGap',
  'thinThickThinLargeGap', 'wave', 'doubleWave', 'dashSmallGap', 'dashDotStroked',
  'threeDEmboss', 'threeDEngrave', 'outset', 'inset',
] as const);

const WORD_BORDER_NUMBER: Readonly<Record<string, number>> = Object.freeze({
  single: 1,
  thick: 2,
  double: 3,
  dotDash: 8,
  dotDotDash: 9,
  triple: 10,
  thinThickSmallGap: 11,
  thickThinSmallGap: 12,
  thinThickThinSmallGap: 13,
  thinThickMediumGap: 14,
  thickThinMediumGap: 15,
  thinThickThinMediumGap: 16,
  thinThickLargeGap: 17,
  thickThinLargeGap: 18,
  thinThickThinLargeGap: 19,
  wave: 20,
  doubleWave: 21,
  dashSmallGap: 22,
  dashDotStroked: 23,
  threeDEmboss: 24,
  threeDEngrave: 25,
  outset: 26,
  inset: 27,
});

export function wordTableBorderWeight(
  style: string,
  widthPt: number,
): number {
  if (style === 'dotted' || style === 'dashed') return 1;
  return Math.max(0, widthPt) * 8 * (ownValue(WORD_BORDER_NUMBER, style) ?? 0);
}

export function wordTableRowHeightRule(
  normalizedRule: 'exact' | 'atLeast' | 'auto',
  authored: boolean,
): 'exact' | 'atLeast' | 'auto' {
  return authored ? normalizedRule : 'atLeast';
}

export function wordTableCellSpacingValuePt(
  kind: string,
  dxaValuePt: number | null,
): number | null {
  if (kind === 'pct' || kind === 'auto' || kind === 'nil') return 0;
  return dxaValuePt;
}

export function wordTableMarginValuePt(input: Readonly<{
  kind: string;
  dxaValuePt: number | null;
  scope: 'cell' | 'exception' | 'table' | 'style';
  edge: 'top' | 'bottom' | 'start' | 'end';
}>): number | null {
  if (input.kind === 'dxa') return input.dxaValuePt;
  if (input.scope === 'cell' || input.scope === 'exception') return null;
  if (input.edge === 'start' || input.edge === 'end') {
    if (input.kind === 'pct' || input.kind === 'auto' || input.kind === 'nil') return 0;
  }
  return null;
}

/** ECMA-376 §17.4.50 and WORD_TABLE_ORIGIN_COMPATIBILITY: measured
 * center/end controls ignore tblInd for both placement and width fitting.
 * Margin hanging is a separate placement correction, never a width indent.
 * Unmeasured classes retain the established width/placement contract. */
export function wordTableEffectiveIndentPt(input: Readonly<{
  measured: boolean;
  justification: string | null | undefined;
  indentPt: number;
}>): number {
  return input.measured && (input.justification === 'center'
    || input.justification === 'right' || input.justification === 'end') ? 0 : input.indentPt;
}

/** WORD_TABLE_ORIGIN_COMPATIBILITY. An unresolved/unmeasured class retains
 * its established translation rather than inventing a margin default. */
export function wordTableOriginTranslationPt(input: Readonly<{
  compat: Pick<DocxCompatProfile, 'tableMode15Geometry'>;
  measured: boolean;
  justification: string | null | undefined;
  indentPt: number;
  indentAuthored: boolean;
  firstLeftMarginPt: number;
  rowLeftMarginPt: number;
  /** WORD_TABLE_NONLEADING_INDENT_SCOPE outside the origin observation. */
  nonleadingIndentIgnored?: boolean;
}>): number {
  const indentPt = wordTableEffectiveIndentPt({ ...input,
    measured: input.measured || input.nonleadingIndentIgnored === true });
  if (!input.measured) return indentPt;
  if (input.justification === 'center') return indentPt;
  if (input.justification === 'right' || input.justification === 'end') {
    return indentPt + (input.compat.tableMode15Geometry ? 0 : input.rowLeftMarginPt);
  }
  return indentPt
    - (!input.compat.tableMode15Geometry && input.indentAuthored ? input.firstLeftMarginPt : 0);
}

/** Preserve grid indexes while dropping only the Office-measured unused track.
 * O(rows + columns + cells), with no parser mutation or extra content walk. */
export function wordFixedOccupiedGridInput(
  input: TableColumnLayoutInput,
  compat: Pick<DocxCompatProfile, 'tableOriginMeasuredMode'>,
  measuredScope: boolean,
): TableColumnLayoutInput {
  if (!compat.tableOriginMeasuredMode || !measuredScope
    || input.layout !== 'fixed' || input.rows.length === 0
    || !input.rows.every((row) => row.cells.length > 0
      && row.cells[0]?.columnStart === 1 && row.after === null
      && row.cells.every((cell) => cell.columnSpan === 1 && cell.preferredWidth?.kind === 'dxa' && cell.preferredWidth.value > 0))) {
    return input;
  }
  // A row occupying column zero is a measured counterexample. Only the single
  // universally skipped leading track was varied; spans, trailing skips,
  // missing preferences, and AutoFit retain their previous solver input.
  return {
    ...input,
    gridWidthsPt: input.gridWidthsPt.map((width, i) => i === 0 ? 0 : width),
    gridWidthKeys: input.gridWidthKeys?.map((key, i) => i === 0 ? null : key),
    rows: input.rows.map((row) => ({ ...row, before: null })),
  };
}

/*
 * Cell-owner row hosts. table-owner-runs.ts owns the structure (owner
 * contexts, SourceRef creation, partitioning, segment inputs, host geometry,
 * caching); the bounded compatibility policies it delegates to live here.
 *
 * Evidence boundary. ECMA-376 §17.3.1.11 defines framePr paragraph frames and
 * §17.4.57 makes a table without effective tblpPr nonfloating; neither
 * assigns a table row to a cell paragraph's frame, and MS-OI29500 2.1.43
 * defines no row promotion. The rules below are therefore office
 * observations, bounded to their control class; whatever they do not name is
 * library inference.
 *
 * Common control class (Word 16.113.3, macOS 27.0, public synthetic sources):
 * page/page notBeside carriers unless noted, a fixed 2×189pt grid with frame
 * w 378pt, horizontal text, no explicit row heights, no rotated, RTL or
 * tblpPr content. Original-open and saved/reopened full pages were exactly
 * equal, and the probes left the source and saved packages unchanged.
 * Numeric compatibility mode, exact fonts, explicit row heights, page splits,
 * tblpPr together with a carrier, other anchors and wrap modes, RTL and
 * rotated content were not observed; no universal Word behavior is claimed.
 */

/**
 * Context promotion. The 52-case context matrix varied the owning context of
 * one- and two-cell tables carrying a page/page notBeside framePr (and, where
 * tested, a width-omitted margin/text-around one). Word kept the framePr and
 * moved the row for direct roots of the body, header and footer stories. It
 * removed the framePr on save and rendered exactly the frame-free control for
 * tables nested in a body or footer cell (frame inside and beyond the flow
 * region), text box roots (unrotated, 90°, ±30°) and footnote/endnote roots.
 * The two-cell counterexamples rule out a cell-count rule and the
 * margin/text-around ones a page/page or notBeside rule: the owning context
 * decides. Contexts outside these classes are unobserved.
 */
export const WORD_CELL_OWNER_ROW_CONTEXT = defineCompatibilityRule({
  id: 'word-cell-owner-row-context',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'cell-owner-row-context-matrix',
    application: 'Microsoft Word',
    version: '16.113.3',
    platform: 'macOS 27.0',
  },
  description: 'A leading cell paragraph frame promotes its table row only for direct body, header and footer story roots; nested-cell, text box and note-root tables keep the frame-free result.',
});

/** {@link WORD_CELL_OWNER_ROW_CONTEXT}: `rootStory` is the story a table is a
 * direct root of, or null for a table owned by a cell at any depth. */
export function wordCellOwnerContextPromotesRows(rootStory: SourceRef['story'] | null): boolean {
  return rootStory === 'body' || rootStory === 'header' || rootStory === 'footer';
}

/**
 * Selector. Observed: a framePr on the first paragraph of the first cell moves
 * the whole row (borders and the ordinary sibling paragraphs of both cells)
 * to the frame position; the same framePr on the second paragraph renders
 * identically to the frame-free table. First-column vMerge controls (restart
 * carrier 144/90; continuation cell holding one empty paragraph): a
 * continuation whose pPr states an equal frame kept one two-row table, a
 * differing one (90/210) saved two one-row tables each at its own frame, and
 * no continuation frame saved two one-row tables with the later row in
 * ordinary flow. The continuation paragraph paints nothing, yet its own frame
 * takes part in row identity, and the restart frame is not inherited.
 *
 * Library choices: the logical first cell `cells[0]` is used, not the
 * visual-left cell of a bidiVisual table (RTL unobserved); explicit and
 * parser-defaulted dropCap `none` are one value, and `drop`, `margin` or
 * unknown tokens do not elect.
 */
export const WORD_CELL_OWNER_ROW_SELECTOR = defineCompatibilityRule({
  id: 'word-cell-owner-row-selector',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'cell-owner-row-selector-matrix',
    application: 'Microsoft Word',
    version: '16.113.3',
    platform: 'macOS 27.0',
  },
  description: 'A row elects the frame of its first cell\'s first paragraph, including a vertical-merge continuation\'s own stated frame; a frame on a later paragraph is inert, and an absent continuation frame is not inherited.',
});

/** {@link WORD_CELL_OWNER_ROW_SELECTOR}: the frame a row elects, or null. */
export function wordCellOwnerLeadingFrame(row: TableLayoutSource['rows'][number]): FramePr | null {
  const block = row.cells[0]?.content[0];
  if (block?.type !== 'paragraph' || !block.framePr || block.framePr.dropCap !== 'none') return null;
  return block.framePr as FramePr;
}

/**
 * Run identity. Observed: two rows with equal parsed carriers stay one
 * two-row table, the second row at its natural grid offset; rows whose
 * carriers differ in position, or only in hSpace with the same authored x/y,
 * render as independently placed hosts and save as two one-row tables, and
 * same-coordinate hosts overlap (no collision avoidance). With distinct top,
 * insideH and bottom border styles, differing carriers saved two one-row
 * tables each with the original top and bottom and no insideH, while equal
 * carriers kept top, insideH between the rows, and bottom; exact border
 * geometry was not compared. A second-column restart/continue across two rows
 * with the carrier on the first row only, on both rows at different
 * positions, or on the later row only saved two one-row tables keeping the
 * restart and an empty continue: merge markup does not decide ownership. An
 * equal-carrier merge was observed only in the first column (one two-row
 * table, merge kept; see {@link WORD_CELL_OWNER_ROW_SELECTOR}); a
 * second-column merge inside one run follows the library's grid model and
 * ECMA-376 §17.4.84, not an observation.
 */
export const WORD_CELL_OWNER_ROW_RUN_IDENTITY = defineCompatibilityRule({
  id: 'word-cell-owner-row-run-identity',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'cell-owner-row-identity-merge-matrix',
    application: 'Microsoft Word',
    version: '16.113.3',
    platform: 'macOS 27.0',
  },
  description: 'Adjacent electing rows with equal parsed frames stay one table; rows whose frames differ, even only in hSpace, become separate one-row hosts capped by the table top and bottom borders, regardless of vertical merges.',
});

/**
 * {@link WORD_CELL_OWNER_ROW_RUN_IDENTITY}: adjacent electing rows share one
 * run when these parsed framePr fields are equal. Only equality and the
 * hSpace difference were observed; the other fields are keyed by inference,
 * unprobed one by one. Absent w/h/x/y/xAlign/yAlign differ from any explicit
 * value, including 0 (library choice). dropCap is fixed by the selector and
 * `lines` only sizes drop caps. This is not §17.3.1.11 paragraph grouping and
 * not a claim about Word's native row-identity tuple.
 */
export function wordCellOwnerRunKey(framePr: FramePr): string {
  return stableFingerprint('w:framePr:owner-run', [
    framePr.hAnchor, framePr.vAnchor,
    framePr.x ?? null, framePr.xAlign ?? null,
    framePr.y ?? null, framePr.yAlign ?? null,
    framePr.w ?? null, framePr.h ?? null, framePr.hRule,
    framePr.wrap, framePr.hSpace, framePr.vSpace,
  ]);
}

/**
 * Host width clip. Observed: a 100pt frame over a 378pt grid clipped the
 * second cell's paint; a 378pt frame painted it. Which edge clips, the
 * absence of a vertical host clip and the retained text, search and
 * selection ownership of clipped cells are library choices made in
 * table-owner-runs.ts (finishOwnerHostLayout); RTL frames were not observed.
 */
export const WORD_CELL_OWNER_HOST_WIDTH_CLIP = defineCompatibilityRule({
  id: 'word-cell-owner-host-width-clip',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'cell-owner-host-frame-width-clip',
    application: 'Microsoft Word',
    version: '16.113.3',
    platform: 'macOS 27.0',
  },
  description: 'A row host whose explicit frame width is narrower than its natural grid keeps the grid and clips its paint at the frame width.',
});

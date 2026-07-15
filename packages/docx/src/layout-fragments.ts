/**
 * Measured layout fragments for the DOCX body flow (PR 5 of the layout-context /
 * fragments migration; see docs/docx-layout-context-fragments-design.md §"Measured
 * Fragment Model").
 *
 * A fragment belongs to a {@link DocumentLayout} result, NOT to the parsed document
 * model — the parsed {@link DocParagraph} is never mutated with layout state. A
 * A retained {@link ParagraphLayout} owns the exact line geometry for its recorded
 * placement. A paragraph split across pages or columns is represented by several
 * immutable layouts over the acquired paragraph geometry.
 *
 * All coordinates are points at scale 1 — the measurement coordinate system. Paint
 * scales the stored geometry; it never repeats text layout.
 *
 * Tables use the retained {@link TableLayout} tree, including page-local
 * {@link import('./layout/table-pagination.js').TableFragmentLayout} subtypes.
 */
import type { SectionGeom } from './types';
import type { SectionLayoutContext } from './layout-context.js';
import type { ParagraphLayout, TableLayout } from './layout/types.js';

export type FlowFragment = ParagraphLayout | TableLayout;

export interface PlacedFragment {
  readonly fragment: FlowFragment;
  /** ECMA-376 §17.6.4 newspaper column (0-based). */
  readonly columnIndex: number;
  /** Page-absolute point coordinates (scale 1). `yPt` is the top of the fragment's
   *  leading spacing; `heightPt` is the cursor advancement the paginator charged for
   *  this fragment (== leadingSpacePt + line advances + trailingSpacePt). */
  readonly xPt: number;
  readonly yPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
}

export interface BodyFragmentLayoutPage {
  readonly pageIndex: number;
  readonly section: SectionLayoutContext;
  readonly geometry: SectionGeom;
  readonly fragments: readonly PlacedFragment[];
}

export interface BodyFragmentDocumentLayout {
  readonly pages: readonly BodyFragmentLayoutPage[];
}

/**
 * Sum of the measured line advances this fragment paints, in scale-1 points.
 *
 * For a content fragment this is `measured.lines[i].advancePt` for the first line
 * in range plus, for each subsequent line, the (non-negative) inter-line gap to the
 * previous line's bottom followed by its advance — the same per-line extent the
 * paginator charges. For a markOnly fragment it is the paragraph-mark line box
 * height. It excludes the fragment's leading and trailing paragraph spacing (those
 * are `leadingSpacePt` / `trailingSpacePt`), so spacing is never double-counted.
 */
export function fragmentLineAdvancesPt(fragment: ParagraphLayout): number {
  if (fragment.lines.length === 0) {
    return fragment.paragraphMark?.bounds.heightPt ?? 0;
  }
  let sum = 0;
  for (let i = 0; i < fragment.lines.length; i++) {
    const line = fragment.lines[i];
    if (!line) break;
    if (i === 0) {
      sum += line.advancePt;
      continue;
    }
    const previous = fragment.lines[i - 1];
    const previousBottom = previous.bounds.yPt + previous.advancePt;
    sum += Math.max(0, line.bounds.yPt - previousBottom) + line.advancePt;
  }
  return sum;
}

/**
 * The cursor advancement a paragraph fragment charges: its leading spacing, plus the
 * measured line advances it paints, plus its trailing spacing. This is the invariant
 * a paginator's per-fragment height must equal — proving paragraph spacing is owned
 * by the fragment and added exactly once (design §"Pagination and paint invariants" 1).
 */
export function paragraphFragmentAdvancePt(fragment: ParagraphLayout): number {
  return fragment.advancePt;
}

/** Flow height (scale-1 pt) of any {@link FlowFragment}: a paragraph fragment's
 *  leading + line advances + trailing, or a table layout's retained advance. */
export function flowFragmentAdvancePt(fragment: FlowFragment): number {
  return fragment.kind === 'table'
    ? fragment.advancePt
    : paragraphFragmentAdvancePt(fragment);
}

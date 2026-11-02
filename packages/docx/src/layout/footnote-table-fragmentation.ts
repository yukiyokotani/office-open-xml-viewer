import type { RetainedTableAcquisition } from './table-acquisition.js';
import { startTableFragmentCursor, takeTableFragment, type TableFragmentCursor,
  type TableFragmentContext, type TableFragmentPageContext } from './table-pagination.js';
import type { LayoutServices, TableLayout } from './types.js';

// Acquisition ownership stays outside the immutable paint graph. Weak keys
// release the semantic table input with its shaped note epoch; no parser model
// or callback is stamped onto a retained layout.
const sources = new WeakMap<TableLayout, Readonly<{
  acquisition: RetainedTableAcquisition;
  services: LayoutServices;
  page: TableFragmentPageContext;
}>>();
const continuations = new WeakMap<TableFragmentCursor, NonNullable<ReturnType<typeof sources.get>>>();

export function registerFootnoteTable(
  layout: TableLayout, acquisition: RetainedTableAcquisition, services: LayoutServices,
  page: TableFragmentPageContext,
): void {
  const input = { ...acquisition.input, flowDomainId: layout.flowDomainId };
  sources.set(layout, { acquisition: { ...acquisition, input }, services, page });
}

/** Library content-retention policy, not an Office row-cut observation.
 * ECMA-376 §17.4.6/.49 govern row breaking and header repetition; reuse the
 * body paginator for row/cell-line/merge/border geometry and those policies.
 * A row cursor keeps its original shaped epoch:
 * reflowing cell lines at a new width or PAGE value would invalidate that cut.
 * Subsequent fragments use the destination's placement and page ownership;
 * column-width reflow inside a continued row is intentionally not claimed.
 * The minimum asks the body paginator for its first complete cell unit, with
 * full ink even when it exceeds capacity. Later cell units continue; retaining
 * an entire row suffix instead would strand those units below the page edge.
 * Page-dependent tables remain whole until they have a row-level page-frame
 * reacquisition contract. Their caller can still retain the complete block.
 */
export function partitionFootnoteTable(
  table: TableLayout, cursor: TableFragmentCursor | undefined, capacityPt: number,
  minimum: boolean,
): Readonly<{ fragment: TableLayout; nextCursor: TableFragmentCursor | null }> | null {
  const destination = sources.get(table);
  const source = (cursor && continuations.get(cursor)) || destination;
  if (!source) return null;
  const { acquisition, services } = source;
  const bounds = { ...table.flowBounds, heightPt: minimum ? 0 : capacityPt };
  const context: TableFragmentContext = {
    availableHeightPt: bounds.heightPt,
    // A fit trial may defer to another band; only the explicit minimum grants
    // the first-unit over-band fallback. These are capacity sentinels,
    // not fitted page/row dimensions.
    freshPageHeightPt: minimum ? 0 : Number.MAX_VALUE,
    ...(minimum ? { oversizedCellPolicy: 'overflow-first-unit' as const } : {}),
    compatibility: 'standard', services,
    placement: { container: { id: table.flowDomainId, kind: 'footnote', bounds },
      cursor: { xPt: bounds.xPt, yPt: bounds.yPt }, availableBounds: bounds },
    page: destination?.page ?? source.page,
  };
  const result = takeTableFragment(acquisition, cursor ?? startTableFragmentCursor(), context);
  if (result.nextCursor) continuations.set(result.nextCursor, source);
  return result.fragment ? { fragment: result.fragment, nextCursor: result.nextCursor } : null;
}

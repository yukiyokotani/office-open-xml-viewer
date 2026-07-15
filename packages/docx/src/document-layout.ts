/**
 * Body {@link DocumentLayout} production for the DOCX renderer (PR 5).
 *
 * `layoutDocument(doc)` returns the immutable fragment layout for the document body:
 * pages of {@link PlacedFragment}s over body paragraphs, each carrying the
 * placement-aware {@link MeasuredParagraph} the paginator measured. The measurement
 * and page-assignment engine lives in `renderer.ts` (it needs the renderer's private
 * measure state, section resolution and float registration); this module is the
 * public entry point for the fragment result and the natural home for future
 * body-flow composition as the migration proceeds (PR 6 adds table fragments).
 *
 * See docs/docx-layout-context-fragments-design.md §"Measured Fragment Model".
 */
export { layoutDocument } from './renderer.js';
export type {
  BodyFragmentDocumentLayout,
  BodyFragmentLayoutPage,
  PlacedFragment,
  FlowFragment,
} from './layout-fragments.js';
export type { ParagraphLayout } from './layout/types.js';

/**
 * Progressive document layout: paint the first pages long before the whole
 * document has been paginated.
 *
 * ## One resumable pagination session
 *
 * The canonical paginator already suspends between body entries. At selected
 * page-count checkpoints it now composes the committed page drafts through the
 * same layout-to-paint boundary used by the final document, publishes that
 * immutable snapshot, and then resumes the SAME generator state. No truncated
 * source, growing-prefix replay, or second paginator exists.
 *
 * Publications remain provisional because later anchor/header/footer/field
 * convergence can repaginate the document. They are nevertheless based on the
 * complete source input, so unbounded `keepNext` lookahead is no longer cut off
 * at an artificial preview boundary.
 *
 * ## What is guaranteed
 *
 * The final layout is produced by the ordinary full `paginateBody`, so it is
 * byte-identical to a blocking load. Publications only affect what is on screen
 * BEFORE that finishes.
 */
import type { BodyLayoutInput } from './body-layout-input.js';
import { paginateBodySteps } from './body-paginator.js';
import {
  drainPaginationAsync,
  type PaginationSchedulerOptions,
} from './pagination-scheduler.js';
import type { LayoutOptions } from './options.js';
import type { DocumentLayout, LayoutServices } from './types.js';

/** A document this small is likely to finish before a provisional paint helps. */
const MIN_PROGRESSIVE_ENTRIES = 12;

export interface ProgressiveLayoutOptions {
  /**
   * Receives each provisional prefix layout, in growing order.
   *
   * Called zero or more times before the returned promise settles. The first
   * call is the opening preview a caller can resolve `load()` on — delivered
   * asynchronously, since the preview is itself laid out in scheduler slices;
   * later calls extend it as the session progresses. Never called with fewer
   * pages than the previous call.
   */
  readonly onPreview?: (preview: ProgressiveLayoutPreview) => void;
  /** Scheduling for the full layout that follows the preview. */
  readonly scheduler?: PaginationSchedulerOptions;
}

export interface ProgressiveLayoutPreview {
  /** A complete, paintable layout of the document's opening pages. */
  readonly layout: DocumentLayout;
  /**
   * Whether these pages are known to match the final layout. Currently false:
   * a later convergence pass may still replace them.
   */
  readonly exact: boolean;
  /** Body entries the preview covers, for diagnostics. */
  readonly coveredEntries: number;
}

/**
 * Lay out and publish from one resumable canonical pagination session.
 *
 * The returned layout is the authoritative one; the preview is strictly a
 * stopgap for the viewport. When no useful preview can be produced — a document
 * short enough that previewing is pointless, or one whose first attempt yields
 * a single page — `onPreview` simply never fires and this degrades to an
 * ordinary sliced layout.
 */
export async function layoutDocumentProgressively(
  input: BodyLayoutInput,
  services: LayoutServices,
  options: LayoutOptions,
  progressive: ProgressiveLayoutOptions = {},
): Promise<DocumentLayout> {
  const { onPreview, scheduler } = progressive;
  const observer = onPreview && input.sequence.length > MIN_PROGRESSIVE_ENTRIES
    ? {
        onPages: (layout: DocumentLayout, processedEntries: number) => {
          onPreview(Object.freeze({
            layout,
            exact: false,
            coveredEntries: processedEntries,
          }));
        },
      }
    : undefined;
  return drainPaginationAsync(
    paginateBodySteps(input, services, options, observer),
    scheduler,
  );
}

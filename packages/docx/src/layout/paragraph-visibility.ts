import type { DocParagraph } from '../types.js';
import { runIsOmitted } from './revision-visibility.js';

/**
 * A paragraph whose mark is vanished (ECMA-376 §17.3.2.41 on the paragraph
 * mark run properties) and which has no ink joins the following paragraph, so
 * a flow that continues past it lays out no block for it. Retained layout
 * sources carry unavailable drawings as runs, which count as ink here.
 */
export function isSwallowedHiddenParagraph(
  paragraph: Readonly<{ markVanish?: boolean; runs?: readonly Readonly<{ type: string; text?: string }>[] }>,
): boolean {
  return paragraph.markVanish === true && !(paragraph.runs ?? []).some((candidate) => (
    candidate.type !== 'text' || (candidate.text?.length ?? 0) > 0));
}

export function isInklessParagraph(paragraph: Readonly<DocParagraph>): boolean {
  return !(paragraph.runs ?? []).some((candidate) => {
    if (candidate.type === 'text') return candidate.text.length > 0;
    return true;
  });
}

type AnchorOnlyRunFacts = Readonly<{
  type: string;
  text?: string;
  anchor?: boolean;
  inline?: boolean;
  listLabel?: unknown;
  __listLabel?: unknown;
  anchorAcquisitionInput?: unknown;
  __anchorAcquisition?: unknown;
  revision?: Readonly<{ kind?: string }> | null;
}>;

/** WORD_ANCHOR_HOST_NO_LINE_METRIC: the paragraph displays nothing but
 * floating drawings (their anchor hosts), so it is laid out and paginated as
 * its empty paragraph mark. Text hidden by revision display and unselected
 * optional hyphens are no content; a picture list label's metric run is. The
 * single classification for layout (paragraph-measure) and pagination. */
export function isAnchorOnlyParagraph(
  runs: readonly AnchorOnlyRunFacts[],
  showTrackedChanges: boolean | undefined,
): boolean {
  let anchored = false;
  for (const run of runs) {
    // A run the view omits emits nothing, whatever its type (the segment
    // builder's own predicate).
    if (runIsOmitted(run, showTrackedChanges)) continue;
    if (run.type === 'anchorHost') {
      if (run.listLabel !== undefined || run.__listLabel !== undefined) return false;
      anchored = true;
      continue;
    }
    if ((run.type === 'image' || run.type === 'chart') && run.anchor === true) continue;
    if (run.type === 'shape' && run.inline !== true) continue;
    if (run.type === 'unavailableDrawing'
      && (run.anchorAcquisitionInput !== undefined || run.__anchorAcquisition !== undefined)) continue;
    if (run.type === 'text' && (run.text ?? '').replace(/\u00AD/g, '') === '') continue;
    return false;
  }
  return anchored;
}

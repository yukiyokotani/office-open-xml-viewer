import type { DocParagraph } from '../types.js';
import { hasVisibleAnchoringRun, metricVisibleRuns } from './metric-visible-runs.js';

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

type MarkOnlyRunFacts = Readonly<{
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

/** WORD_ANCHOR_HOST_NO_LINE_METRIC: the paragraph displays nothing but its
 * mark in the given revision view: its metric-visible runs (metricVisibleRuns:
 * no floating drawing, no anchoring character, nothing the view omits) are
 * inkless exactly as isInklessParagraph defines for a source paragraph. Such
 * a paragraph is laid out and paginated as that empty paragraph, whether or
 * not it anchors floating drawings. A picture list label's metric run is
 * content. The single classification for layout (paragraph-measure) and
 * pagination. */
export function displaysOnlyParagraphMark(
  runs: readonly MarkOnlyRunFacts[],
  showTrackedChanges: boolean | undefined,
): boolean {
  return metricVisibleRuns(runs, showTrackedChanges).every((run) => run.type === 'text' && (run.text ?? '') === '');
}

/** A list marker (text or picture bullet) is painted: visible ink that no
 * spacing, collapsing or structural-mark role may remove. */
export function hasVisibleNumbering(paragraph: Readonly<{
  numbering?: Readonly<{ text?: string; picBulletImagePath?: string | null }> | null;
}>): boolean {
  const numbering = paragraph.numbering;
  return numbering != null && ((numbering.text ?? '') !== '' || Boolean(numbering.picBulletImagePath));
}

/** A paragraph that displays only its mark and anchors at least one floating
 * drawing the view displays. */
export function isAnchorOnlyParagraph(
  runs: readonly MarkOnlyRunFacts[],
  showTrackedChanges: boolean | undefined,
): boolean {
  return hasVisibleAnchoringRun(runs, showTrackedChanges) && displaysOnlyParagraphMark(runs, showTrackedChanges);
}

import type { DocParagraph } from '../types.js';

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

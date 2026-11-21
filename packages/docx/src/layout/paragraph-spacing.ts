import {
  WORD_FIXED_PARAGRAPH_AUTO_SPACING_STORED_MARGINS,
  WORD_HTML_PARAGRAPH_AUTO_SPACING,
} from './line-compatibility.js';
export interface ParagraphSpacingParticipant {
  readonly contextualSpacing?: boolean;
  readonly styleId?: string | null;
}

/**
 * ECMA-376 §17.3.1.33 paragraph gap, with the per-side
 * `w:contextualSpacing` suppression from §17.3.1.9.
 *
 * The returned value is the complete distance between the preceding line block
 * and the current line block. Callers use this one rule for ordinary flow,
 * table-cell block folds, and DrawingML/WPS text boxes.
 */
export function paragraphGapPt(
  previous: ParagraphSpacingParticipant | null,
  current: ParagraphSpacingParticipant,
  previousAfterPt: number,
  currentBeforePt: number,
): number {
  if (!previous) return currentBeforePt;
  const sameStyle = !!(previous.styleId && previous.styleId === current.styleId);
  const dropPrevious = !!(sameStyle && previous.contextualSpacing);
  const dropCurrent = !!(sameStyle && current.contextualSpacing);
  if (dropPrevious && dropCurrent) return 0;
  if (dropCurrent) return previousAfterPt;
  if (dropPrevious) return Math.max(currentBeforePt - previousAfterPt, 0);
  return Math.max(previousAfterPt, currentBeforePt);
}

/** Cursor adjustment expressed from the shared total-gap authority. */
export function paragraphGapAdjustment(
  previous: ParagraphSpacingParticipant | null,
  current: ParagraphSpacingParticipant,
  previousAfterPt: number,
  currentBeforePt: number,
): { readonly suppressBefore: boolean; readonly overlap: number } {
  const gapPt = paragraphGapPt(previous, current, previousAfterPt, currentBeforePt);
  const suppressBefore = gapPt <= previousAfterPt;
  return {
    suppressBefore,
    overlap: previousAfterPt + (suppressBefore ? 0 : currentBeforePt) - gapPt,
  };
}

/** Automatic (HTML) paragraph margin of {@link WORD_HTML_PARAGRAPH_AUTO_SPACING}. */
const WORD_HTML_AUTO_SPACING_PT = 14;

interface AutomaticSpacingParagraph {
  readonly type?: string;
  readonly spaceBefore?: number;
  readonly spaceAfter?: number;
  readonly beforeAutospacing?: boolean;
  readonly afterAutospacing?: boolean;
  readonly numbering?: Readonly<{ numId: number; abstractNumId?: number }> | null;
}

/**
 * Where a paragraph sits for {@link WORD_HTML_PARAGRAPH_AUTO_SPACING}: whether
 * it opens or closes a box whose edges drop automatic margins, and whether its
 * adjacent sibling paragraphs belong to the same list (abstract definition).
 * The default (no neighbourhood) is a standalone paragraph.
 */
export interface AutomaticSpacingNeighbourhood {
  readonly dropBefore: boolean;
  readonly dropAfter: boolean;
}

const STANDALONE: AutomaticSpacingNeighbourhood = Object.freeze({ dropBefore: false, dropAfter: false });

function sameList(
  a: AutomaticSpacingParagraph | undefined,
  b: AutomaticSpacingParagraph | undefined,
): boolean {
  const first = a?.type === undefined || a.type === 'paragraph' ? a?.numbering : null;
  const second = b?.type === undefined || b.type === 'paragraph' ? b?.numbering : null;
  if (!first || !second) return false;
  // One list is one abstract definition: numIds that share it continue one
  // counter (§17.9.15). A numId without an abstract reference is its own list.
  return first.abstractNumId !== undefined || second.abstractNumId !== undefined
    ? first.abstractNumId === second.abstractNumId
    : first.numId === second.numId;
}

/**
 * Neighbourhood of `blocks[index]` inside one sibling sequence (a story body,
 * a cell, a text box). `edges` says whether the sequence's start and end are
 * box edges that drop automatic margins: both edges of a table cell, and the
 * start of the document body, a header or a text box
 * (WORD_HTML_PARAGRAPH_AUTO_SPACING (2)). Only an
 * immediately adjacent paragraph counts as a list neighbour, so a table or a
 * hoisted page break in between separates the items.
 */
export function automaticSpacingNeighbourhood(
  blocks: readonly AutomaticSpacingParagraph[],
  index: number,
  edges: Readonly<{ start: boolean; end: boolean }>,
): AutomaticSpacingNeighbourhood {
  const current = blocks[index];
  const dropBefore = (edges.start && index === 0) || sameList(blocks[index - 1], current);
  const dropAfter = (edges.end && index === blocks.length - 1) || sameList(current, blocks[index + 1]);
  return dropBefore || dropAfter ? { dropBefore, dropAfter } : STANDALONE;
}

/**
 * ECMA-376 §17.3.1.33 gives an active automatic flag priority over the stored
 * before/after values (including line-unit values) and delegates the amount
 * to the consumer's HTML paragraph policy; Part 4 §14.8.3.15 repeats that
 * while doNotUseHTMLParagraphAutoSpacing is off. The amount and where it
 * applies are not specified (the Microsoft implementation notes on §17.3.1.33
 * cover only line-unit inheritance); {@link WORD_HTML_PARAGRAPH_AUTO_SPACING}
 * records the compatibility evidence: 14pt per automatic side whatever the
 * font size, dropped at the box edges and between same-list items that
 * `neighbourhood` reports. Explicit false, and every non-automatic side,
 * retains the stored value.
 *
 * Part 4 §14.8.3.15 specifies 5pt/10pt under the fixed compatibility setting.
 * WORD_FIXED_PARAGRAPH_AUTO_SPACING_STORED_MARGINS records the approved Word
 * deviation: that setting preserves the stored resolved pair (and no edge or
 * list suppression is observed or applied). Missing values remain zero;
 * line-unit interpretation stays at its existing parser owner.
 * Resolve at the immutable model boundary, before both pagination and line
 * acquisition consume the same pair; paint and paragraph-gap folding do not
 * need to know about automatic flags.
 */
export function resolveAutomaticParagraphMarginsPt(
  paragraph: Readonly<{
    spaceBefore?: number;
    spaceAfter?: number;
    beforeAutospacing?: boolean;
    afterAutospacing?: boolean;
  }>,
  fixed: boolean,
  neighbourhood: AutomaticSpacingNeighbourhood = STANDALONE,
): Readonly<{ spaceBefore: number; spaceAfter: number }> {
  void WORD_FIXED_PARAGRAPH_AUTO_SPACING_STORED_MARGINS;
  void WORD_HTML_PARAGRAPH_AUTO_SPACING;
  if (fixed) return { spaceBefore: paragraph.spaceBefore ?? 0, spaceAfter: paragraph.spaceAfter ?? 0 };
  return {
    spaceBefore: paragraph.beforeAutospacing === true
      ? (neighbourhood.dropBefore ? 0 : WORD_HTML_AUTO_SPACING_PT) : paragraph.spaceBefore ?? 0,
    spaceAfter: paragraph.afterAutospacing === true
      ? (neighbourhood.dropAfter ? 0 : WORD_HTML_AUTO_SPACING_PT) : paragraph.spaceAfter ?? 0,
  };
}

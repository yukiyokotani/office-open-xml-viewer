/**
 * Text highlights: a set of term groups an application keeps highlighted in a
 * viewer, independent of the interactive find (`findText`).
 *
 * `viewer.highlightText(groups)` replaces the whole set. Each group is searched
 * on its own by the on-demand term search (`term-search.ts`); groups do not
 * affect each other, so where matches of two groups overlap both are kept and
 * both are painted (in group order, the later group on top). Within one group
 * the longer match wins an overlap. The combined list is in document order,
 * and matches that start at the same place follow the group order.
 */
import type { FindMatch } from './find-match';

/** One group of terms highlighted together, for example all terms from one
 *  source. */
export interface TextHighlightGroup {
  /** Application-chosen identifier, unique within the set. Reported as
   *  {@link TextHighlightMatch.group} on every match of this group. */
  id: string;
  /** The texts to highlight. Each is matched literally, so a phrase is one term
   *  that contains spaces. Empty and repeated terms are ignored; terms are not
   *  trimmed. */
  terms: readonly string[];
  /** CSS background of this group's highlight boxes, applied verbatim like
   *  `findHighlightColors`. Use a translucent value so the text stays readable.
   *  Any CSS `background` value is accepted, so a pattern (for example a
   *  `repeating-linear-gradient`) can mark a group without relying on hue.
   *  XLSX also draws the cell border in this value when it is a plain colour.
   *  The value becomes part of an inline style, so pass only values the
   *  application controls. Omitted: a translucent green. */
  color?: string;
  /** Match case exactly. Default `false`. */
  caseSensitive?: boolean;
  /** Match whole words only; see `FindMatchesOptions.wholeWord` for the rule.
   *  Default `false`. */
  wholeWord?: boolean;
}

/** One highlighted occurrence, as returned by `highlightText`,
 *  `highlightNext` and `highlightPrev`. */
export interface TextHighlightMatch<Loc = unknown> extends FindMatch<Loc> {
  /** The {@link TextHighlightGroup.id} of the group that produced the match. */
  group: string;
  /** The term of that group that matched, as the application supplied it. */
  term: string;
}

/** Inline declarations that mark the active text highlight (the one
 *  `highlightNext` / `highlightPrev` last landed on). A ring instead of a
 *  colour change: the group background keeps its meaning, and the marker does
 *  not depend on hue. A dark ring with a light one inside it, visible on any
 *  page colour, drawn inside the box so that a clipping container (a PPTX
 *  shape frame) cannot cut it off. */
export const ACTIVE_TEXT_HIGHLIGHT_RING = 'box-shadow:inset 0 0 0 2px #111,inset 0 0 0 3px #fff;';

/**
 * Validate a set and return a copy that later changes to the caller's objects
 * cannot reach, without empty terms and without groups that have no term left.
 * An empty result therefore means "highlight nothing". Group ids must be
 * distinct, or the matches could not be attributed: a repeated id throws, also
 * in an otherwise empty set, before anything is cleared or searched.
 */
export function adoptTextHighlightGroups(groups: readonly TextHighlightGroup[]): TextHighlightGroup[] {
  if (new Set(groups.map((group) => group.id)).size !== groups.length) {
    throw new TypeError('duplicate group id');
  }
  return groups
    .map((group) => ({ ...group, terms: group.terms.filter((term) => term !== '') }))
    .filter((group) => group.terms.length);
}

/** Load the term-list matcher (several terms, whole words, group search). It
 *  is a separate chunk: the single-query find does not need it. */
export const loadTermSearch = (): Promise<typeof import('./term-search')> => import('./term-search');

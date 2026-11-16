/**
 * Term search: everything beyond the plain single-query find. Several terms at
 * once, whole-word boundaries, text-highlight groups, and the matchers the find
 * controllers run per unit (DOCX page, PPTX slide) or per cell (XLSX).
 *
 * Loaded on demand (`loadTermSearch`), because only `highlightText` and
 * `findText(..., { wholeWord: true })` need it; the plain find in
 * `text-index.ts` stays in the static bundle unchanged. To keep that module
 * out of a shared chunk, nothing is imported from it at run time: the callers
 * pass in its `foldPreservingLength` and `sliceRange`.
 */
import type { FindMatchesOptions, SearchRun, TextIndex } from './text-index';
import type { TextHighlightGroup } from './text-highlight';
import { wordBoundaryTest } from './word-boundary';

/** The searched text and its case-folded twin of equal length (a
 *  {@link TextIndex}, or one XLSX cell text). */
type SearchText = Pick<TextIndex, 'text' | 'folded'>;

/** `foldPreservingLength` of `text-index.ts`. */
export type Fold = (text: string) => string;

/** The whole-word test of one searched text: whether a match may start or end
 *  at an offset. */
export type WordBoundaryTest = (offset: number) => boolean;

/** Whether layout separates two runs that are adjacent in a unit's run list (a
 *  paragraph end, a line change, another text container), although the joined
 *  search text has no character between them. Supplied by the format. */
export type RunBreakTest<Run> = (previous: Run, next: Run) => boolean;

/** Background of a text-highlight group that sets no `color`. */
export const DEFAULT_TEXT_HIGHLIGHT = 'rgba(0, 184, 148, 0.35)';

/**
 * The whole-word test of an indexed text (see `word-boundary.ts` for the
 * rule). `runs` and `runBreak` add the layout breaks between the runs the
 * index was built from: a whole word may end at the last run of a paragraph
 * and start at the first run of the next.
 */
export function wordBoundaries<Run extends SearchRun>(
  index: Pick<TextIndex, 'text' | 'runStart'>,
  runs?: readonly Run[],
  runBreak?: RunBreakTest<Run>,
): WordBoundaryTest {
  let breaks: Set<number> | undefined;
  if (runs && runBreak) {
    let previous: Run | undefined;
    runs.forEach((run, i) => {
      if (run.text.length === 0) return;
      if (previous && runBreak(previous, run)) (breaks ??= new Set()).add(index.runStart[i]);
      previous = run;
    });
  }
  return wordBoundaryTest(index.text, breaks);
}

/** One match of a term list: its `[start, end)` range in the searched text and
 *  the index of the term that found it. */
export interface TermRange {
  start: number;
  end: number;
  termIndex: number;
}

/**
 * Visit every occurrence of `term` as a `[start, end)` range, in order and
 * without overlap, exactly as the plain find does: scanning resumes just past
 * each match. With a whole-word test, an occurrence that does not start and end
 * at a boundary is skipped, and scanning resumes one position later so it
 * cannot hide an occurrence that starts inside it.
 */
function forEachOccurrence(
  searched: SearchText,
  term: string,
  opts: FindMatchesOptions,
  fold: Fold,
  boundary: WordBoundaryTest | null,
  visit: (start: number, end: number) => void,
): void {
  if (term.length === 0) return;
  const haystack = opts.caseSensitive ? searched.text : searched.folded;
  const needle = opts.caseSensitive ? term : fold(term);
  let from = 0;
  for (;;) {
    const at = haystack.indexOf(needle, from);
    if (at === -1) break;
    const end = at + needle.length;
    if (boundary && !(boundary(at) && boundary(end))) {
      from = at + 1;
      continue;
    }
    visit(at, end);
    from = end;
  }
}

/**
 * Keep a non-overlapping subset of `ranges`, preferring the longer range where
 * two collide (so a phrase term wins over one of its own words searched
 * alongside it), then the earlier one, then the earlier term. Returned in
 * document order. Cost is linear in the total length of the candidate ranges.
 */
function dropOverlaps(ranges: TermRange[], textLength: number): TermRange[] {
  const covered = new Uint8Array(textLength);
  const byPreference = [...ranges].sort((a, b) => (b.end - b.start) - (a.end - a.start) || a.start - b.start);
  const kept: TermRange[] = [];
  for (const range of byPreference) {
    let free = true;
    for (let i = range.start; i < range.end; i++) {
      if (covered[i]) {
        free = false;
        break;
      }
    }
    if (!free) continue;
    covered.fill(1, range.start, range.end);
    kept.push(range);
  }
  return kept.sort((a, b) => a.start - b.start);
}

/**
 * Find every occurrence of several `terms` at once. Each term is searched like
 * a single find query and the results are merged into one document-ordered
 * list. Where matches of different terms overlap, the longer match is kept, so
 * every matched character belongs to exactly one match. Empty and repeated
 * terms are ignored (no trimming, as for a single query).
 *
 * `wordBoundary` builds the whole-word test of `searched`; it is called only
 * for `opts.wholeWord`.
 */
export function findTermRanges(
  searched: SearchText,
  terms: readonly string[],
  opts: FindMatchesOptions,
  fold: Fold,
  wordBoundary: () => WordBoundaryTest = () => wordBoundaryTest(searched.text),
): TermRange[] {
  const boundary = opts.wholeWord ? wordBoundary() : null;
  const ranges: TermRange[] = [];
  const searchedTerms = new Set<string>();
  terms.forEach((term, termIndex) => {
    if (searchedTerms.has(term)) return;
    searchedTerms.add(term);
    forEachOccurrence(searched, term, opts, fold, boundary, (start, end) => {
      ranges.push({ start, end, termIndex });
    });
  });
  // One term's matches are already disjoint and in order; only a merge of
  // several terms can collide.
  return searchedTerms.size <= 1 ? ranges : dropOverlaps(ranges, searched.text.length);
}

/** One text-highlight occurrence in one searched text: its range, its group,
 *  and the term that matched. */
export interface TextHighlightRange {
  start: number;
  end: number;
  group: TextHighlightGroup;
  term: string;
}

/** Search every group in one text. Document order; occurrences that start at
 *  the same offset follow the group order. `wordBoundary` is called at most
 *  once, and only when a group asks for whole words. */
export function findTextHighlightRanges(
  searched: SearchText,
  groups: readonly TextHighlightGroup[],
  fold: Fold,
  wordBoundary: () => WordBoundaryTest = () => wordBoundaryTest(searched.text),
): TextHighlightRange[] {
  const found: TextHighlightRange[] = [];
  let boundary: WordBoundaryTest | undefined;
  const sharedBoundary = (): WordBoundaryTest => boundary ??= wordBoundary();
  for (const group of groups) {
    for (const { start, end, termIndex } of findTermRanges(searched, group.terms, group, fold, sharedBoundary)) {
      found.push({ start, end, group, term: group.terms[termIndex] });
    }
  }
  // Array.prototype.sort is stable, so equal starts keep the group order.
  return groups.length > 1 ? found.sort((a, b) => a.start - b.start) : found;
}

/** What a text-highlight match carries besides its position: the group's
 *  background, and the `{ group, term }` reported on the public match. */
interface HighlightSource {
  color: string;
  source: { group: string; term: string };
}

const highlightSource = ({ group, term }: TextHighlightRange): HighlightSource => ({
  color: group.color ?? DEFAULT_TEXT_HIGHLIGHT,
  source: { group: group.id, term },
});

/** Reject a set the matches of which could not be attributed: group ids must
 *  be distinct. */
function assertUniqueGroupIds(groups: readonly TextHighlightGroup[]): void {
  if (new Set(groups.map((group) => group.id)).size !== groups.length) {
    throw new TypeError('highlightText: group ids must be unique');
  }
}

/** Resolves one unit's index and runs to its matches. */
type UnitMatcher<Run, Hit> = (index: TextIndex, runs: readonly Run[]) => Hit[];
type SliceRange<Slice> = (index: TextIndex, start: number, end: number) => Slice[];
type RunBreakLoader<Run> = (() => Promise<RunBreakTest<Run>>) | undefined;

/**
 * The per-unit matcher of `findText(query, { wholeWord: true })`. `fold` and
 * `slice` are the static index primitives; `loadRunBreak` fetches the format's
 * layout-break test.
 */
export async function wholeWordMatcher<Run extends SearchRun, Slice>(
  query: string,
  opts: FindMatchesOptions,
  fold: Fold,
  slice: SliceRange<Slice>,
  loadRunBreak: RunBreakLoader<Run>,
): Promise<UnitMatcher<Run, { slices: Slice[] }>> {
  const runBreak = await loadRunBreak?.();
  return (index, runs) =>
    findTermRanges(index, [query], opts, fold, () => wordBoundaries(index, runs, runBreak))
      .map(({ start, end }) => ({ slices: slice(index, start, end) }));
}

/**
 * The per-unit matcher of a text-highlight set (rejects one with a repeated
 * group id). The layout-break test is fetched only when a group asks for whole
 * words.
 */
export async function textHighlightMatcher<Run extends SearchRun, Slice>(
  groups: readonly TextHighlightGroup[],
  fold: Fold,
  slice: SliceRange<Slice>,
  loadRunBreak: RunBreakLoader<Run>,
): Promise<UnitMatcher<Run, HighlightSource & { slices: Slice[] }>> {
  assertUniqueGroupIds(groups);
  const runBreak = groups.some((group) => group.wholeWord) ? await loadRunBreak?.() : undefined;
  return (index, runs) =>
    findTextHighlightRanges(index, groups, fold, () => wordBoundaries(index, runs, runBreak))
      .map((hit) => ({ slices: slice(index, hit.start, hit.end), ...highlightSource(hit) }));
}

/** Resolves one cell's display texts (the full one first) to its matches. A
 *  cell is a search unit of its own, so its ends are word boundaries. */
type CellMatcher<Hit> = (candidates: readonly string[]) => Hit[];

/** The per-cell matcher of `findText(query, { wholeWord: true })` on a sheet:
 *  the matches come from the first display text that has any. */
export function wholeWordCellMatcher(
  query: string,
  opts: FindMatchesOptions,
  fold: Fold,
): CellMatcher<{ text: string }> {
  return (candidates) => {
    for (const text of candidates) {
      const ranges = findTermRanges({ text, folded: fold(text) }, [query], opts, fold);
      if (ranges.length > 0) return ranges.map(({ start, end }) => ({ text: text.slice(start, end) }));
    }
    return [];
  };
}

/** The per-cell matcher of a text-highlight set (throws for one with a repeated
 *  group id). Each group is searched on its own; like the find, a group's
 *  matches in a cell come from the first of the cell's display texts that has
 *  any. */
export function textHighlightCellMatcher(
  groups: readonly TextHighlightGroup[],
  fold: Fold,
): CellMatcher<HighlightSource & { text: string }> {
  assertUniqueGroupIds(groups);
  return (candidates) => {
    const searched = candidates.map((text) => ({ text, folded: fold(text) }));
    const found: (HighlightSource & { text: string; start: number })[] = [];
    for (const group of groups) {
      for (const candidate of searched) {
        const hits = findTextHighlightRanges(candidate, [group], fold);
        for (const hit of hits) {
          found.push({ text: candidate.text.slice(hit.start, hit.end), start: hit.start, ...highlightSource(hit) });
        }
        if (hits.length > 0) break;
      }
    }
    // Stable: occurrences that start at the same offset keep the group order.
    return found.sort((a, b) => a.start - b.start).map(({ start: _start, ...hit }) => hit);
  };
}

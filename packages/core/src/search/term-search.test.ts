import { describe, expect, it } from 'vitest';
import {
  buildTextIndex, findMatches, foldPreservingLength as fold, sliceRange,
  type FindMatchesOptions, type MatchRunSlice, type SearchRun, type TextIndex,
} from './text-index.js';
import {
  findTermRanges,
  findTextHighlightRanges,
  matchCellGroups,
  matchGroups,
  matchCellWords,
  matchWords,
  wordBoundaries,
} from './term-search.js';
import type { TextHighlightGroup } from './text-highlight.js';

/**
 * The on-demand term search: several terms at once, whole words, and
 * text-highlight groups. The static index supplies the case fold and the
 * run-slice resolution, as in the find controllers.
 */
const runs = (...texts: string[]): SearchRun[] => texts.map((text) => ({ text }));
const index = (...texts: string[]): TextIndex => buildTextIndex(runs(...texts));
const terms = (searched: TextIndex, list: readonly string[], opts: FindMatchesOptions = {}) =>
  findTermRanges(searched, list, opts, fold);
const starts = (ranges: { start: number }[]) => ranges.map((range) => range.start);
const wholeWords = (searched: TextIndex, query: string) => starts(terms(searched, [query], { wholeWord: true }));

describe('findTermRanges — several terms', () => {
  it('gives a one-element list the matches of the plain find', () => {
    const searched = index('the quick brown fox, quick');
    expect(terms(searched, ['quick']).map(({ start, end }) => sliceRange(searched, start, end)))
      .toEqual(findMatches(searched, 'quick').map((match) => match.slices));
    expect(starts(terms(index('aaaa'), ['aa']))).toEqual([0, 2]);
  });

  it('merges every term into one document-ordered list', () => {
    const found = terms(index('fox and dog and fox'), ['dog', 'fox']);
    expect(starts(found)).toEqual([0, 8, 16]);
    expect(found.map((range) => range.termIndex)).toEqual([1, 0, 1]);
  });

  it('matches a spaced phrase contiguously, across runs', () => {
    const searched = index('quarterly ', 'earnings call; earnings');
    const found = terms(searched, ['quarterly earnings']);
    expect(found).toHaveLength(1);
    expect(sliceRange(searched, found[0].start, found[0].end)).toEqual([
      { runIndex: 0, start: 0, end: 10 },
      { runIndex: 1, start: 0, end: 8 },
    ]);
  });

  it('keeps the longer match where terms overlap, so no character is matched twice', () => {
    const found = terms(index('earnings call and earnings'), ['earnings', 'earnings call']);
    expect(found).toEqual([
      { start: 0, end: 13, termIndex: 1 },
      { start: 18, end: 26, termIndex: 0 },
    ]);
  });

  it('keeps an occurrence that only overlaps a losing candidate of its own term', () => {
    // "returns returns" occurs at 11 and at 19. The one at 11 loses to the
    // longer "guaranteed returns"; the one at 19 overlaps neither winner.
    const searched = index('guaranteed returns returns returns');
    expect(terms(searched, ['guaranteed returns', 'returns returns'], { wholeWord: true })).toEqual([
      { start: 0, end: 18, termIndex: 0 },
      { start: 19, end: 34, termIndex: 1 },
    ]);
    // Within one term, and between equal candidates, the earlier one still wins.
    expect(starts(terms(index('aaaaa'), ['aa', 'aaa']))).toEqual([0, 3]);
    expect(starts(terms(index('aaaa'), ['aa', 'b']))).toEqual([0, 2]);
    expect(starts(terms(index('aaaa'), ['aa']))).toEqual([0, 2]);
  });

  it('folds case for every term by default, and honours caseSensitive', () => {
    const searched = index('Alpha beta GAMMA');
    expect(starts(terms(searched, ['alpha', 'gamma']))).toEqual([0, 11]);
    expect(starts(terms(searched, ['alpha', 'GAMMA'], { caseSensitive: true }))).toEqual([11]);
  });

  it('uses the same length-preserving fold as the index', () => {
    // U+0130 lowercases to two code units; the index keeps it unfolded.
    expect(starts(terms(index('xİy'), ['İy']))).toEqual([1]);
  });

  it('collapses terms that fold to the same text into one match of the first term', () => {
    const found = terms(index('Report'), ['report', 'REPORT', 'report']);
    expect(found).toEqual([{ start: 0, end: 6, termIndex: 0 }]);
  });

  it('searches for nothing when every term is empty', () => {
    expect(terms(index('anything'), [])).toEqual([]);
    expect(terms(index('anything'), ['', ''])).toEqual([]);
  });

  it('applies wholeWord to every term', () => {
    const searched = index('concatenate the cat and dogma dog');
    expect(starts(terms(searched, ['cat', 'dog'], { wholeWord: true }))).toEqual([16, 30]);
  });
});

describe('findTermRanges — wholeWord', () => {
  it('rejects a match inside a longer word and keeps the standalone one', () => {
    const searched = index('concatenate the cat');
    expect(wholeWords(searched, 'cat')).toEqual([16]);
    expect(starts(terms(searched, ['cat']))).toEqual([3, 16]);
  });

  it('accepts a word bounded by punctuation or the ends of the text', () => {
    expect(wholeWords(index('cat, (cat) cat'), 'cat')).toEqual([0, 6, 11]);
  });

  it('only demands a boundary where the term itself starts or ends with a word character', () => {
    // "C++" ends in "+", so a following "," or "x" does not join it to a word.
    expect(wholeWords(index('use C++, not C++x'), 'C++')).toEqual([4, 13]);
    // Likewise at the start: "#tag" begins with "#", so the "x" before it does not join it.
    expect(wholeWords(index('see x#tag'), '#tag')).toEqual([5]);
  });

  it('does not let a rejected occurrence hide a valid one that starts inside it', () => {
    // "a a" at 1 is glued to "b"; the whole-word "a a" at 3 overlaps it.
    expect(wholeWords(index('ba a a'), 'a a')).toEqual([3]);
  });

  it('treats non-ASCII letters, including astral ones, as word characters', () => {
    expect(wholeWords(index('cafécat'), 'cat')).toEqual([]);
    expect(wholeWords(index('\u{1D400}cat'), 'cat')).toEqual([]);
  });

  it('joins runs that continue a word when no layout break separates them', () => {
    expect(wholeWords(index('bob', 'cat cat'), 'cat')).toEqual([7]);
  });

  it('accepts a word that ends or starts at a layout break between runs', () => {
    const paragraphs = [{ text: 'Introduction', p: 0 }, { text: '', p: 1 }, { text: 'The intro', p: 1 }];
    const searched = buildTextIndex(paragraphs);
    const boundary = () => wordBoundaries(searched, paragraphs, (previous, next) => previous.p !== next.p);
    const at = (query: string) => starts(findTermRanges(searched, [query], { wholeWord: true }, fold, boundary));
    expect(at('introduction')).toEqual([0]);
    expect(at('the')).toEqual([12]);
    expect(at('intro')).toEqual([16]);
    // Without the break the same seam joins "Introduction" and "The" into one word.
    expect(wholeWords(searched, 'the')).toEqual([]);
  });
});

describe('findTextHighlightRanges', () => {
  const summary = (hits: ReturnType<typeof findTextHighlightRanges>) =>
    hits.map((hit) => [hit.start, hit.end, hit.group.id, hit.term]);
  const search = (text: string, groups: readonly TextHighlightGroup[]) =>
    summary(findTextHighlightRanges(index(text), groups, fold));

  it('searches each group on its own and merges the hits in document order', () => {
    expect(search('revenue and guaranteed returns, revenue', [
      { id: 'user', terms: ['revenue'] },
      { id: 'validated', terms: ['guaranteed returns'] },
    ])).toEqual([
      [0, 7, 'user', 'revenue'],
      [12, 30, 'validated', 'guaranteed returns'],
      [32, 39, 'user', 'revenue'],
    ]);
  });

  it('keeps overlapping hits of different groups, in group order at the same start', () => {
    expect(search('guaranteed returns', [
      { id: 'user', terms: ['returns', 'guaranteed'] },
      { id: 'validated', terms: ['guaranteed returns'] },
    ])).toEqual([
      [0, 10, 'user', 'guaranteed'],
      [0, 18, 'validated', 'guaranteed returns'],
      [11, 18, 'user', 'returns'],
    ]);
  });

  it('lets the longer term win an overlap inside one group', () => {
    expect(search('earnings call and earnings', [{ id: 'a', terms: ['earnings', 'earnings call'] }]))
      .toEqual([[0, 13, 'a', 'earnings call'], [18, 26, 'a', 'earnings']]);
  });

  it('applies caseSensitive and wholeWord per group', () => {
    expect(search('Cat concatenate cat', [
      { id: 'loose', terms: ['cat'] },
      { id: 'word', terms: ['cat'], wholeWord: true },
      { id: 'exact', terms: ['Cat'], caseSensitive: true },
    ])).toEqual([
      [0, 3, 'loose', 'cat'], [0, 3, 'word', 'cat'], [0, 3, 'exact', 'Cat'],
      [7, 10, 'loose', 'cat'],
      [16, 19, 'loose', 'cat'], [16, 19, 'word', 'cat'],
    ]);
  });

  it('builds the whole-word test once, and only when a group asks for it', () => {
    const searched = index('cat cat');
    let built = 0;
    const boundary = () => { built++; return wordBoundaries(searched); };
    findTextHighlightRanges(searched, [{ id: 'a', terms: ['cat'] }], fold, boundary);
    expect(built).toBe(0);
    findTextHighlightRanges(searched, [
      { id: 'a', terms: ['cat'], wholeWord: true }, { id: 'b', terms: ['cat'], wholeWord: true },
    ], fold, boundary);
    expect(built).toBe(1);
  });

  it('reports the term as supplied, not as it appears in the text', () => {
    expect(search('REVENUE', [{ id: 'a', terms: ['Revenue'] }])).toEqual([[0, 7, 'a', 'Revenue']]);
  });

  it('ignores empty terms and groups without terms', () => {
    expect(search('abc', [{ id: 'a', terms: [''] }, { id: 'b', terms: [] }])).toEqual([]);
  });
});

describe('unit matchers', () => {
  const paragraphs = [{ text: 'Introduction', p: 0 }, { text: 'The cat', p: 1 }];
  type Paragraph = (typeof paragraphs)[number];
  const loadBreak = async () => ({ default: (previous: Paragraph, next: Paragraph) => previous.p !== next.p });
  const accept = (): boolean => true;
  /** The format rule of PPTX table cells: a match stays within one container. */
  const withinParagraph = (unitRuns: Paragraph[], slices: MatchRunSlice[]): boolean =>
    slices.every((slice) => unitRuns[slice.runIndex].p === unitRuns[slices[0].runIndex].p);

  it('matchWords finds whole words using the format\'s layout breaks', async () => {
    const searched = buildTextIndex(paragraphs);
    const withBreaks = await matchWords<Paragraph, MatchRunSlice>(
      'the', { wholeWord: true }, fold, sliceRange, loadBreak, accept);
    expect(withBreaks(searched, paragraphs)).toEqual([{ slices: [{ runIndex: 1, start: 0, end: 3 }] }]);
    const without = await matchWords<Paragraph, MatchRunSlice>(
      'the', { wholeWord: true }, fold, sliceRange, undefined, accept);
    expect(without(searched, paragraphs)).toEqual([]);
  });

  it('matchGroups resolves slices, colour and source, loading breaks only for whole words', async () => {
    const searched = buildTextIndex(paragraphs);
    let loads = 0;
    const counted = () => { loads++; return loadBreak(); };
    const plain = await matchGroups<Paragraph, MatchRunSlice>(
      [{ id: 'user', terms: ['cat'], color: 'red' }, { id: 'system', terms: ['the'] }], fold, sliceRange, counted, accept);
    expect(loads).toBe(0);
    expect(plain(searched, paragraphs)).toEqual([
      { slices: [{ runIndex: 1, start: 0, end: 3 }], color: 'rgba(0, 184, 148, 0.35)', layer: 1, source: { group: 'system', term: 'the' } },
      { slices: [{ runIndex: 1, start: 4, end: 7 }], color: 'red', layer: 0, source: { group: 'user', term: 'cat' } },
    ]);
    const words = await matchGroups<Paragraph, MatchRunSlice>(
      [{ id: 'system', terms: ['introduction'], wholeWord: true }], fold, sliceRange, counted, accept);
    expect(loads).toBe(1);
    expect(words(searched, paragraphs)).toHaveLength(1);
  });

  it('applies the format rule before the overlap rule, so a rejected long match keeps a shorter valid one', async () => {
    const cells = [{ text: 'cat', p: 0 }, { text: 'dog cat', p: 1 }];
    const searched = buildTextIndex(cells);
    const highlights = await matchGroups<Paragraph, MatchRunSlice>(
      [{ id: 'a', terms: ['catdog', 'cat'] }], fold, sliceRange, undefined, withinParagraph);
    expect(highlights(searched, cells).map((hit) => hit.slices)).toEqual([
      [{ runIndex: 0, start: 0, end: 3 }],
      [{ runIndex: 1, start: 4, end: 7 }],
    ]);
    // A rejected occurrence does not hide one that starts inside it either.
    const seam = [{ text: 'xa', p: 0 }, { text: 'aa', p: 1 }];
    const words = await matchWords<Paragraph, MatchRunSlice>(
      'aa', { wholeWord: true }, fold, sliceRange, loadBreak, withinParagraph);
    expect(words(buildTextIndex(seam), seam)).toEqual([{ slices: [{ runIndex: 1, start: 0, end: 2 }] }]);
  });
});

describe('cell matchers', () => {
  it('matchCellWords takes the first display text that has a whole-word match', () => {
    const match = matchCellWords('cat', { wholeWord: true }, fold);
    expect(match(['category Cat'])).toEqual([{ text: 'Cat' }]);
    expect(match(['cats', 'a cat'])).toEqual([{ text: 'cat' }]);
    expect(match(['cats'])).toEqual([]);
  });

  it('matchCellGroups searches each group in the first display text that matches it', () => {
    const match = matchCellGroups([
      { id: 'full', terms: ['456'], color: 'red' },
      { id: 'short', terms: ['235'] },
    ], fold);
    expect(match(['0.123456789', '0.1235'])).toEqual([
      { text: '235', color: 'rgba(0, 184, 148, 0.35)', layer: 1, source: { group: 'short', term: '235' } },
      { text: '456', color: 'red', layer: 0, source: { group: 'full', term: '456' } },
    ]);
  });
});

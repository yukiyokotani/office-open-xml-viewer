import { describe, it, expect } from 'vitest';
import { buildTextIndex, findMatches, normalizeFindQuery, type SearchRun } from './text-index.js';

/**
 * IX2 core search index. `buildTextIndex` joins an ordered run list; `findMatches`
 * locates a query in the joined text and maps each hit back to the run-slices it
 * covers so a viewer can highlight the right glyphs — including a match that
 * straddles a run boundary (the sole subtle case). Case-insensitive by default.
 */
const runs = (...texts: string[]): SearchRun[] => texts.map((text) => ({ text }));

describe('buildTextIndex', () => {
  it('joins runs in order and records each run start offset', () => {
    const index = buildTextIndex(runs('Hello ', 'World'));
    expect(index.text).toBe('Hello World');
    expect(index.folded).toBe('hello world');
    expect(index.runStart).toEqual([0, 6]);
    expect(index.runCount).toBe(2);
  });

  it('handles an empty run list', () => {
    const index = buildTextIndex([]);
    expect(index.text).toBe('');
    expect(index.runCount).toBe(0);
    expect(index.runStart).toEqual([]);
  });
});

describe('findMatches — single-run matches', () => {
  it('finds a substring wholly inside one run', () => {
    const index = buildTextIndex(runs('the quick brown fox'));
    const m = findMatches(index, 'quick');
    expect(m).toHaveLength(1);
    expect(m[0].matchIndex).toBe(0);
    expect(m[0].slices).toEqual([{ runIndex: 0, start: 4, end: 9 }]);
  });

  it('returns non-overlapping matches in document order with ascending matchIndex', () => {
    const index = buildTextIndex(runs('abababab'));
    const m = findMatches(index, 'ab');
    expect(m.map((x) => x.matchIndex)).toEqual([0, 1, 2, 3]);
    expect(m.map((x) => x.slices[0].start)).toEqual([0, 2, 4, 6]);
  });

  it('advances past each match so overlapping occurrences are not double-counted', () => {
    const index = buildTextIndex(runs('aaaa'));
    const m = findMatches(index, 'aa');
    // Browser find-in-page semantics: matches at 0 and 2 (not 0,1,2,3).
    expect(m.map((x) => x.slices[0].start)).toEqual([0, 2]);
  });
});

describe('findMatches — case sensitivity', () => {
  it('is case-insensitive by default', () => {
    const index = buildTextIndex(runs('Hello HELLO hello'));
    const m = findMatches(index, 'hello');
    expect(m).toHaveLength(3);
  });

  it('honors caseSensitive: true', () => {
    const index = buildTextIndex(runs('Hello HELLO hello'));
    const m = findMatches(index, 'hello', { caseSensitive: true });
    expect(m).toHaveLength(1);
    expect(m[0].slices[0].start).toBe(12);
  });
});

describe('findMatches — cross-run matches', () => {
  it('resolves a match split across two runs into two slices', () => {
    // "Hello" split as "Hel" | "lo".
    const index = buildTextIndex(runs('Hel', 'lo World'));
    const m = findMatches(index, 'Hello');
    expect(m).toHaveLength(1);
    expect(m[0].slices).toEqual([
      { runIndex: 0, start: 0, end: 3 },
      { runIndex: 1, start: 0, end: 2 },
    ]);
  });

  it('resolves a match spanning three runs (whole middle run)', () => {
    // "abcdef" as "ab" | "cd" | "ef".
    const index = buildTextIndex(runs('ab', 'cd', 'ef'));
    const m = findMatches(index, 'bcde');
    expect(m).toHaveLength(1);
    expect(m[0].slices).toEqual([
      { runIndex: 0, start: 1, end: 2 },
      { runIndex: 1, start: 0, end: 2 },
      { runIndex: 2, start: 0, end: 1 },
    ]);
  });

  it('skips zero-length runs that fall inside a match range', () => {
    // Empty run between two text runs must not produce an empty slice.
    const index = buildTextIndex(runs('ab', '', 'cd'));
    const m = findMatches(index, 'abcd');
    expect(m).toHaveLength(1);
    expect(m[0].slices).toEqual([
      { runIndex: 0, start: 0, end: 2 },
      { runIndex: 2, start: 0, end: 2 },
    ]);
  });

  it('matches case-insensitively across a run boundary', () => {
    const index = buildTextIndex(runs('FIND', 'text'));
    const m = findMatches(index, 'dt');
    expect(m).toHaveLength(1);
    expect(m[0].slices).toEqual([
      { runIndex: 0, start: 3, end: 4 },
      { runIndex: 1, start: 0, end: 1 },
    ]);
  });
});

describe('findMatches — length-changing case folds must not shift offsets', () => {
  // U+0130 (İ, LATIN CAPITAL LETTER I WITH DOT ABOVE) lowercases to "i" +
  // U+0307 (combining dot above): 1 → 2 UTF-16 code units. A whole-string
  // toLowerCase() therefore de-syncs `folded` offsets from `text` offsets and
  // shifts every later match right by the length delta (reviewer repro:
  // "İstanbul" + "bul" sliced "ul"). The fold must preserve length — such code
  // points stay unfolded — so a match offset always indexes `text` correctly.
  it('finds a match after İ at the correct text offset', () => {
    const index = buildTextIndex(runs('İstanbul'));
    expect(index.folded.length).toBe(index.text.length);
    const m = findMatches(index, 'bul');
    expect(m).toHaveLength(1);
    expect(m[0].slices).toEqual([{ runIndex: 0, start: 5, end: 8 }]);
    // The slice must select the actual matched glyphs in the original text.
    expect('İstanbul'.slice(m[0].slices[0].start, m[0].slices[0].end)).toBe('bul');
  });

  it('keeps offsets aligned across a run boundary after İ', () => {
    const index = buildTextIndex(runs('İst', 'anbul'));
    const m = findMatches(index, 'STAN');
    expect(m).toHaveLength(1);
    expect(m[0].slices).toEqual([
      { runIndex: 0, start: 1, end: 3 },
      { runIndex: 1, start: 0, end: 2 },
    ]);
  });

  it('matches İ by identity (query İ ↔ text İ, both kept unfolded)', () => {
    const index = buildTextIndex(runs('İstanbul'));
    const m = findMatches(index, 'İstan');
    expect(m).toHaveLength(1);
    expect(m[0].slices).toEqual([{ runIndex: 0, start: 0, end: 5 }]);
  });

  it('ordinary folding still applies around the preserved code point', () => {
    const index = buildTextIndex(runs('İSTANBUL and istanbul'));
    // The İ variant's trailing letters and the plain-I word both fold normally.
    const m = findMatches(index, 'stanbul');
    expect(m).toHaveLength(2);
    expect(m.map((x) => x.slices[0].start)).toEqual([1, 14]);
  });
});

describe('findMatches — edge cases', () => {
  it('returns [] for an empty query', () => {
    const index = buildTextIndex(runs('anything'));
    expect(findMatches(index, '')).toEqual([]);
  });

  it('returns [] when the query is absent', () => {
    const index = buildTextIndex(runs('hello world'));
    expect(findMatches(index, 'zzz')).toEqual([]);
  });

  it('returns [] over an empty document', () => {
    const index = buildTextIndex([]);
    expect(findMatches(index, 'x')).toEqual([]);
  });

  it('matches a query that equals the whole text', () => {
    const index = buildTextIndex(runs('exact'));
    const m = findMatches(index, 'exact');
    expect(m).toHaveLength(1);
    expect(m[0].slices).toEqual([{ runIndex: 0, start: 0, end: 5 }]);
  });
});

describe('findMatches — several terms', () => {
  const starts = (m: ReturnType<typeof findMatches>) => m.map((x) => x.slices[0].start);

  it('treats a one-element list exactly like the bare string', () => {
    const index = buildTextIndex(runs('the quick brown fox, quick'));
    expect(findMatches(index, ['quick'])).toEqual(findMatches(index, 'quick'));
  });

  it('merges every term into one document-ordered list and renumbers matchIndex', () => {
    const index = buildTextIndex(runs('fox and dog and fox'));
    const m = findMatches(index, ['dog', 'fox']);
    expect(starts(m)).toEqual([0, 8, 16]);
    expect(m.map((x) => x.matchIndex)).toEqual([0, 1, 2]);
  });

  it('matches a spaced phrase contiguously, across runs', () => {
    const index = buildTextIndex(runs('quarterly ', 'earnings call; earnings'));
    const m = findMatches(index, ['quarterly earnings']);
    expect(m).toHaveLength(1);
    expect(m[0].slices).toEqual([
      { runIndex: 0, start: 0, end: 10 },
      { runIndex: 1, start: 0, end: 8 },
    ]);
  });

  it('keeps the longer match where terms overlap, so no character is highlighted twice', () => {
    const index = buildTextIndex(runs('earnings call and earnings'));
    const m = findMatches(index, ['earnings', 'earnings call']);
    expect(m.map((x) => x.slices[0])).toEqual([
      { runIndex: 0, start: 0, end: 13 },
      { runIndex: 0, start: 18, end: 26 },
    ]);
  });

  it('folds case for every term by default, and honours caseSensitive', () => {
    const index = buildTextIndex(runs('Alpha beta GAMMA'));
    expect(starts(findMatches(index, ['alpha', 'gamma']))).toEqual([0, 11]);
    expect(starts(findMatches(index, ['alpha', 'GAMMA'], { caseSensitive: true }))).toEqual([11]);
  });

  it('collapses terms that fold to the same text into one match per occurrence', () => {
    const index = buildTextIndex(runs('Report'));
    expect(findMatches(index, ['report', 'REPORT'])).toHaveLength(1);
  });

  it('searches for nothing when every term is empty', () => {
    const index = buildTextIndex(runs('anything'));
    expect(findMatches(index, [])).toEqual([]);
    expect(findMatches(index, ['', ''])).toEqual([]);
  });
});

describe('findMatches — term colours', () => {
  it("tags each match with its own term's colour, and leaves plain terms uncoloured", () => {
    const index = buildTextIndex(runs('fox and dog'));
    const m = findMatches(index, [{ text: 'dog', color: 'red' }, 'fox']);
    expect(m.map((x) => [x.slices[0].start, x.color])).toEqual([[0, undefined], [8, 'red']]);
    expect('color' in m[0]).toBe(false);
  });

  it('treats a term object without a colour exactly like the bare string', () => {
    const index = buildTextIndex(runs('fox and fox'));
    expect(findMatches(index, [{ text: 'fox' }])).toEqual(findMatches(index, 'fox'));
  });

  it('keeps the colour of the longer term where coloured and plain terms overlap', () => {
    const index = buildTextIndex(runs('earnings call and earnings'));
    const m = findMatches(index, ['earnings', { text: 'earnings call', color: 'red' }]);
    expect(m.map((x) => x.color)).toEqual(['red', undefined]);
  });
});

describe('findMatches — wholeWord', () => {
  const starts = (m: ReturnType<typeof findMatches>) => m.map((x) => x.slices[0].start);

  it('rejects a match inside a longer word and keeps the standalone one', () => {
    const index = buildTextIndex(runs('concatenate the cat'));
    expect(starts(findMatches(index, 'cat', { wholeWord: true }))).toEqual([16]);
    expect(starts(findMatches(index, 'cat'))).toEqual([3, 16]);
  });

  it('accepts a word bounded by punctuation or the ends of the text', () => {
    const index = buildTextIndex(runs('cat, (cat) cat'));
    expect(starts(findMatches(index, 'cat', { wholeWord: true }))).toEqual([0, 6, 11]);
  });

  it('only demands a boundary where the term itself starts or ends with a word character', () => {
    const index = buildTextIndex(runs('use C++, not C++x'));
    // "C++" ends in "+", so a following "," or "x" does not join it to a word.
    expect(starts(findMatches(index, 'C++', { wholeWord: true }))).toEqual([4, 13]);
    // Likewise at the start: "#tag" begins with "#", so the "x" before it does not join it.
    expect(starts(findMatches(buildTextIndex(runs('see x#tag')), '#tag', { wholeWord: true }))).toEqual([5]);
  });

  it('does not let a rejected occurrence hide a valid one that starts inside it', () => {
    const index = buildTextIndex(runs('ba a a'));
    // "a a" at 1 is glued to "b"; the whole-word "a a" at 3 overlaps it.
    expect(starts(findMatches(index, 'a a', { wholeWord: true }))).toEqual([3]);
  });

  it('treats non-ASCII letters, including astral ones, as word characters', () => {
    expect(findMatches(buildTextIndex(runs('cafécat')), 'cat', { wholeWord: true })).toEqual([]);
    expect(findMatches(buildTextIndex(runs('\u{1D400}cat')), 'cat', { wholeWord: true })).toEqual([]);
  });

  it('checks the boundary on the original text across a run seam', () => {
    const index = buildTextIndex(runs('bob', 'cat cat'));
    expect(starts(findMatches(index, 'cat', { wholeWord: true }))).toEqual([4]);
  });
});

describe('normalizeFindQuery', () => {
  it('drops empty terms and duplicates but does not trim', () => {
    expect(normalizeFindQuery(['a', '', 'a', ' b '])).toEqual([{ text: 'a' }, { text: ' b ' }]);
    expect(normalizeFindQuery('')).toEqual([]);
    expect(normalizeFindQuery([{ text: '', color: 'red' }])).toEqual([]);
  });

  it('keeps the first colour given for a repeated term', () => {
    expect(normalizeFindQuery([{ text: 'a', color: 'red' }, { text: 'a', color: 'blue' }, 'a']))
      .toEqual([{ text: 'a', color: 'red' }]);
  });
});

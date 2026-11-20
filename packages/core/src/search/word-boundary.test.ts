import { describe, expect, it } from 'vitest';
import { wordBoundaryTest, type WordSegmenter } from './word-boundary.js';

const boundaries = (text: string, test: (offset: number) => boolean): number[] => {
  const out: number[] = [];
  for (let offset = 0; offset <= text.length; offset++) if (test(offset)) out.push(offset);
  return out;
};

/** A segmenter that splits at fixed offsets of the string it is given. */
const fixedSegmenter = (...cuts: number[]): WordSegmenter & { calls: string[] } => {
  const calls: string[] = [];
  return {
    calls,
    segment(text) {
      calls.push(text);
      return [0, ...cuts].filter((index) => index < text.length).map((index) => ({ index }));
    },
  };
};

describe('wordBoundaryTest — scripts that separate words', () => {
  it('puts a boundary next to any non-word character and at both ends', () => {
    expect(boundaries('ab, cd', wordBoundaryTest('ab, cd', undefined, null))).toEqual([0, 2, 3, 4, 6]);
  });

  it('counts letters, combining marks, numbers and underscore as word characters', () => {
    const text = 'e\u0301a_1';
    expect(boundaries(text, wordBoundaryTest(text, undefined, null))).toEqual([0, text.length]);
  });

  it('never asks the segmenter, so the result does not depend on the engine', () => {
    const segmenter = fixedSegmenter(1, 2, 3);
    const text = 'can\'t stop example.com';
    expect(boundaries(text, wordBoundaryTest(text, undefined, segmenter)))
      .toEqual(boundaries(text, wordBoundaryTest(text, undefined, null)));
    expect(segmenter.calls).toEqual([]);
  });

  it('does not split a surrogate pair letter from its neighbours', () => {
    const text = 'a\u{1D400}b';
    expect(boundaries(text, wordBoundaryTest(text, undefined, null))).toEqual([0, text.length]);
  });

  it('treats every layout break as a boundary', () => {
    expect(boundaries('abcd', wordBoundaryTest('abcd', new Set([2]), null))).toEqual([0, 2, 4]);
  });
});

describe('wordBoundaryTest — scripts written without spaces', () => {
  it('takes boundaries inside a word-character run from the segmenter', () => {
    const text = '東京都に行く。';
    const segmenter = fixedSegmenter(2, 3, 4);
    expect(boundaries(text, wordBoundaryTest(text, undefined, segmenter))).toEqual([0, 2, 3, 4, 6, 7]);
  });

  it('segments each word-character run once, bounded by punctuation and breaks', () => {
    const text = '東京都、大阪府京都';
    const segmenter = fixedSegmenter(2);
    const test = wordBoundaryTest(text, new Set([7]), segmenter);
    expect(boundaries(text, test)).toEqual([0, 2, 3, 4, 6, 7, 9]);
    expect(segmenter.calls).toEqual(['東京都', '大阪府', '京都']);
  });

  it('walks a long run once, however many offsets are tested and in any order', () => {
    // 60,000 characters without a space. Looking up every offset is linear
    // overall; walking back to the run start per offset would take minutes.
    const text = '東京都大阪'.repeat(12_000);
    const segmenter = fixedSegmenter();
    const test = wordBoundaryTest(text, undefined, segmenter);
    const started = performance.now();
    let found = 0;
    for (let offset = text.length; offset >= 0; offset--) if (test(offset)) found++;
    for (let offset = 0; offset <= text.length; offset++) if (test(offset)) found++;
    expect(performance.now() - started).toBeLessThan(2_000);
    expect(found).toBe(4);
    expect(segmenter.calls).toHaveLength(1);
  });

  it('keeps runs apart when they are consulted out of order', () => {
    const text = '東京、大阪、京都';
    const segmenter = fixedSegmenter(1);
    const test = wordBoundaryTest(text, undefined, segmenter);
    expect([7, 1, 4].map((offset) => test(offset))).toEqual([true, true, true]);
    expect(boundaries(text, test)).toEqual([0, 1, 2, 3, 4, 5, 6, 7, 8]);
    expect(segmenter.calls).toEqual(['京都', '東京', '大阪']);
  });

  it('asks the segmenter where a spaced script meets an unspaced one', () => {
    const text = 'Tokyo東京';
    const segmenter = fixedSegmenter(5);
    expect(boundaries(text, wordBoundaryTest(text, undefined, segmenter))).toEqual([0, 5, 7]);
    expect(segmenter.calls).toEqual(['Tokyo東京']);
  });

  it('includes the prolonged sound mark with kana', () => {
    const text = 'コーヒー';
    expect(boundaries(text, wordBoundaryTest(text, undefined, fixedSegmenter()))).toEqual([0, 4]);
  });

  it('uses Intl.Segmenter by default when the engine has it', () => {
    const text = '東京都に行く';
    const segmenter = (Intl as { Segmenter?: unknown }).Segmenter;
    const found = boundaries(text, wordBoundaryTest(text));
    if (segmenter) {
      // A dictionary never reports every offset, and never none inside a sentence.
      expect(found.length).toBeGreaterThan(2);
      expect(found.length).toBeLessThan(text.length + 1);
    } else {
      expect(found).toEqual([0, 1, 2, 3, 4, 5, 6]);
    }
  });

  it('without a segmenter, treats every offset there as a boundary except before a combining mark', () => {
    const text = 'か\u3099き';
    expect(boundaries(text, wordBoundaryTest(text, undefined, null))).toEqual([0, 2, 3]);
    expect(boundaries('abc東京', wordBoundaryTest('abc東京', undefined, null))).toEqual([0, 3, 4, 5]);
  });
});

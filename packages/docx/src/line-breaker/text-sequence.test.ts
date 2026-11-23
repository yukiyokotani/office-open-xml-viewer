import { afterEach, describe, expect, it, vi } from 'vitest';
import type { ParagraphLayoutRun } from '../layout/text.js';
import type { LineLayoutEnvironment } from './model.js';
import { acquireTextSequences, type SequenceLanguages } from './text-sequence.js';

const run = (text: string, lang?: string, extra: Record<string, unknown> = {}) => ({
  type: 'text', text, fontSize: 10, fontFamily: 'Face', bold: false, italic: false,
  ...(lang ? { lang } : {}), ...extra,
}) as unknown as ParagraphLayoutRun;

const sequences = (runs: ParagraphLayoutRun[]) =>
  acquireTextSequences(runs, {} as LineLayoutEnvironment, (text) => text);

const languages = (sequence: { run: unknown } | undefined) =>
  (sequence?.run as SequenceLanguages | undefined)?.westernLanguageRanges?.map((range) => [range.end, range.lang]);

describe('text sequences across §17.3.2.20 language seams', () => {
  afterEach(() => vi.restoreAllMocks());

  it('joins a w:val-only seam inside a grapheme cluster and transfers only the crossing cluster', () => {
    const joined = sequences([run('xa', 'sr-rs'), run('\u0301b', 'en-us'), run('c', 'en-us'), run('d', 'sr-rs')]);
    expect([...joined.keys()]).toEqual([0]);
    const sequence = joined.get(0)!;
    // The next run with the second run's formatting joins as usual; a
    // language seam at a cluster boundary still delimits.
    expect(sequence.run.text).toBe('xa\u0301bc');
    expect(sequence.run.lang).toBe('sr-rs');
    // U+0301 completes the cluster 'a' began: it takes that run's language.
    // 'x', 'b' and 'c' keep their own.
    expect(languages(sequence)).toEqual([[3, 'sr-rs'], [5, 'en-us']]);
    // Source ranges stay those of the authored runs.
    expect(sequence.sources.map((source) => [source.runIndex, source.start, source.end]))
      .toEqual([[0, 0, 2], [1, 2, 4], [2, 4, 5]]);
  });

  it('keeps a cluster whole across three runs whose last carries the base language again', () => {
    const joined = sequences([run('\u0433', 'sr-rs'), run('\u0301', 'en-us'), run('\u0323x', 'sr-rs')]);
    const sequence = joined.get(0)!;
    expect(sequence.run.text).toBe('\u0433\u0301\u0323x');
    expect(languages(sequence)).toBeUndefined();
    expect(sequence.run.lang).toBe('sr-rs');
    expect(sequence.sources.map((source) => [source.runIndex, source.start, source.end]))
      .toEqual([[0, 0, 1], [1, 1, 2], [2, 2, 4]]);
    // With a third language the units after the cluster keep their own.
    const third = sequences([run('\u0433', 'sr-rs'), run('\u0301', 'en-us'), run('\u0323x', 'de-de')]).get(0)!;
    expect(languages(third)).toEqual([[3, 'sr-rs'], [4, 'de-de']]);
  });

  it('joins a base and a mark however long the text before the base', () => {
    for (const length of [0, 63, 64, 65, 5000]) {
      const joined = sequences([run(`${'x'.repeat(length)}a`, 'sr-rs'), run('\u0301', 'en-us')]);
      expect(languages(joined.get(0))).toBeUndefined();
      expect(joined.get(0)?.run.text.length).toBe(length + 2);
      expect(joined.get(0)?.run.lang).toBe('sr-rs');
    }
  });

  it('declines a seam inside a cluster longer than the bound', () => {
    // 64 units already in the cluster: one more mark would exceed the bound.
    expect(sequences([run(`a${'\u0301'.repeat(63)}`, 'sr-rs'), run('\u0301', 'en-us')]).size).toBe(0);
    expect(sequences([run(`a${'\u0301'.repeat(62)}`, 'sr-rs'), run('\u0301', 'en-us')]).size).toBe(1);
    // The bound is on the whole joined cluster, not on the second run.
    expect(sequences([run('a', 'sr-rs'), run('\u0301'.repeat(64), 'en-us')]).size).toBe(0);
    expect(sequences([run('a', 'sr-rs'), run(`${'\u0301'.repeat(63)}b`, 'en-us')]).size).toBe(1);
  });

  it('puts no language boundary inside a cluster longer than the bound', () => {
    // 63 units in Serbian, then two English runs of one mark each: the second
    // joins by formatting; the cluster (65 units) keeps one language.
    const joined = sequences([run(`a${'\u0301'.repeat(62)}`, 'sr-rs'), run('\u0301', 'en-us'), run('\u0301', 'en-us')]);
    const sequence = joined.get(0)!;
    expect(sequence.run.text.length).toBe(65);
    expect(languages(sequence)).toBeUndefined();
    expect(sequence.run.lang).toBe('sr-rs');
  });

  it('keeps delimiting at a boundary and at w:eastAsia or w:bidi seams', () => {
    expect(sequences([run('a', 'sr-rs'), run('b', 'en-us')]).size).toBe(0);
    expect(sequences([
      run('a', 'sr-rs', { langEastAsia: 'ja-jp' }), run('\u0301', 'sr-rs', { langEastAsia: 'zh-cn' }),
    ]).size).toBe(0);
    // Any other formatting difference delimits too, cluster or not.
    expect(sequences([run('a', 'sr-rs'), run('\u0301', 'en-us', { bold: true })]).size).toBe(0);
  });

  it('carries a ZWJ sequence across three runs and regional-indicator pairs by parity', () => {
    const zwj = sequences([
      run('\u{1f469}', 'fr-ch'), run('\u200d', 'en-us'), run('\u{1f467}b', 'en-us'), run('c', 'en-us'),
    ]);
    expect(zwj.get(0)?.run.text).toBe('\u{1f469}\u200d\u{1f467}bc');
    // The whole cluster takes its first unit's language, across both seams.
    expect(languages(zwj.get(0))).toEqual([[5, 'fr-ch'], [7, 'en-us']]);
    // Two complete flags: the third indicator starts a new cluster.
    const flags = sequences([run('\u{1f1ef}\u{1f1f5}', 'ja-jp'), run('\u{1f1fa}', 'en-us')]);
    expect(flags.size).toBe(0);
    const pair = sequences([run('\u{1f1ef}\u{1f1f5}\u{1f1fa}', 'ja-jp'), run('\u{1f1f8}', 'en-us')]);
    expect(pair.get(0)?.run.text).toBe('\u{1f1ef}\u{1f1f5}\u{1f1fa}\u{1f1f8}');
    expect(languages(pair.get(0))).toBeUndefined();
  });

  it('segments each unit once plus a bounded cluster per seam', () => {
    let units = 0;
    const segment = Intl.Segmenter.prototype.segment;
    vi.spyOn(Intl.Segmenter.prototype, 'segment').mockImplementation(function (this: Intl.Segmenter, input: string) {
      units += input.length;
      return segment.call(this, input);
    });
    // Every seam continues the cluster until it reaches the bound.
    const runs = [run('a', 'sr-rs')];
    for (let index = 0; index < 4000; index += 1) runs.push(run('\u0301', index % 2 ? 'sr-rs' : 'en-us'));
    sequences(runs);
    expect(units).toBeLessThanOrEqual(4000 * 66);
    // Long runs between seams are scanned once, not once per seam.
    units = 0;
    const long = [run('a', 'sr-rs')];
    for (let index = 0; index < 200; index += 1) {
      long.push(run(`\u0301${'x'.repeat(1000)}a`, index % 2 ? 'sr-rs' : 'en-us'));
    }
    expect(sequences(long).get(0)?.sources.length).toBe(201);
    expect(units).toBeLessThanOrEqual(200 * (1002 + 2 * 66));
    // Untagged text is never segmented here.
    units = 0;
    sequences([run('a\u0301'), run('\u0301b'), run('c', undefined, { bold: true })]);
    expect(units).toBe(0);
  });
});

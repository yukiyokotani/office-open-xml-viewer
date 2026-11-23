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

  it('keeps a cluster of any length whole and transfers only its units', () => {
    // 64 units already in the cluster, then one more mark: still one cluster.
    expect(sequences([run(`a${'\u0301'.repeat(63)}`, 'sr-rs'), run('\u0301', 'en-us')]).size).toBe(1);
    expect(sequences([run('a', 'sr-rs'), run('\u0301'.repeat(64), 'en-us')]).size).toBe(1);
    // Only the crossing cluster takes the base language; 'b' keeps its own.
    const joined = sequences([run(`a${'\u0301'.repeat(62)}`, 'sr-rs'), run('\u0301', 'en-us'), run('\u0301b', 'en-us')]);
    expect(languages(joined.get(0))).toEqual([[65, 'sr-rs'], [66, 'en-us']]);
  });

  it('treats runs naming one language on every axis as one language fact', () => {
    // OpenDocument runs: each names its script class's language on all axes.
    const uniform = (text: string, lang: string) => run(text, lang, { langEastAsia: lang, langBidi: lang });
    const joined = sequences([uniform('x\u0433', 'sr-rs'), uniform('\u0301b', 'en-us')]).get(0)!;
    expect(joined.run.text).toBe('x\u0433\u0301b');
    expect(languages(joined)).toEqual([[3, 'sr-rs'], [4, 'en-us']]);
    expect((joined.run as SequenceLanguages).uniformLanguageRanges).toBe(true);
    // A boundary still delimits.
    expect(sequences([uniform('a', 'sr-rs'), uniform('b', 'en-us')]).size).toBe(0);
    // Non-uniform runs differing beyond w:val still delimit, as before.
    expect(sequences([
      run('\u0433', 'sr-rs', { langEastAsia: 'ja-jp' }), run('\u0301', 'en-us', { langEastAsia: 'zh-cn' }),
    ]).size).toBe(0);
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
    // Every seam continues one ever-growing cluster: a bounded context per
    // seam (64 units of the cluster, the seam pre-check, the next run).
    const marks = (count: number) => {
      const runs = [run('a', 'sr-rs')];
      for (let index = 0; index < count; index += 1) runs.push(run('\u0301', index % 2 ? 'sr-rs' : 'en-us'));
      units = 0;
      sequences(runs);
      return units;
    };
    const thousand = marks(1000);
    const fourThousand = marks(4000);
    expect(fourThousand).toBeLessThanOrEqual(4000 * 70);
    expect(fourThousand / thousand).toBeLessThan(4.2);
    // Long runs between seams are scanned once, not once per seam.
    units = 0;
    const long = [run('a', 'sr-rs')];
    for (let index = 0; index < 200; index += 1) {
      long.push(run(`\u0301${'x'.repeat(1000)}a`, index % 2 ? 'sr-rs' : 'en-us'));
    }
    expect(sequences(long).get(0)?.sources.length).toBe(201);
    expect(units).toBeLessThanOrEqual(200 * (1002 + 3 * 70));
    // Untagged text is never segmented here.
    units = 0;
    sequences([run('a\u0301'), run('\u0301b'), run('c', undefined, { bold: true })]);
    expect(units).toBe(0);
  });
});

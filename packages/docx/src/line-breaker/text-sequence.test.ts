import { afterEach, describe, expect, it, vi } from 'vitest';
import type { ParagraphLayoutRun } from '../layout/text.js';
import type { LineLayoutEnvironment } from './model.js';
import { acquireTextSequences } from './text-sequence.js';

const run = (text: string, lang?: string, extra: Record<string, unknown> = {}) => ({
  type: 'text', text, fontSize: 10, fontFamily: 'Face', bold: false, italic: false,
  ...(lang ? { lang } : {}), ...extra,
}) as unknown as ParagraphLayoutRun;

const sequences = (runs: ParagraphLayoutRun[]) =>
  acquireTextSequences(runs, {} as LineLayoutEnvironment, (text) => text);

describe('text sequences across §17.3.2.20 language seams', () => {
  afterEach(() => vi.restoreAllMocks());

  it('joins a w:val-only seam inside a grapheme cluster and drops the western language', () => {
    const joined = sequences([run('a', 'sr-rs'), run('́b', 'en-us'), run('c', 'en-us')]);
    expect([...joined.keys()]).toEqual([0]);
    const sequence = joined.get(0)!;
    expect(sequence.run.text).toBe('áb');
    expect(sequence.run.lang).toBeUndefined();
    expect(sequence.sources.map((source) => [source.runIndex, source.start, source.end])).toEqual([[0, 0, 1], [1, 1, 3]]);
  });

  it('keeps delimiting at a boundary and at w:eastAsia or w:bidi seams', () => {
    expect(sequences([run('a', 'sr-rs'), run('b', 'en-us')]).size).toBe(0);
    expect(sequences([
      run('a', 'sr-rs', { langEastAsia: 'ja-jp' }), run('́', 'sr-rs', { langEastAsia: 'zh-cn' }),
    ]).size).toBe(0);
  });

  it('carries a ZWJ sequence across three runs and regional-indicator pairs by parity', () => {
    const zwj = sequences([run('\u{1f469}', 'fr-ch'), run('‍', 'en-us'), run('\u{1f467}', 'de-de')]);
    expect(zwj.get(0)?.run.text).toBe('\u{1f469}‍\u{1f467}');
    // Two complete flags: the third indicator starts a new cluster.
    const flags = sequences([run('\u{1f1ef}\u{1f1f5}', 'ja-jp'), run('\u{1f1fa}', 'en-us')]);
    expect(flags.size).toBe(0);
    const pair = sequences([run('\u{1f1ef}\u{1f1f5}\u{1f1fa}', 'ja-jp'), run('\u{1f1f8}', 'en-us')]);
    expect(pair.get(0)?.run.text).toBe('\u{1f1ef}\u{1f1f5}\u{1f1fa}\u{1f1f8}');
  });

  it('bounds the text segmented again over many continuing-mark seams', () => {
    let units = 0;
    const segment = Intl.Segmenter.prototype.segment;
    vi.spyOn(Intl.Segmenter.prototype, 'segment').mockImplementation(function (this: Intl.Segmenter, input: string) {
      units += input.length;
      return segment.call(this, input);
    });
    const runs = [run('a', 'sr-rs')];
    for (let index = 0; index < 4000; index += 1) runs.push(run('́', index % 2 ? 'sr-rs' : 'en-us'));
    sequences(runs);
    expect(units).toBeLessThan(4000 * 80);
  });
});

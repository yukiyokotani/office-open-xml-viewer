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

const ranges = (sequence: { run: unknown } | undefined) =>
  (sequence?.run as SequenceLanguages | undefined)?.languageRanges;

/** [end, w:val] of each range. */
const languages = (sequence: { run: unknown } | undefined) =>
  ranges(sequence)?.map((range) => [range.end, range.lang]);

/** The grapheme cluster boundaries of `text` (UAX #29). */
const clusterBoundaries = (text: string): Set<number> => {
  const boundaries = new Set<number>([text.length]);
  for (const part of new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(text)) {
    boundaries.add(part.index);
  }
  return boundaries;
};

describe('text sequences across §17.3.2.20 language seams', () => {
  afterEach(() => vi.restoreAllMocks());

  it('displays a sequence again only from its first language seam', () => {
    const calls: string[] = [];
    const acquire = (runs: ParagraphLayoutRun[]) => {
      calls.length = 0;
      const result = acquireTextSequences(runs, {} as LineLayoutEnvironment, (text) => {
        calls.push(text);
        return text.toUpperCase();
      });
      return result;
    };
    // One language state: each run is displayed once, as before languages,
    // and no range text is collected.
    acquire([run('a'), run('b'), run('c')]);
    expect(calls).toEqual(['a', 'b', 'c']);
    // A seam displays the runs before it once more; ranges index the
    // displayed text.
    const seamed = acquire([run('a'), run('b'), run('c', 'sr-rs'), run('d')]);
    expect(calls).toEqual(['a', 'b', 'c', 'a', 'b', 'd']);
    expect(languages(seamed.get(0))).toEqual([[2, undefined], [3, 'sr-rs'], [4, undefined]]);
  });

  it('joins runs that differ only in language and records each stretch', () => {
    const joined = sequences([run('xa', 'sr-rs'), run('́b', 'en-us'), run('c', 'en-us'), run('d', 'sr-rs')]);
    expect([...joined.keys()]).toEqual([0]);
    const sequence = joined.get(0)!;
    expect(sequence.run.text).toBe('xábcd');
    expect(sequence.run.lang).toBe('sr-rs');
    // U+0301 completes the cluster 'a' began: it takes that run's language.
    // 'x', 'b', 'c' and 'd' keep their own.
    expect(languages(sequence)).toEqual([[3, 'sr-rs'], [5, 'en-us'], [6, 'sr-rs']]);
    // Source ranges stay those of the authored runs.
    expect(sequence.sources.map((source) => [source.runIndex, source.start, source.end]))
      .toEqual([[0, 0, 2], [1, 2, 4], [2, 4, 5], [3, 5, 6]]);
  });

  it('records nothing for runs of one language state, tagged or not', () => {
    const untagged = sequences([run('á'), run('́b')]).get(0)!;
    expect(untagged.run.text).toBe('á́b');
    expect('languageRanges' in untagged.run).toBe(false);
    expect(ranges(sequences([run('a', 'sr-rs'), run('b', 'sr-rs')]).get(0))).toBeUndefined();
    // x-none, an ill-formed tag and no tag are one state (canvasLanguageTag).
    expect(ranges(sequences([run('a'), run('b', 'x-none'), run('c', 'not a tag')]).get(0))).toBeUndefined();
    // A run forced onto the complex-script slot never uses w:val.
    const forced = sequences([
      run('س', 'en-us', { cs: true, langBidi: 'ar-sa' }), run('ل', 'fr-fr', { cs: true, langBidi: 'ar-sa' }),
    ]).get(0);
    expect(forced?.run.text).toBe('سل');
    expect(ranges(forced)).toBeUndefined();
  });

  it('keeps a cluster whole across three runs whose last carries the base language again', () => {
    const joined = sequences([run('г', 'sr-rs'), run('́', 'en-us'), run('̣x', 'sr-rs')]);
    const sequence = joined.get(0)!;
    expect(sequence.run.text).toBe('ѓ̣x');
    expect(languages(sequence)).toBeUndefined();
    expect(sequence.run.lang).toBe('sr-rs');
    expect(sequence.sources.map((source) => [source.runIndex, source.start, source.end]))
      .toEqual([[0, 0, 1], [1, 1, 2], [2, 2, 4]]);
    // With a third language the units after the cluster keep their own.
    const third = sequences([run('г', 'sr-rs'), run('́', 'en-us'), run('̣x', 'de-de')]).get(0)!;
    expect(languages(third)).toEqual([[3, 'sr-rs'], [4, 'de-de']]);
  });

  it('keeps a cluster of any length whole and transfers only its units', () => {
    for (const length of [0, 63, 64, 65, 5000]) {
      const joined = sequences([run(`${'x'.repeat(length)}a`, 'sr-rs'), run('́', 'en-us')]);
      expect(languages(joined.get(0))).toBeUndefined();
      expect(joined.get(0)?.run.text.length).toBe(length + 2);
    }
    for (const marks of [62, 63, 64, 200]) {
      const joined = sequences([
        run(`a${'́'.repeat(marks)}`, 'sr-rs'), run('́', 'en-us'), run('́b', 'en-us'),
      ]);
      expect(languages(joined.get(0))).toEqual([[marks + 3, 'sr-rs'], [marks + 4, 'en-us']]);
    }
  });

  it('judges a long emoji ZWJ cluster in its full context (UAX #29 GB11)', () => {
    // An emoji, 66 extending marks, then ZWJ and a second emoji in other
    // runs: one cluster of 71 units, whatever its length.
    const text = `\u{1f469}${'́'.repeat(66)}‍\u{1f467}b`;
    expect(Math.min(...[...clusterBoundaries(text)].filter((offset) => offset > 0))).toBe(71);
    const joined = sequences([
      run(`\u{1f469}${'́'.repeat(66)}`, 'fr-ch'), run('‍', 'en-us'), run('\u{1f467}b', 'en-us'),
    ]).get(0)!;
    expect(joined.run.text).toBe(text);
    expect(languages(joined)).toEqual([[71, 'fr-ch'], [72, 'en-us']]);
  });

  it('keeps an Indic conjunct with its first consonant (UAX #29 GB9c)', () => {
    // KA, 70 nuktas, VIRAMA | SSA 'b': the conjunct is one cluster.
    const head = `क${'़'.repeat(70)}्`;
    const text = `${head}षb`;
    const boundaries = clusterBoundaries(text);
    const joined = sequences([run(head, 'hi-in'), run('षb', 'en-us')]).get(0)!;
    const found = ranges(joined)!;
    for (const range of found) expect(boundaries.has(range.end)).toBe(true);
    // Where the engine implements GB9c the conjunct is one cluster.
    const conjunctEnd = boundaries.has(head.length) ? head.length : head.length + 1;
    expect(languages(joined)).toEqual([[conjunctEnd, 'hi-in'], [text.length, 'en-us']]);
  });

  it('places every range boundary on a grapheme cluster boundary of the whole text', () => {
    const pieces = ['a', '́', '‍', '\u{1f469}', '\u{1f467}', '️', 'क', '्', '\u{1f1ef}',
      '\u{1f1f5}', 'x', 'ᄀ', 'ᅡ', '\r', '\n', '؀', 'ः'];
    const tags = ['sr-rs', 'en-us', 'ja-jp', undefined];
    let seed = 7;
    const next = (bound: number) => {
      seed = (seed * 1103515245 + 12345) & 0x7fffffff;
      return seed % bound;
    };
    for (let round = 0; round < 200; round += 1) {
      const runs = Array.from({ length: 2 + next(6) }, () => run(
        Array.from({ length: 1 + next(4) }, () => pieces[next(pieces.length)]!).join(''), tags[next(tags.length)],
      ));
      const sequence = sequences(runs).get(0)!;
      const boundaries = clusterBoundaries(sequence.run.text);
      const found = ranges(sequence);
      if (!found) continue;
      expect(found.at(-1)!.end).toBe(sequence.run.text.length);
      let start = 0;
      for (const range of found) {
        expect(boundaries.has(range.end)).toBe(true);
        // The range's language is that of the run owning its first unit.
        const owner = sequence.sources.find((source) => source.start <= start && start < source.end)!;
        expect(range.lang).toBe((runs[owner.runIndex] as { lang?: string }).lang);
        expect(range.end).toBeGreaterThan(start);
        start = range.end;
      }
    }
  });

  it('carries the whole language tuple of the owning run in every range', () => {
    // A base in sr-RS on every axis, a mark and kana in ja-JP on every axis,
    // then a mark and Han with en-US on w:val and ja-JP on the other axes.
    const all = (lang: string) => ({ langEastAsia: lang, langBidi: lang });
    const joined = sequences([
      run('г', 'sr-rs', all('sr-rs')), run('́か', 'ja-jp', all('ja-jp')),
      run('́漢', 'en-us', all('ja-jp')),
    ]).get(0)!;
    expect(joined.run.text).toBe('ѓか́漢');
    expect(ranges(joined)).toEqual([
      { end: 2, lang: 'sr-rs', langEastAsia: 'sr-rs', langBidi: 'sr-rs', langScriptClass: undefined },
      // U+304B and its mark U+0301 (third run) are the second run's cluster.
      { end: 4, lang: 'ja-jp', langEastAsia: 'ja-jp', langBidi: 'ja-jp', langScriptClass: undefined },
      { end: 5, lang: 'en-us', langEastAsia: 'ja-jp', langBidi: 'ja-jp', langScriptClass: undefined },
    ]);
  });

  it('joins across w:eastAsia, w:bidi and script-class seams, each range with its own', () => {
    const eastAsia = sequences([
      run('a', 'sr-rs', { langEastAsia: 'ja-jp' }), run('́漢', 'sr-rs', { langEastAsia: 'zh-cn' }),
    ]).get(0)!;
    expect(ranges(eastAsia)?.map((range) => [range.end, range.langEastAsia])).toEqual([[2, 'ja-jp'], [3, 'zh-cn']]);
    const bidi = sequences([run('a', undefined, { langBidi: 'ar-sa' }), run('b', undefined, { langBidi: 'he-il' })]).get(0)!;
    expect(ranges(bidi)?.map((range) => [range.end, range.langBidi])).toEqual([[1, 'ar-sa'], [2, 'he-il']]);
    // An OpenDocument class language is a fact of its own: present and
    // empty (the class names none) differs from absent.
    const classed = sequences([
      run('г', 'fr-fr', { langScriptClass: 'fr-fr' }), run('́字', 'fr-fr', { langScriptClass: 'zh-cn' }),
      run('x', 'fr-fr', { langScriptClass: '' }),
    ]).get(0)!;
    expect(ranges(classed)?.map((range) => [range.end, range.langScriptClass]))
      .toEqual([[2, 'fr-fr'], [3, 'zh-cn'], [4, '']]);
    // Any other formatting difference still delimits, cluster or not.
    expect(sequences([run('a', 'sr-rs'), run('́', 'en-us', { bold: true })]).size).toBe(0);
  });

  it('takes the languages of the first non-empty run', () => {
    const joined = sequences([run('', 'sr-rs'), run('ab', 'en-us')]).get(0)!;
    expect(joined.run.lang).toBe('sr-rs');
    expect(languages(joined)).toEqual([[2, 'en-us']]);
  });

  it('carries a ZWJ sequence across three runs and regional-indicator pairs by parity', () => {
    const zwj = sequences([
      run('\u{1f469}', 'fr-ch'), run('‍', 'en-us'), run('\u{1f467}b', 'en-us'), run('c', 'en-us'),
    ]);
    expect(zwj.get(0)?.run.text).toBe('\u{1f469}‍\u{1f467}bc');
    // The whole cluster takes its first unit's language, across both seams.
    expect(languages(zwj.get(0))).toEqual([[5, 'fr-ch'], [7, 'en-us']]);
    // Two complete flags: the third indicator starts a new cluster.
    const flags = sequences([run('\u{1f1ef}\u{1f1f5}', 'ja-jp'), run('\u{1f1fa}', 'en-us')]);
    expect(languages(flags.get(0))).toEqual([[4, 'ja-jp'], [6, 'en-us']]);
    const pair = sequences([run('\u{1f1ef}\u{1f1f5}\u{1f1fa}', 'ja-jp'), run('\u{1f1f8}', 'en-us')]);
    expect(pair.get(0)?.run.text).toBe('\u{1f1ef}\u{1f1f5}\u{1f1fa}\u{1f1f8}');
    expect(languages(pair.get(0))).toBeUndefined();
  });

  it('keeps the languages of symbol-encoded runs as sequence boundaries', () => {
    const symbol = (text: string, lang: string) => run(text, lang, { fontFamily: 'Symbol' });
    expect(sequences([symbol('a', 'sr-rs'), symbol('b', 'en-us')]).size).toBe(0);
    expect(sequences([symbol('a', 'sr-rs'), symbol('b', 'sr-rs')]).get(0)?.run.text).toBe('ab');
  });

  it('segments a sequence at most once, and only when a seam may fall inside a cluster', () => {
    let units = 0;
    let calls = 0;
    const segment = Intl.Segmenter.prototype.segment;
    vi.spyOn(Intl.Segmenter.prototype, 'segment').mockImplementation(function (this: Intl.Segmenter, input: string) {
      units += input.length;
      calls += 1;
      return segment.call(this, input);
    });
    // Every seam continues one ever-growing cluster: a three-scalar pre-check
    // per seam and one pass over the text, whatever the number of seams.
    const marks = (count: number) => {
      const runs = [run('a', 'sr-rs')];
      for (let index = 0; index < count; index += 1) runs.push(run('́', index % 2 ? 'sr-rs' : 'en-us'));
      units = 0;
      expect(languages(sequences(runs).get(0))).toBeUndefined();
      return units;
    };
    expect(marks(1000)).toBeLessThanOrEqual(1001 + 1000 * 3);
    expect(marks(4000)).toBeLessThanOrEqual(4001 + 4000 * 3);
    // Long runs between seams are scanned once, not once per seam.
    units = 0;
    const long = [run('a', 'sr-rs')];
    for (let index = 0; index < 200; index += 1) {
      long.push(run(`́${'x'.repeat(1000)}a`, index % 2 ? 'sr-rs' : 'en-us'));
    }
    expect(sequences(long).get(0)?.sources.length).toBe(201);
    expect(units).toBeLessThanOrEqual(200 * 1002 + 1 + 200 * 3);
    // Seams between scalars below U+0300 are boundaries without a segmenter;
    // other seams that cannot join cost only their pre-check.
    calls = 0;
    sequences([run('Hello ', 'en-us'), run('world', 'en-gb'), run('!', 'fr-fr')]);
    expect(calls).toBe(0);
    units = 0;
    sequences([run('漢'.repeat(1000), 'ja-jp'), run('字'.repeat(1000), 'zh-cn')]);
    expect(units).toBeLessThanOrEqual(3);
    // Untagged text is never segmented here.
    calls = 0;
    sequences([run('á'), run('́b'), run('c', undefined, { bold: true })]);
    expect(calls).toBe(0);
  });
});

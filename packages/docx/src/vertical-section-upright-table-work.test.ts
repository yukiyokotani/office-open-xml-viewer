import { beforeAll, describe, expect, it, vi } from 'vitest';
import {
  P,
  T,
  layOutPages,
  loadParser,
  paragraph,
  table,
} from './test-support/vertical-section-upright.test-support.js';

// `word-vertical-section-upright-block-table`: the work of placing upright
// block tables grows with their number, not with its square.
//
// What is counted: every element an Array method or array iteration visits
// while a document is laid out, when that element is
// - a box of the page's upright block index (upright-block-index.ts), known
//   by identity through a passthrough around its registration,
// - an upright block's text exclusion among the page's wrap rectangles, a
//   slot of a table placed ahead of the flow, or a float registry entry of an
//   upright block (there are none: the blocks are kept beside that registry),
// - a retained row of a table.
// These are the collections whose rescans made earlier revisions quadratic.
// Indexed loops and Map lookups are not counted, and nothing here measures
// time: the counts bound array traversals only.

const work = vi.hoisted(() => ({ boxes: new WeakSet<object>(), registered: 0 }));

vi.mock('./layout/upright-block-index.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('./layout/upright-block-index.js')>();
  return {
    ...actual,
    addUprightBlockBox: (...args: Parameters<typeof actual.addUprightBlockBox>) => {
      work.boxes.add(args[1]);
      work.registered += 1;
      return actual.addUprightBlockBox(...args);
    },
  };
});

type Counts = { boxes: number; exclusions: number; slots: number; registryEntries: number; rows: number };

/** Lays `run` out while counting the visits described above. */
function counting<Result>(run: () => Result): Readonly<{ result: Result; counts: Counts }> {
  const counts: Counts = { boxes: 0, exclusions: 0, slots: 0, registryEntries: 0, rows: 0 };
  const visit = (value: unknown): void => {
    if (value === null || typeof value !== 'object') return;
    if (work.boxes.has(value)) counts.boxes += 1;
    else if ('bandStartInclusive' in value) counts.exclusions += 1;
    else if ('reserved' in value && 'sourceIndex' in value) counts.slots += 1;
    else if ('kind' in value && value.kind === 'table-row') counts.rows += 1;
    else if ('occurrenceId' in value && typeof value.occurrenceId === 'string'
      && value.occurrenceId.includes('upright-table')) counts.registryEntries += 1;
  };
  const prototype = Array.prototype as unknown as Record<string | symbol, (...args: never[]) => unknown>;
  const restore: Array<() => void> = [];
  const replace = (name: string | symbol, replacement: (original: (...args: never[]) => unknown) => unknown) => {
    const original = prototype[name]!;
    prototype[name] = replacement(original) as (...args: never[]) => unknown;
    restore.push(() => { prototype[name] = original; });
  };
  // Plain overrides: a spy's own bookkeeping would iterate arrays in turn.
  replace(Symbol.iterator, (original) => function (this: unknown[]) {
    const iterator = original.call(this) as Iterator<unknown>;
    return (function* () {
      for (let next = iterator.next(); !next.done; next = iterator.next()) {
        visit(next.value);
        yield next.value;
      }
    })();
  });
  for (const name of ['map', 'filter', 'some', 'every', 'find', 'findIndex', 'forEach', 'flatMap']) {
    replace(name, (original) => function (this: unknown[], callback: (...args: unknown[]) => unknown, self?: unknown) {
      return original.call(this, (function (this: unknown, value: unknown, ...rest: unknown[]) {
        visit(value);
        return callback.call(this, value, ...rest);
      }) as never, self as never);
    });
  }
  replace('reduce', (original) => function (this: unknown[], callback: (...args: unknown[]) => unknown, ...initial: unknown[]) {
    return original.call(this, ((accumulator: unknown, value: unknown, ...rest: unknown[]) => {
      visit(value);
      return callback(accumulator, value, ...rest);
    }) as never, ...(initial as never[]));
  });
  for (const name of ['slice', 'concat']) {
    replace(name, (original) => function (this: unknown[], ...args: unknown[]) {
      const copy = original.call(this, ...(args as never[])) as unknown[];
      for (let index = 0; index < copy.length; index += 1) visit(copy[index]);
      return copy;
    });
  }
  try {
    return { result: run(), counts };
  } finally {
    for (const undo of restore.reverse()) undo();
  }
}

beforeAll(loadParser);

describe('upright block table placement work', () => {
  it('places tables that run on from one another in work proportional to their number', () => {
    // Consecutive one-row tables of a two-column page are one block.
    const measure = (count: number) => {
      const registeredBefore = work.registered;
      const { result, counts } = counting(() => layOutPages('tbRl', Array.from({ length: count }, () => table(1)), 2));
      expect(result.flatMap((page) => page.tables)).toHaveLength(count);
      return { ...counts, registered: work.registered - registeredBefore };
    };
    for (const count of [16, 32, 64]) {
      const counts = measure(count);
      // One box for the whole run, which the tables in it do not read again.
      expect(counts.registered, `${count} tables`).toBe(1);
      expect(counts.boxes, `${count} tables`).toBeLessThanOrEqual(2);
      expect(counts.registryEntries, `${count} tables`).toBe(0);
      // Each table's slot and exclusion is handled a bounded number of times.
      expect(counts.slots, `${count} tables`).toBeLessThanOrEqual(4 * count);
      expect(counts.exclusions, `${count} tables`).toBeLessThanOrEqual(4 * count);
    }
  }, 120_000);

  it('places tables between paragraphs of one page in work proportional to their number', () => {
    // Tables one strip apart stack down the column, a handful to a stack,
    // then start the next stack past it: each placement reads only the boxes
    // near it, however many the document holds.
    const measure = (count: number) => {
      const body = Array.from({ length: count }, () => [P, table(6)]).flat();
      const { result, counts } = counting(() => layOutPages('tbRl', body, 2));
      expect(result.flatMap((page) => page.tables)).toHaveLength(count);
      return counts;
    };
    const small = measure(24);
    const large = measure(96);
    expect(large.registryEntries).toBe(0);
    // Four times the tables: about four times the box reads (a scan of every
    // box placed so far would read sixteen times as many).
    expect(small.boxes).toBeGreaterThan(0);
    expect(large.boxes).toBeLessThanOrEqual(6 * small.boxes);
    // A paragraph is laid out against the exclusions its lines can meet,
    // not against those of every table of the page.
    expect(small.exclusions).toBeGreaterThan(0);
    expect(large.exclusions).toBeLessThanOrEqual(6 * small.exclusions);
  }, 120_000);

  it('places a table on each page in work that does not grow with the pages before it', () => {
    // A full-length strip, a table it wraps around and a page break per page.
    const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
    const measure = (pageCount: number) => {
      const body = Array.from({ length: pageCount }, (_, page) => [
        paragraph('長'.repeat(66)), T, ...(page + 1 < pageCount ? [pageBreak] : []),
      ]).flat();
      const { result, counts } = counting(() => layOutPages('tbRl', body));
      expect(result).toHaveLength(pageCount);
      // Every strip before a table wraps around it: two strips per page.
      for (const page of result) expect(page.strips[0]).toHaveLength(2);
      return counts;
    };
    const small = measure(8);
    const large = measure(32);
    const perPage = (counts: Counts) => counts.boxes + counts.exclusions + counts.slots;
    expect(perPage(large)).toBeLessThanOrEqual(4 * perPage(small));
    expect(large.rows).toBeLessThanOrEqual(4 * small.rows);
    expect(large.registryEntries).toBe(0);
  }, 120_000);

  it('continues a table over pages in row work proportional to its rows', () => {
    // 115 pt rows that never split: six rows a page.
    const tall = (rows: number) => table(rows, 115).replaceAll('<w:trPr>', '<w:trPr><w:cantSplit/>');
    const measure = (rows: number) => {
      const { result, counts } = counting(() => layOutPages('tbRl', [tall(rows)]));
      expect(result.length).toBeGreaterThanOrEqual(rows / 6);
      return counts.rows;
    };
    const small = measure(40);
    const large = measure(160);
    // Four times the rows over four times the pages: summing the remaining
    // rows for every page would read sixteen times as many.
    expect(large).toBeLessThanOrEqual(5 * small);
  }, 120_000);
});

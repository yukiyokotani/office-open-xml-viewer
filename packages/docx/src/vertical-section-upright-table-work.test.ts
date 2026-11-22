import { beforeAll, describe, expect, it, vi } from 'vitest';
import {
  P,
  T,
  anchoredRectangle,
  layOutPages,
  loadParser,
  paragraph,
  romanPageTable,
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
// - an entry of the page's history of placed table boxes,
// - a member of a logical table as it is handed to its acquisition,
// - an entry of the page's index of wrapped anchored drawings,
// - a retained row of a table.
// Reads of a table's retained record are counted by a second passthrough.
// These are the collections whose rescans made earlier revisions quadratic.
// Indexed loops and Map lookups are not counted, and nothing here measures
// time: the counts bound array traversals only.

const work = vi.hoisted(() => ({
  boxes: new WeakSet<object>(), registered: 0, recordReads: 0, failPinnedLayouts: false, failures: 0,
  pageIndexes: 0,
}));

vi.mock('./layout/upright-block-index.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('./layout/upright-block-index.js')>();
  return {
    ...actual,
    // One per layout of a page, and one per prescan that indexes drawings.
    emptyUprightBlockPageIndex: () => {
      work.pageIndexes += 1;
      return actual.emptyUprightBlockPageIndex();
    },
    addUprightBlockBox: (...args: Parameters<typeof actual.addUprightBlockBox>) => {
      work.boxes.add(args[1]);
      work.registered += 1;
      return actual.addUprightBlockBox(...args);
    },
    // Stands for any failure of a layout that holds pins (see the test
    // that sets it).
    uprightBlockBoxesMet: (...args: Parameters<typeof actual.uprightBlockBoxesMet>) => {
      if (work.failPinnedLayouts && args[0].segmentPins.size > 0) {
        work.failures += 1;
        throw new Error('pinned layout failure');
      }
      return actual.uprightBlockBoxesMet(...args);
    },
  };
});

vi.mock('./layout/acquisition-state.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('./layout/acquisition-state.js')>();
  return {
    ...actual,
    retainedTableRecord: (...args: Parameters<typeof actual.retainedTableRecord>) => {
      work.recordReads += 1;
      return actual.retainedTableRecord(...args);
    },
  };
});

type Counts = {
  boxes: number; exclusions: number; slots: number; registryEntries: number; rows: number; history: number;
  members: number; drawings: number;
  /** Float registry entries of wrapped drawings. */
  floats: number;
  /** Frames of anchor occurrences serialized into keys (JSON.stringify of
   * [occurrence, frame] pairs). */
  frameKeys: number;
  /** Table pin keys read by iterating a Set. */
  pinKeys: number;
  /** Elements an insertion into or removal from an array of drawings moves. */
  drawingShifts: number;
};

/** Lays `run` out while counting the visits described above. */
function counting<Result>(
  run: () => Result,
  /** Count `pinKeys` too: every Set iteration of the layout is then wrapped. */
  setIteration = false,
): Readonly<{ result: Result; counts: Counts }> {
  const counts: Counts = {
    boxes: 0, exclusions: 0, slots: 0, registryEntries: 0, rows: 0, history: 0, members: 0, drawings: 0, floats: 0,
    frameKeys: 0, pinKeys: 0,
    drawingShifts: 0,
  };
  const visit = (value: unknown): void => {
    if (value === null || typeof value !== 'object') return;
    if (work.boxes.has(value)) counts.boxes += 1;
    else if ('bandStartInclusive' in value) counts.exclusions += 1;
    else if ('reserved' in value && 'sourceIndex' in value) counts.slots += 1;
    else if ('box' in value && 'sourceIndex' in value) counts.history += 1;
    else if ('table' in value && 'sourceIndex' in value) counts.members += 1;
    else if ('anchorIndex' in value && 'frame' in value) counts.drawings += 1;
    else if ('kind' in value && value.kind === 'table-row') counts.rows += 1;
    else if ('kind' in value && value.kind === 'shape' && 'exclusionBounds' in value) counts.floats += 1;
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
  replace('splice', (original) => function (this: unknown[], ...args: unknown[]) {
    const first = this[0];
    if (first !== null && typeof first === 'object' && 'anchorIndex' in first && 'frame' in first) {
      const start = Number(args[0]);
      const removed = args.length > 1 ? Number(args[1]) : this.length - start;
      counts.drawingShifts += Math.max(0, this.length - start - removed);
    }
    return original.call(this, ...(args as never[]));
  });
  for (const name of ['slice', 'concat']) {
    replace(name, (original) => function (this: unknown[], ...args: unknown[]) {
      const copy = original.call(this, ...(args as never[])) as unknown[];
      for (let index = 0; index < copy.length; index += 1) visit(copy[index]);
      return copy;
    });
  }
  const stringify = JSON.stringify;
  JSON.stringify = ((value: unknown, ...rest: never[]) => {
    if (Array.isArray(value) && Array.isArray(value[0]) && typeof value[0][0] === 'string'
      && value[0][1] !== null && typeof value[0][1] === 'object' && 'xPt' in value[0][1]) {
      counts.frameKeys += value.length;
    }
    return stringify(value, ...rest);
  }) as typeof JSON.stringify;
  restore.push(() => { JSON.stringify = stringify; });
  if (setIteration) {
    const setIterator = Set.prototype[Symbol.iterator];
    Set.prototype[Symbol.iterator] = function (this: Set<unknown>) {
      const iterator = setIterator.call(this);
      return (function* () {
        for (let next = iterator.next(); !next.done; next = iterator.next()) {
          if (typeof next.value === 'string' && next.value.includes('body:body:') && next.value.includes('|')) {
            counts.pinKeys += 1;
          }
          yield next.value;
        }
      })() as never;
    };
    restore.push(() => { Set.prototype[Symbol.iterator] = setIterator; });
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
    // No paragraph here anchors a drawing: none reads the history of placed
    // table boxes.
    expect(large.history).toBe(0);
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

  it('pins a run raised over the first column once, however many tables it holds', () => {
    // A paragraph and a column break, then one-row tables with nothing
    // between them ending with an 800 pt one: the run is raised over column 1
    // and the page laid out once more. The tables after the first are already
    // pinned then and ask for nothing further.
    const columnBreak = '<w:p><w:r><w:br w:type="column"/></w:r></w:p>';
    for (const count of [4, 8, 16]) {
      const before = work.registered;
      const pages = layOutPages('tbLrV', [
        P, columnBreak, ...Array.from({ length: count - 1 }, () => table(1, 18)), table(1, 800),
      ], 2);
      expect(pages).toHaveLength(1);
      expect(pages[0]!.tables).toHaveLength(count);
      // One box when the strip before the run plans it, one when the page
      // starts again with the pin.
      expect(work.registered - before, `${count} tables`).toBeLessThanOrEqual(2);
    }
  }, 120_000);

  it('reads the tables of a run that spans pages in proportion to their number', () => {
    // One 400 pt row each: one table a page. The run's dimensions are taken
    // once, not again from every page's first table.
    const measure = (count: number) => {
      const before = work.recordReads;
      const pages = layOutPages('tbLrV', Array.from({ length: count }, () => table(1, 400)));
      expect(pages).toHaveLength(count);
      return work.recordReads - before;
    };
    const small = measure(20);
    const large = measure(80);
    expect(large).toBeLessThanOrEqual(5 * small);
  }, 120_000);

  it('lays out a run whose tables measure differently on every page in proportion to its tables', () => {
    // One table a page (a 400 pt row), each holding 37 roman PAGE fields: its
    // text grows with the page number, so every table measures differently
    // on the page it lands on than when the run was taken. Only the table
    // placed is measured again, never the run.
    const measure = (count: number) => {
      const before = work.recordReads;
      const pages = layOutPages('tbRl', Array.from({ length: count }, () => romanPageTable(37, 400)));
      expect(pages.flatMap((page) => page.tables).length).toBeGreaterThanOrEqual(count);
      return work.recordReads - before;
    };
    const small = measure(10);
    const large = measure(20);
    expect(large).toBeLessThanOrEqual(2.6 * small);
  }, 300_000);

  it('registers the pins one layout of a page finds together', () => {
    // Tables one strip apart, each followed by a paragraph anchoring a small
    // wrapped drawing over it: every table has to be pinned for its drawing
    // to lie below it. The page is laid out again once for all of them, not
    // once per table.
    const measure = (count: number) => {
      const body = Array.from({ length: count }, (_, at) => [
        table(1), P.replace('</w:p>', `${anchoredRectangle(60 + 18 * at, 62 + 11.25 * at, 6, 5)}</w:p>`),
      ]).flat();
      const before = work.registered;
      const pages = layOutPages('tbLrV', body);
      expect(pages).toHaveLength(1);
      expect(pages[0]!.tables).toHaveLength(count);
      expect(pages[0]!.drawings).toHaveLength(count);
      return work.registered - before;
    };
    const small = measure(8);
    const large = measure(16);
    expect(large).toBeLessThanOrEqual(3 * small);
  }, 300_000);

  it('lays a page out again once for all the drawings their paragraphs place elsewhere', () => {
    // Pairs of paragraphs after a table: the first anchors a rectangle
    // positioned down its paragraph, the second one positioned on the page
    // that may not overlap and lies over the first. Each second rectangle
    // is registered when the page starts, before the first is known, and
    // its paragraph then places it below the first. The page is laid out
    // again once for all pairs, and each rectangle is placed against the
    // rectangles near it: neither the layouts of the page nor the drawing
    // records read per pair grow with the number of pairs.
    const measure = (pairs: number) => {
      const body = [table(2), ...Array.from({ length: pairs }, (_, at) => [
        P.replace('</w:p>', `${anchoredRectangle(70 + 30 * at, 310, 10, 20, 'paragraph')}</w:p>`),
        P.replace('</w:p>', `${anchoredRectangle(70 + 30 * at, 380, 10, 20, 'page', 'page', false)}</w:p>`),
      ]).flat()];
      const before = work.pageIndexes;
      const { result, counts } = counting(() => layOutPages('tbLrV', body, 1, 240));
      expect(result).toHaveLength(1);
      expect(result[0]!.drawings).toHaveLength(2 * pairs);
      // Every second rectangle lies directly below the first of its pair.
      for (let at = 0; at < pairs; at += 1) {
        const [first, second] = [result[0]!.drawings[2 * at]!, result[0]!.drawings[2 * at + 1]!];
        expect(second.left).toBeCloseTo(first.left, 6);
        expect(second.top).toBeCloseTo(first.top + first.height, 6);
      }
      return { layouts: work.pageIndexes - before, drawings: counts.drawings };
    };
    const small = measure(4);
    const large = measure(8);
    expect(large.layouts).toBe(small.layouts);
    expect(large.drawings).toBeLessThanOrEqual(2 * small.drawings + 16);
  }, 300_000);

  it('keys a paragraph on the pinned frames of its own drawings', () => {
    // The pairs of the test above: after the page is laid out again every
    // page-positioned rectangle is pinned. Each anchoring paragraph is keyed
    // on the pins of its own rectangles, not on every pin of the page.
    const measure = (pairs: number) => {
      const body = [table(2), ...Array.from({ length: pairs }, (_, at) => [
        P.replace('</w:p>', `${anchoredRectangle(70 + 30 * at, 310, 10, 20, 'paragraph')}</w:p>`),
        P.replace('</w:p>', `${anchoredRectangle(70 + 30 * at, 380, 10, 20, 'page', 'page', false)}</w:p>`),
      ]).flat()];
      const { result, counts } = counting(() => layOutPages('tbLrV', body, 1, 240));
      expect(result[0]!.drawings).toHaveLength(2 * pairs);
      return counts.frameKeys;
    };
    const small = measure(4);
    const large = measure(8);
    expect(small).toBeGreaterThan(0);
    expect(large).toBeLessThanOrEqual(2 * small);
  }, 300_000);

  it('lays a page out again for its drawings without reading the pins asked for on earlier pages', () => {
    // Pages that each hold a table pinned for the rectangle anchored after
    // it, and a pair as above, for which the page is laid out again and its
    // table pins are found anew. What was asked for on the pages before is
    // not read.
    const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
    const measure = (pages: number) => {
      const body = Array.from({ length: pages }, () => [
        table(1),
        P.replace('</w:p>', `${anchoredRectangle(60, 62, 6, 5)}</w:p>`),
        P.replace('</w:p>', `${anchoredRectangle(300, 310, 10, 20, 'paragraph')}</w:p>`),
        P.replace('</w:p>', `${anchoredRectangle(300, 380, 10, 20, 'page', 'page', false)}</w:p>`),
        pageBreak,
      ]).flat();
      const { result, counts } = counting(() => layOutPages('tbLrV', body), true);
      expect(result.length).toBeGreaterThanOrEqual(pages);
      for (let page = 0; page < pages; page += 1) {
        expect(result[page]!.tables).toHaveLength(1);
        expect(result[page]!.drawings).toHaveLength(3);
      }
      return counts.pinKeys;
    };
    expect(measure(8)).toBe(measure(4));
  }, 300_000);

  it('compares each drawing registered at the page start with its own registration only', () => {
    // A table, then paragraphs each anchoring a small rectangle positioned
    // on the page, clear of one another: every rectangle is registered when
    // the page starts and none moves. A paragraph looks up what its own
    // drawing was registered with; it does not read the page's registry.
    const measure = (count: number) => {
      const body = [table(2), ...Array.from({ length: count }, (_, at) =>
        P.replace('</w:p>', `${anchoredRectangle(45 + 30 * at, 500, 12, 18)}</w:p>`))];
      const { result, counts } = counting(() => layOutPages('tbLrV', body));
      expect(result).toHaveLength(1);
      expect(result[0]!.drawings).toHaveLength(count);
      return counts.floats;
    };
    const small = measure(8);
    const large = measure(16);
    expect(large).toBeLessThanOrEqual(2 * small + 16);
  }, 300_000);

  it('reads no placed table for a drawing that lies clear of the tables', () => {
    // Paragraphs anchoring a drawing at the far block edge of the page,
    // alternating with tables near its start: the drawing's block range
    // holds no table box, so placing it reads none.
    for (const count of [4, 8]) {
      const body = Array.from({ length: count }, () => [
        table(1), P.replace('</w:p>', `${anchoredRectangle(0, 66, 80, 40)}</w:p>`),
      ]).flat();
      const { result, counts } = counting(() => layOutPages('tbRl', body, 2, 240));
      expect(result.flatMap((page) => page.drawings)).toHaveLength(count);
      expect(counts.history, `${count} pairs`).toBe(0);
    }
  }, 300_000);

  it('continues a logical table over pages without acquiring its members again on each', () => {
    // A default table style makes the tables one logical table; one 400 pt
    // row each puts one of them on a page. Its members hold nothing
    // page-dependent, so the table is acquired once, not once per page.
    const styles = '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style>';
    const measure = (count: number) => {
      const body = Array.from({ length: count }, () => table(1, 400));
      const { result, counts } = counting(() => layOutPages('tbRl', body, 1, 360, undefined, styles));
      expect(result).toHaveLength(count);
      return counts.members;
    };
    const small = measure(8);
    const large = measure(16);
    expect(small).toBeGreaterThan(0);
    expect(large).toBeLessThanOrEqual(2.6 * small);
  }, 300_000);

  it('places tables without reading the drawings that lie clear of them', () => {
    // Tables alternating with paragraphs that anchor a small wrapped drawing
    // in the page margin, far from every table along the block axis: placing
    // a table reads the drawings near its block range only, so none here.
    const measure = (count: number) => {
      const body = Array.from({ length: count }, () => [
        table(1), P.replace('</w:p>', `${anchoredRectangle(2, 2, 5, 5)}</w:p>`),
      ]).flat();
      const { result, counts } = counting(() => layOutPages('tbRl', body, 2, 240));
      expect(result.flatMap((page) => page.tables)).toHaveLength(count);
      expect(result.flatMap((page) => page.drawings)).toHaveLength(count);
      return counts.drawings;
    };
    const small = measure(4);
    const large = measure(8);
    expect(large).toBeLessThanOrEqual(Math.max(8, 2.6 * small));
  }, 300_000);
  it('continues a page-dependent logical table over pages acquiring only the members each page takes', () => {
    // As above, but every member's cell holds a PAGE field: its members are
    // page-dependent. A continuation page reads the rows it takes only, not
    // every member of the table again.
    const styles = '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style>';
    const measure = (count: number) => {
      const body = Array.from({ length: count }, () => romanPageTable(1, 400, 'Arabic'));
      const { result, counts } = counting(() => layOutPages('tbRl', body, 1, 360, undefined, styles));
      expect(result).toHaveLength(count);
      return counts.members;
    };
    const small = measure(8);
    const medium = measure(16);
    const large = measure(32);
    expect(small).toBeGreaterThan(0);
    expect(large).toBeLessThanOrEqual(4.5 * small);
    expect(large).toBeLessThanOrEqual(2.5 * medium);
  }, 300_000);

  it('places tables without reading the drawings that lie before the column top', () => {
    // Paragraphs anchoring 340 x 1 pt rules in the top margin alternate with
    // one-row tables: the rules share the tables' block range but end before
    // the column top, so no table placement reads them.
    const measure = (count: number) => {
      const body = Array.from({ length: count }, (_, at) => [
        table(1), P.replace('</w:p>', `${anchoredRectangle(41, 4 + 2 * at, 340, 1)}</w:p>`),
      ]).flat();
      const { result, counts } = counting(() => layOutPages('tbLrV', body, 2, 240));
      expect(result.flatMap((page) => page.tables)).toHaveLength(count);
      expect(result.flatMap((page) => page.drawings)).toHaveLength(count);
      return counts.drawings;
    };
    const small = measure(8);
    const medium = measure(16);
    const large = measure(24);
    expect(large).toBeLessThanOrEqual(Math.max(8, 3.5 * small));
    expect(large).toBeLessThanOrEqual(Math.max(8, 1.8 * medium));
  }, 300_000);
  it('places tables without reading the drawings that lie below them along the column', () => {
    // The same rules near the column end, far below the one-row tables at
    // the column top: they share the tables' block range but lie beyond
    // their inline extent.
    const measure = (count: number) => {
      const body = Array.from({ length: count }, (_, at) => [
        table(1), P.replace('</w:p>', `${anchoredRectangle(41, 695 + 2 * at, 340, 1)}</w:p>`),
      ]).flat();
      const { result, counts } = counting(() => layOutPages('tbLrV', body, 2, 240));
      expect(result.flatMap((page) => page.tables)).toHaveLength(count);
      expect(result.flatMap((page) => page.drawings)).toHaveLength(count);
      return counts.drawings;
    };
    const small = measure(8);
    const medium = measure(16);
    const large = measure(24);
    expect(large).toBeLessThanOrEqual(Math.max(8, 3.5 * small));
    expect(large).toBeLessThanOrEqual(Math.max(8, 1.8 * medium));
  }, 300_000);
  it('lays a page out without its pins when a layout that holds them fails', () => {
    // A long table's first fragment covers a rectangle anchored after it and
    // is pinned; every placement in a layout holding that pin fails here.
    // The page is laid out once more without pins instead of failing.
    work.failPinnedLayouts = true;
    work.failures = 0;
    try {
      const pages = layOutPages('tbRl', [
        table(1), table(70), P.replace('</w:p>', `${anchoredRectangle(440, 480, 80, 10)}</w:p>`), P,
      ]);
      expect(work.failures).toBe(1);
      expect(pages.flatMap((page) => page.tables).length).toBeGreaterThanOrEqual(2);
    } finally {
      work.failPinnedLayouts = false;
    }
  }, 120_000);
  it('records drawings ending at the same inline position without shifting the earlier ones', () => {
    // Narrow rectangles side by side along the block axis at one inline
    // position share their cells and their inline end: each is recorded
    // after the others.
    const measure = (count: number) => {
      const body = [P.replace('</w:p>', `${Array.from({ length: count }, (_, at) =>
        anchoredRectangle(382 + at, 700, 0.5, 5)).join('')}</w:p>`), T, P];
      const { result, counts } = counting(() => layOutPages('tbRl', body));
      expect(result.flatMap((page) => page.drawings)).toHaveLength(count);
      return counts.drawingShifts;
    };
    const small = measure(16);
    const large = measure(64);
    expect(large).toBeLessThanOrEqual(Math.max(16, 4.5 * small));
  }, 300_000);
  it('moves drawings ending at the same inline position below a table without shifting the others', () => {
    // Narrow rectangles side by side over a table, anchored after it: all
    // move below the table, from cells they share, to cells they share.
    const measure = (count: number) => {
      const body = [P, T, P.replace('</w:p>', `${Array.from({ length: count }, (_, at) =>
        anchoredRectangle(383 + 0.5 * at, 62, 0.45, 5)).join('')}</w:p>`), P];
      const { result, counts } = counting(() => layOutPages('tbRl', body));
      const [page] = result;
      expect(page!.drawings).toHaveLength(count);
      for (const drawing of page!.drawings) expect(drawing.top).toBeCloseTo(page!.tables[0]!.bottom, 6);
      return counts.drawingShifts;
    };
    const small = measure(16);
    const large = measure(128);
    expect(large).toBeLessThanOrEqual(Math.max(16, 9 * small));
  }, 300_000);
});

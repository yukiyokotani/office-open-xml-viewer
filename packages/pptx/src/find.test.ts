import { describe, it, expect } from 'vitest';
import { PptxFindController } from './find.js';
import type { PptxTextRunInfo } from './renderer';

/**
 * IX2 pptx find controller. Exercised with stubbed per-slide runs: joins a
 * slide's runs, matches across run boundaries, aggregates in document order
 * tagged `{ slide }`, and cycles the active match across slides.
 */
function run(text: string, overrides: Partial<PptxTextRunInfo> = {}): PptxTextRunInfo {
  return {
    text,
    inShapeX: 0,
    inShapeY: 0,
    w: text.length,
    h: 10,
    fontSize: 10,
    font: '10px monospace',
    shapeX: 0,
    shapeY: 0,
    shapeW: 100,
    shapeH: 20,
    rotation: 0,
    ...overrides,
  };
}

function controllerFor(slides: PptxTextRunInfo[][]): PptxFindController {
  return new PptxFindController(
    () => slides.length,
    (slide) => Promise.resolve(slides[slide] ?? []),
  );
}

describe('PptxFindController.find', () => {
  it('finds matches across slides tagged with their slide index', async () => {
    const c = controllerFor([[run('hello world')], [run('a world here')]]);
    const matches = await c.find('world');
    expect(matches).toHaveLength(2);
    expect(matches[0].location.slide).toBe(0);
    expect(matches[1].location.slide).toBe(1);
  });

  it('resolves a match straddling two runs on one slide', async () => {
    const c = controllerFor([[run('Hel'), run('lo there')]]);
    const matches = await c.find('Hello');
    expect(matches).toHaveLength(1);
    expect(matches[0].text).toBe('Hello');
  });

  it('matches across runs in one table cell but never across cell boundaries', async () => {
    const sameCell = await controllerFor([[
      run('Hel', { elementIndex: 0, tableCell: { row: 0, column: 0 } }),
      run('lo', { elementIndex: 0, tableCell: { row: 0, column: 0 } }),
    ]]).find('Hello');
    expect(sameCell).toHaveLength(1);

    const adjacentCells = await controllerFor([[
      run('foo', { elementIndex: 0, tableCell: { row: 0, column: 0 } }),
      run('bar', { elementIndex: 0, tableCell: { row: 0, column: 1 } }),
    ]]).find('foobar');
    expect(adjacentCells).toEqual([]);
  });

  it('is case-insensitive by default; caseSensitive honored', async () => {
    const ci = await controllerFor([[run('FOO foo')]]).find('foo');
    expect(ci).toHaveLength(2);
    const cs = await controllerFor([[run('FOO foo')]]).find('foo', { caseSensitive: true });
    expect(cs).toHaveLength(1);
  });
});

describe('PptxFindController cursor + highlights', () => {
  it('suppresses a rejected collection after it becomes stale', async () => {
    let rejectRuns!: (reason: Error) => void;
    const c = new PptxFindController(
      () => 1,
      () => new Promise((_resolve, reject) => { rejectRuns = reject; }),
    );
    const pending = c.find('a');

    c.invalidate();
    rejectRuns(new Error('old presentation closed'));

    await expect(pending).resolves.toEqual([]);
  });

  it('propagates a rejected collection for the current query', async () => {
    const c = new PptxFindController(
      () => 1,
      () => Promise.reject(new Error('current presentation failed')),
    );
    await expect(c.find('a')).rejects.toThrow('current presentation failed');
  });

  it('cycles the active match with wrap-around across slides', async () => {
    const c = controllerFor([[run('x')], [run('x')]]);
    await c.find('x');
    expect(c.next()?.matchIndex).toBe(0);
    expect(c.activeSlide()).toBe(0);
    expect(c.next()?.matchIndex).toBe(1);
    expect(c.activeSlide()).toBe(1);
    expect(c.next()?.matchIndex).toBe(0); // wrap
  });

  it('slideHighlights scopes to slide and marks active', async () => {
    const c = controllerFor([[run('a a')]]);
    await c.find('a');
    c.next();
    const hl = c.slideHighlights(0);
    expect(hl).toHaveLength(2);
    expect(hl[0].active).toBe(true);
    expect(hl[1].active).toBe(false);
  });

  it('invalidate clears everything', async () => {
    const c = controllerFor([[run('a')]]);
    await c.find('a');
    c.invalidate();
    expect(c.matches()).toHaveLength(0);
    expect(c.activeSlide()).toBeNull();
  });

  it('prevents a pending find from restoring matches after invalidate', async () => {
    let resolveRuns!: (runs: PptxTextRunInfo[]) => void;
    const c = new PptxFindController(
      () => 1,
      () => new Promise((resolve) => { resolveRuns = resolve; }),
    );
    const pending = c.find('a');

    c.invalidate();
    resolveRuns([run('a')]);

    await expect(pending).resolves.toEqual([]);
    expect(c.matches()).toEqual([]);
    expect(c.slideRuns(0)).toBeUndefined();
    expect(c.next()).toBeNull();
  });

  it('commits collected slide runs atomically only after the complete scan', async () => {
    let resolveSecond!: (runs: PptxTextRunInfo[]) => void;
    const c = new PptxFindController(
      () => 2,
      (slide) => slide === 0
        ? Promise.resolve([run('a')])
        : new Promise((resolve) => { resolveSecond = resolve; }),
    );
    const pending = c.find('a');
    await Promise.resolve();
    await Promise.resolve();

    expect(c.slideRuns(0)).toBeUndefined();
    resolveSecond([run('a')]);
    await expect(pending).resolves.toHaveLength(2);
    expect(c.slideRuns(0)).toEqual([run('a')]);
    expect(c.slideRuns(1)).toEqual([run('a')]);
  });

  it('does not overwrite newer visible-render geometry when a pending scan completes', async () => {
    let resolveSecond!: (runs: PptxTextRunInfo[]) => void;
    const oldVisibleRuns = [{ ...run('a'), inShapeX: 1 }];
    const freshVisibleRuns = [{ ...run('a'), inShapeX: 20 }];
    const c = new PptxFindController(
      () => 2,
      (slide) => slide === 0
        ? Promise.resolve(oldVisibleRuns)
        : new Promise((resolve) => { resolveSecond = resolve; }),
    );
    c.setSlideRuns(0, oldVisibleRuns);
    const pending = c.find('a');
    await Promise.resolve();

    // A zoom-settle render publishes fresh geometry while the full-deck scan
    // is still waiting on another slide.
    c.setSlideRuns(0, freshVisibleRuns);
    resolveSecond([run('a')]);

    await expect(pending).resolves.toHaveLength(2);
    expect(c.slideRuns(0)).toBe(freshVisibleRuns);
  });
});

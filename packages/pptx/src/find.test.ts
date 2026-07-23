import { describe, it, expect } from 'vitest';
import { PptxFindController } from './find.js';
import type { PptxTextRunInfo } from './renderer';

/**
 * IX2 pptx find controller. Exercised with stubbed per-slide runs: joins a
 * slide's runs, matches across run boundaries, aggregates in document order
 * tagged `{ slide }`, and cycles the active match across slides.
 */
function run(text: string): PptxTextRunInfo {
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

  it('is case-insensitive by default; caseSensitive honored', async () => {
    const ci = await controllerFor([[run('FOO foo')]]).find('foo');
    expect(ci).toHaveLength(2);
    const cs = await controllerFor([[run('FOO foo')]]).find('foo', { caseSensitive: true });
    expect(cs).toHaveLength(1);
  });
});

describe('PptxFindController cursor + highlights', () => {
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
});

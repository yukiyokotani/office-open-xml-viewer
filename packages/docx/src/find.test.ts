import { describe, it, expect } from 'vitest';
import { DocxFindController } from './find.js';
import type { DocxTextRunInfo } from './renderer';

/**
 * IX2 docx find controller. Exercised with stubbed per-page runs (no real
 * render): the controller must join a page's runs, match the query across run
 * boundaries, aggregate matches in document order tagged with `{ page }`, and
 * cycle the active match with wrap-around across pages.
 */
function run(text: string): DocxTextRunInfo {
  return { text, x: 0, y: 0, w: text.length, h: 10, fontSize: 10, font: '10px monospace' };
}

/** Build a controller over fixed per-page run lists. */
function controllerFor(pages: DocxTextRunInfo[][]): DocxFindController {
  return new DocxFindController(
    () => pages.length,
    (page) => Promise.resolve(pages[page] ?? []),
  );
}

describe('DocxFindController.find', () => {
  it('searches several terms at once, and treats an all-empty list as a clear', async () => {
    const c = controllerFor([[run('the cat sat')], [run('a concatenated dog')]]);
    const matches = await c.find(['dog', 'cat'], { wholeWord: true });
    expect(matches.map((m) => [m.location.page, m.text])).toEqual([[0, 'cat'], [1, 'dog']]);
    expect(await c.find([''])).toEqual([]);
    expect(c.matches()).toEqual([]);
  });

  it("carries each term's colour into the matches and the highlights", async () => {
    const c = controllerFor([[run('the cat and the dog')]]);
    const matches = await c.find([{ text: 'dog', color: 'red' }, 'cat']);
    expect(matches.map((m) => [m.text, m.color])).toEqual([['cat', undefined], ['dog', 'red']]);
    expect(c.pageHighlights(0).map((h) => h.color)).toEqual([undefined, 'red']);
    expect(c.next()?.color).toBeUndefined();
    expect(c.next()?.color).toBe('red');
  });

  it('finds matches across pages and tags each with its page', async () => {
    const c = controllerFor([
      [run('hello world')],
      [run('the world is round')],
    ]);
    const matches = await c.find('world');
    expect(matches).toHaveLength(2);
    expect(matches[0].location.page).toBe(0);
    expect(matches[1].location.page).toBe(1);
    expect(matches.map((m) => m.matchIndex)).toEqual([0, 1]);
  });

  it('resolves a match that straddles two runs on one page', async () => {
    const c = controllerFor([[run('Hel'), run('lo World')]]);
    const matches = await c.find('Hello');
    expect(matches).toHaveLength(1);
    expect(matches[0].location.page).toBe(0);
    // The reported text carries the document's original case.
    expect(matches[0].text).toBe('Hello');
  });

  it('matches the semantic host-to-text-box seam, not the former z-order seam', async () => {
    // B4 sequence: header → host body text → anchored text box → note → footer.
    // Presentation stacking may paint the text box behind the host, but must not
    // reorder the search index.
    const c = controllerFor([[
      run('header '),
      run('host'),
      run(' textbox'),
      run(' note'),
      run(' footer'),
    ]]);
    await expect(c.find('host textbox')).resolves.toHaveLength(1);
    await expect(c.find('textboxhost')).resolves.toHaveLength(0);
  });

  it('is case-insensitive by default and reports original-case text', async () => {
    const c = controllerFor([[run('The FOO and foo')]]);
    const matches = await c.find('foo');
    expect(matches).toHaveLength(2);
    expect(matches.map((m) => m.text)).toEqual(['FOO', 'foo']);
  });

  it('honors caseSensitive: true', async () => {
    const c = controllerFor([[run('FOO foo')]]);
    const matches = await c.find('foo', { caseSensitive: true });
    expect(matches).toHaveLength(1);
    expect(matches[0].text).toBe('foo');
  });

  it('returns [] for an empty query and clears state', async () => {
    const c = controllerFor([[run('anything')]]);
    await c.find('any');
    const cleared = await c.find('');
    expect(cleared).toEqual([]);
  });
});

describe('DocxFindController active-match cursor', () => {
  it('next() lands on the first match, then advances with wrap-around', async () => {
    const c = controllerFor([[run('a a a')]]); // three 'a' matches
    await c.find('a');
    expect(c.next()?.matchIndex).toBe(0);
    expect(c.next()?.matchIndex).toBe(1);
    expect(c.next()?.matchIndex).toBe(2);
    expect(c.next()?.matchIndex).toBe(0); // wrap
  });

  it('prev() from no-active lands on the last match', async () => {
    const c = controllerFor([[run('a a')]]);
    await c.find('a');
    expect(c.prev()?.matchIndex).toBe(1);
  });

  it('activePage() tracks the active match across pages', async () => {
    const c = controllerFor([[run('x')], [run('x')]]);
    await c.find('x');
    c.next();
    expect(c.activePage()).toBe(0);
    c.next();
    expect(c.activePage()).toBe(1);
  });

  it('next()/prev() return null when there are no matches', async () => {
    const c = controllerFor([[run('abc')]]);
    await c.find('zzz');
    expect(c.next()).toBeNull();
    expect(c.prev()).toBeNull();
    expect(c.activePage()).toBeNull();
  });
});

describe('DocxFindController.pageHighlights', () => {
  it('returns per-page slices and marks the active match', async () => {
    const c = controllerFor([[run('a a')]]);
    await c.find('a');
    c.next(); // active = match 0
    const hl = c.pageHighlights(0);
    expect(hl).toHaveLength(2);
    expect(hl[0].active).toBe(true);
    expect(hl[1].active).toBe(false);
    // Each highlight carries the run-slice(s) it covers.
    expect(hl[0].slices[0]).toMatchObject({ runIndex: 0, start: 0, end: 1 });
  });

  it('scopes highlights to the requested page', async () => {
    const c = controllerFor([[run('a')], [run('a')]]);
    await c.find('a');
    expect(c.pageHighlights(0)).toHaveLength(1);
    expect(c.pageHighlights(1)).toHaveLength(1);
    expect(c.pageHighlights(2)).toHaveLength(0);
  });
});

describe('DocxFindController.invalidate', () => {
  it('suppresses a rejected collection after it becomes stale', async () => {
    let rejectRuns!: (reason: Error) => void;
    const c = new DocxFindController(
      () => 1,
      () => new Promise((_resolve, reject) => { rejectRuns = reject; }),
    );
    const pending = c.find('a');

    c.invalidate();
    rejectRuns(new Error('old document closed'));

    await expect(pending).resolves.toEqual([]);
  });

  it('propagates a rejected collection for the current query', async () => {
    const c = new DocxFindController(
      () => 1,
      () => Promise.reject(new Error('current document failed')),
    );
    await expect(c.find('a')).rejects.toThrow('current document failed');
  });

  it('drops matches and cached runs', async () => {
    const c = controllerFor([[run('a')]]);
    await c.find('a');
    c.invalidate();
    expect(c.matches()).toHaveLength(0);
    expect(c.pageHighlights(0)).toHaveLength(0);
    expect(c.activePage()).toBeNull();
  });

  it('prevents a pending find from restoring matches after invalidate', async () => {
    let resolveRuns!: (runs: DocxTextRunInfo[]) => void;
    const c = new DocxFindController(
      () => 1,
      () => new Promise((resolve) => { resolveRuns = resolve; }),
    );
    const pending = c.find('a');

    c.invalidate();
    resolveRuns([run('a')]);

    await expect(pending).resolves.toEqual([]);
    expect(c.matches()).toEqual([]);
    expect(c.pageRuns(0)).toBeUndefined();
    expect(c.next()).toBeNull();
  });

  it('commits collected page runs atomically only after the complete scan', async () => {
    let resolveSecond!: (runs: DocxTextRunInfo[]) => void;
    const c = new DocxFindController(
      () => 2,
      (page) => page === 0
        ? Promise.resolve([run('a')])
        : new Promise((resolve) => { resolveSecond = resolve; }),
    );
    const pending = c.find('a');
    await Promise.resolve();
    await Promise.resolve();

    expect(c.pageRuns(0)).toBeUndefined();
    resolveSecond([run('a')]);
    await expect(pending).resolves.toHaveLength(2);
    expect(c.pageRuns(0)).toEqual([run('a')]);
    expect(c.pageRuns(1)).toEqual([run('a')]);
  });

  it('does not overwrite newer visible-render geometry when a pending scan completes', async () => {
    let resolveSecond!: (runs: DocxTextRunInfo[]) => void;
    const oldVisibleRuns = [{ ...run('a'), x: 1 }];
    const freshVisibleRuns = [{ ...run('a'), x: 20 }];
    const c = new DocxFindController(
      () => 2,
      (page) => page === 0
        ? Promise.resolve(oldVisibleRuns)
        : new Promise((resolve) => { resolveSecond = resolve; }),
    );
    c.setPageRuns(0, oldVisibleRuns);
    const pending = c.find('a');
    await Promise.resolve();

    // A zoom-settle render publishes fresh geometry while the full-document
    // scan is still waiting on another page.
    c.setPageRuns(0, freshVisibleRuns);
    resolveSecond([run('a')]);

    await expect(pending).resolves.toHaveLength(2);
    expect(c.pageRuns(0)).toBe(freshVisibleRuns);
  });
});

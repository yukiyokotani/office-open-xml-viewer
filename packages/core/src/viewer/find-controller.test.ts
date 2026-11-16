import { describe, expect, it, vi } from 'vitest';
import { UnitFindController } from './find-controller';

const runs = (texts: string[]) => texts.map((text) => ({ text }));
const controller = (units: string[][]) => new UnitFindController(
  () => units.length,
  async (unit: number) => runs(units[unit] ?? []),
  (unit: number) => ({ unit }),
);

describe('UnitFindController', () => {
  it('finds across runs in unit order, retains case, and cycles highlights', async () => {
    const find = controller([['Hel', 'lo'], ['hello']]);
    expect(await find.find('hello')).toEqual([
      { matchIndex: 0, text: 'Hello', location: { unit: 0 } },
      { matchIndex: 1, text: 'hello', location: { unit: 1 } },
    ]);
    expect(find.next()?.location.unit).toBe(0);
    expect(find.unitHighlights(0)).toEqual([{
      slices: [{ runIndex: 0, start: 0, end: 3 }, { runIndex: 1, start: 0, end: 2 }],
      active: true,
    }]);
    expect(find.prev()?.location.unit).toBe(1);
    expect(find.activeUnit()).toBe(1);
  });

  it('rejects stale collection errors but propagates current errors', async () => {
    let reject!: (reason: Error) => void;
    const find = new UnitFindController(
      () => 1,
      () => new Promise<{ text: string }[]>((_resolve, rejectRuns) => { reject = rejectRuns; }),
      (unit) => ({ unit }),
    );
    const pending = find.find('x');
    find.invalidate();
    reject(new Error('previous load closed'));
    await expect(pending).resolves.toEqual([]);

    const current = new UnitFindController(
      () => 1, async (): Promise<{ text: string }[]> => { throw new Error('current load failed'); },
      (unit) => ({ unit }),
    );
    await expect(current.find('x')).rejects.toThrow('current load failed');
  });

  it('does not publish a scan that resolves after invalidation', async () => {
    let resolve!: (runs: { text: string }[]) => void;
    const find = new UnitFindController(
      () => 1,
      () => new Promise<{ text: string }[]>((done) => { resolve = done; }),
      (unit) => ({ unit }),
    );
    const pending = find.find('a');
    find.invalidate();
    resolve([{ text: 'a' }]);
    await expect(pending).resolves.toEqual([]);
    expect(find.unitRuns(0)).toBeUndefined();
    expect(find.matches()).toEqual([]);
  });

  it('commits a scan atomically and keeps newer visible render geometry', async () => {
    let resolveSecond!: (runs: { text: string; x: number }[]) => void;
    const original = [{ text: 'a', x: 1 }];
    const fresh = [{ text: 'a', x: 20 }];
    const find = new UnitFindController(
      () => 2,
      (unit) => unit === 0 ? Promise.resolve(original) : new Promise<{ text: string; x: number }[]>((resolve) => { resolveSecond = resolve; }),
      (unit) => ({ unit }),
    );
    const pending = find.find('a');
    await Promise.resolve();
    await Promise.resolve();
    expect(find.unitRuns(0)).toBeUndefined();
    find.setUnitRuns(0, fresh);
    resolveSecond([{ text: 'a', x: 2 }]);
    await expect(pending).resolves.toHaveLength(2);
    expect(find.unitRuns(0)).toBe(fresh);
  });

  it('clears matches and cached runs on an empty query', async () => {
    const find = controller([['a']]);
    await find.find('a');
    await expect(find.find('')).resolves.toEqual([]);
    expect(find.matches()).toEqual([]);
    expect(find.unitRuns(0)).toBeUndefined();
    expect(find.next()).toBeNull();
  });
});

describe('UnitFindController — text highlights', () => {
  const groups = [
    { id: 'user', terms: ['revenue'], color: 'rgba(0,0,255,0.3)' },
    { id: 'validated', terms: ['guaranteed returns'] },
  ];

  it('returns every group match in document order with its group and term', async () => {
    const find = controller([['Revenue and guaranteed ', 'returns'], ['revenue']]);
    expect(await find.highlight(groups)).toEqual([
      { matchIndex: 0, text: 'Revenue', location: { unit: 0 }, group: 'user', term: 'revenue' },
      { matchIndex: 1, text: 'guaranteed returns', location: { unit: 0 }, group: 'validated', term: 'guaranteed returns' },
      { matchIndex: 2, text: 'revenue', location: { unit: 1 }, group: 'user', term: 'revenue' },
    ]);
    expect(find.highlighting).toBe(true);
  });

  it('paints the group colours first and the find matches after them', async () => {
    const find = controller([['revenue']]);
    await find.find('ven');
    await find.highlight(groups);
    expect(find.unitHighlights(0)).toEqual([
      { slices: [{ runIndex: 0, start: 0, end: 7 }], active: false, color: 'rgba(0,0,255,0.3)', layer: 0 },
      { slices: [{ runIndex: 0, start: 2, end: 5 }], active: false },
    ]);
    const validated = controller([['guaranteed returns']]);
    await validated.highlight(groups);
    expect(validated.unitHighlights(0)[0].color).toBe('rgba(0, 184, 148, 0.35)');
  });

  it('keeps the find and the highlights independent: matches, cursors and clears', async () => {
    const find = controller([['revenue revenue']]);
    await find.find('revenue');
    await find.highlight(groups);
    expect(find.next()?.matchIndex).toBe(0);
    expect(find.highlightPrev()).toMatchObject({ matchIndex: 1, group: 'user' });
    expect(find.highlightNext()).toMatchObject({ matchIndex: 0, group: 'user' });
    expect(find.unitHighlights(0).map((box) => box.active)).toEqual([true, false, true, false]);

    // A new find query resets only the find cursor.
    await find.find('rev');
    expect(find.unitHighlights(0).map((box) => box.active)).toEqual([true, false, false, false]);

    find.clearFind();
    expect(find.matches()).toEqual([]);
    expect(find.unitHighlights(0)).toHaveLength(2);
    expect(find.unitRuns(0)).toBeDefined();

    await find.find('revenue');
    find.clearHighlights();
    expect(find.highlightNext()).toBeNull();
    expect(find.highlighting).toBe(false);
    expect(find.matches()).toHaveLength(2);
    expect(find.unitRuns(0)).toBeDefined();

    // The cached runs go with the last armed search.
    find.clearFind();
    expect(find.unitRuns(0)).toBeUndefined();
  });

  it('replaces the previous set, and an empty set clears it', async () => {
    const find = controller([['alpha beta']]);
    await find.highlight([{ id: 'a', terms: ['alpha'] }]);
    expect((await find.highlight([{ id: 'b', terms: ['beta'] }])).map((match) => match.group)).toEqual(['b']);
    expect(find.unitHighlights(0)).toHaveLength(1);
    expect(await find.highlight([{ id: 'b', terms: [''] }])).toEqual([]);
    expect(find.unitHighlights(0)).toEqual([]);
    expect(find.highlighting).toBe(false);
    expect(find.unitRuns(0)).toBeUndefined();
  });

  it('rejects a repeated group id without touching the current set or a request in flight', async () => {
    const find = controller([['alpha']]);
    await find.highlight([{ id: 'a', terms: ['alpha'] }]);
    await expect(find.highlight([{ id: 'x', terms: ['a'] }, { id: 'x', terms: ['b'] }]))
      .rejects.toThrow('duplicate group id');
    // Also when the set is otherwise empty: nothing is cleared.
    await expect(find.highlight([{ id: 'x', terms: [] }, { id: 'x', terms: [''] }]))
      .rejects.toThrow('duplicate group id');
    expect(find.unitHighlights(0)).toHaveLength(1);

    const inFlight = find.highlight([{ id: 'b', terms: ['alp'] }]);
    await expect(find.highlight([{ id: 'x', terms: ['a'] }, { id: 'x', terms: ['b'] }])).rejects.toThrow(TypeError);
    expect((await inFlight).map((match) => match.text)).toEqual(['alp']);
  });

  it('paints groups in the order of the set, wherever their matches start', async () => {
    const find = controller([['abcdef']]);
    const matches = await find.highlight([
      { id: 'first', terms: ['cde'], color: 'red' },
      { id: 'second', terms: ['abcd'], color: 'blue' },
    ]);
    // Navigation follows the document; painting follows the set.
    expect(matches.map((match) => match.group)).toEqual(['second', 'first']);
    expect(find.unitHighlights(0).map((box) => box.color)).toEqual(['red', 'blue']);
    expect(find.highlightNext()).toMatchObject({ group: 'second' });
    expect(find.unitHighlights(0).map((box) => box.active)).toEqual([false, true]);
  });

  it('is not affected by later changes to the caller\'s group objects', async () => {
    let resolve!: (runs: { text: string }[]) => void;
    const find = new UnitFindController(
      () => 1,
      () => new Promise<{ text: string }[]>((done) => { resolve = done; }),
      (unit) => ({ unit }),
    );
    const terms = ['alpha'];
    const group = { id: 'a', terms, color: 'red' };
    const pending = find.highlight([group]);
    terms[0] = 'beta';
    group.color = 'blue';
    // The term search loads on demand, so collection starts a little later.
    await vi.waitFor(() => expect(resolve).toBeTypeOf('function'));
    resolve([{ text: 'alpha beta' }]);
    expect((await pending).map((match) => match.text)).toEqual(['alpha']);
    expect(find.unitHighlights(0)[0].color).toBe('red');
  });

  it('reuses the runs the other search has cached, and collects again once both are cleared', async () => {
    const collected: number[] = [];
    const find = new UnitFindController(
      () => 2,
      async (unit) => { collected.push(unit); return [{ text: 'alpha' }]; },
      (unit) => ({ unit }),
    );
    expect(await find.highlight([{ id: 'a', terms: ['alpha'] }])).toHaveLength(2);
    expect(await find.find('alpha')).toHaveLength(2);
    expect(await find.highlight([{ id: 'b', terms: ['alp'] }])).toHaveLength(2);
    expect(collected).toEqual([0, 1]);
    find.clearFind();
    find.clearHighlights();
    expect(await find.find('alpha')).toHaveLength(2);
    expect(collected).toEqual([0, 1, 0, 1]);
  });

  it('keeps a find and a highlight that collect at the same time apart', async () => {
    const gates: ((runs: { text: string }[]) => void)[] = [];
    const find = new UnitFindController(
      () => 1,
      () => new Promise<{ text: string }[]>((done) => { gates.push(done); }),
      (unit) => ({ unit }),
    );
    const found = find.find('alpha');
    const highlighted = find.highlight([{ id: 'a', terms: ['alpha'] }]);
    find.clearFind();
    await vi.waitFor(() => expect(gates).toHaveLength(2));
    gates.forEach((open) => open([{ text: 'alpha' }]));
    await expect(found).resolves.toEqual([]);
    expect(await highlighted).toHaveLength(1);
    expect(find.matches()).toEqual([]);
  });

  it('cancels a highlight scan that a newer set or a clear supersedes', async () => {
    const gates: ((runs: { text: string }[]) => void)[] = [];
    const find = new UnitFindController(
      () => 1,
      () => new Promise<{ text: string }[]>((done) => { gates.push(done); }),
      (unit) => ({ unit }),
    );
    const first = find.highlight([{ id: 'a', terms: ['alpha'] }]);
    const second = find.highlight([{ id: 'b', terms: ['beta'] }]);
    // The superseded request never starts collecting.
    await vi.waitFor(() => expect(gates).toHaveLength(1));
    gates[0]([{ text: 'alpha beta' }]);
    await expect(first).resolves.toEqual([]);
    expect((await second).map((match) => match.group)).toEqual(['b']);

    const third = find.highlight([{ id: 'c', terms: ['alpha'] }]);
    find.clearHighlights();
    await expect(third).resolves.toEqual([]);
    expect(find.unitHighlights(0)).toEqual([]);
  });

  it('drops both searches on invalidate', async () => {
    const find = controller([['alpha']]);
    await find.find('alpha');
    await find.highlight([{ id: 'a', terms: ['alpha'] }]);
    find.invalidate();
    expect(find.unitHighlights(0)).toEqual([]);
    expect(find.highlighting).toBe(false);
    expect(find.unitRuns(0)).toBeUndefined();
  });

  it('applies the adapter\'s match filter and layout breaks to highlights', async () => {
    const units = [[{ text: 'Introduction', p: 0 }, { text: 'The end', p: 1 }]];
    const find = new UnitFindController(
      () => units.length,
      async (unit: number) => units[unit],
      (unit: number) => ({ unit }),
      (unitRuns, slices) => slices.every((slice) => unitRuns[slice.runIndex].p === unitRuns[slices[0].runIndex].p),
      undefined,
      async () => (previous, next) => previous.p !== next.p,
    );
    const matches = await find.highlight([
      { id: 'word', terms: ['introduction', 'the'], wholeWord: true },
      { id: 'across', terms: ['introductionthe'] },
    ]);
    expect(matches.map((match) => [match.group, match.text])).toEqual([['word', 'Introduction'], ['word', 'The']]);
    expect(await find.find('the', { wholeWord: true })).toHaveLength(1);
  });
});

describe('UnitFindController — whole-word find', () => {
  it('matches whole words and stays cancellable while the term search loads', async () => {
    const find = controller([['concatenate the cat'], ['Cat']]);
    expect((await find.find('cat', { wholeWord: true })).map((match) => [match.location.unit, match.text]))
      .toEqual([[0, 'cat'], [1, 'Cat']]);
    const cancelled = find.find('cat', { wholeWord: true });
    find.clearFind();
    await expect(cancelled).resolves.toEqual([]);
    expect(find.matches()).toEqual([]);
  });
});

describe('UnitFindController — request results', () => {
  it('reports each find its own matches when the runs are already cached', async () => {
    const find = controller([['alpha beta']]);
    await find.find('alpha');
    // Both scans commit synchronously from the cache, one after the other.
    const [alpha, beta] = await Promise.all([find.find('alpha'), find.find('beta')]);
    expect(alpha.map((match) => match.text)).toEqual(['alpha']);
    expect(beta.map((match) => match.text)).toEqual(['beta']);
    expect(find.matches().map((match) => match.text)).toEqual(['beta']);
  });

  it('reports a failed preparation of the current request, and not of a superseded one', async () => {
    let reject!: (error: Error) => void;
    const failing = () => new Promise<never>((_resolve, rejectLoad) => { reject = rejectLoad; });
    const find = new UnitFindController(
      () => 1, async () => [{ text: 'alpha' }], (unit) => ({ unit }), undefined, undefined, failing,
    );
    const current = find.find('alpha', { wholeWord: true });
    await vi.waitFor(() => expect(reject).toBeTypeOf('function'));
    reject(new Error('chunk failed'));
    await expect(current).rejects.toThrow('chunk failed');

    const cleared = find.highlight([{ id: 'a', terms: ['alpha'], wholeWord: true }]);
    const pendingReject = reject;
    await vi.waitFor(() => expect(reject).not.toBe(pendingReject));
    find.clearHighlights();
    reject(new Error('chunk failed late'));
    await expect(cleared).resolves.toEqual([]);
    expect(find.highlighting).toBe(false);
  });
});

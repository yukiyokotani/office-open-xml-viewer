import { describe, expect, it } from 'vitest';
import { deepFreezePlainData, sealPlainData, snapshotPlainData } from './plain-data.js';

describe('plain layout data snapshots', () => {
  it('preserves signed unbounded finite DrawingML source-rectangle percentages exactly', () => {
    const authored = { l: -0.25, t: 1.25, r: 1.5, b: -0.75 };
    const source = { srcRect: { ...authored } };
    const snapshot = snapshotPlainData(source, 'paint resource');
    source.srcRect.l = 0;

    expect(snapshot.srcRect).toEqual(authored);
    expect(structuredClone(snapshot).srcRect).toEqual(authored);
  });

  it('clones and deeply freezes plain data while preserving optional undefined values', () => {
    const source = { optional: undefined, nested: { values: [1, 'two'] } };
    const snapshot = snapshotPlainData(source, 'layout payload');

    source.nested.values[0] = 99;

    expect(snapshot).toEqual({ optional: undefined, nested: { values: [1, 'two'] } });
    expect(Object.isFrozen(snapshot)).toBe(true);
    expect(Object.isFrozen(snapshot.nested)).toBe(true);
    expect(Object.isFrozen(snapshot.nested.values)).toBe(true);
  });

  it.each([
    [() => undefined],
    [Symbol('invalid')],
    [1n],
    [new Map([['key', 'value']])],
  ])('rejects non-plain data %#', (invalid) => {
    expect(() => snapshotPlainData({ invalid }, 'layout payload'))
      .toThrow(/structured-clone-safe plain data/i);
  });

  it('rejects cyclic plain data', () => {
    const cyclic: { self?: object } = {};
    cyclic.self = cyclic;

    expect(() => snapshotPlainData(cyclic, 'layout payload'))
      .toThrow(/structured-clone-safe plain data/i);
  });

  it('preserves a shared plain-data DAG', () => {
    const shared = { value: 7 };
    const snapshot = snapshotPlainData({ first: shared, second: shared }, 'layout payload');
    expect(snapshot.first).toBe(snapshot.second);
  });

  it('preserves sparse-array length and holes exactly', () => {
    const sparse = new Array<string>(4);
    sparse[2] = 'present';

    const snapshot = snapshotPlainData(sparse, 'layout payload');

    expect(snapshot).toHaveLength(4);
    expect(0 in snapshot).toBe(false);
    expect(1 in snapshot).toBe(false);
    expect(2 in snapshot).toBe(true);
    expect(3 in snapshot).toBe(false);
  });

  it('rejects proxies like structuredClone', () => {
    const proxied = new Proxy({ value: 1 }, {});

    expect(() => structuredClone(proxied)).toThrow();
    expect(() => snapshotPlainData(proxied, 'layout payload'))
      .toThrow(/structured-clone-safe plain data/i);
    expect(() => sealPlainData(proxied, 'layout payload'))
      .toThrow(/structured-clone-safe plain data/i);
  });

  it('returns the identical reference for an already-processed graph', () => {
    const snapshot = snapshotPlainData({ nested: { values: [1, 2, 3] } }, 'layout payload');

    expect(snapshotPlainData(snapshot, 'layout payload')).toBe(snapshot);
  });

  it('reuses an already-processed subtree by reference inside a fresh wrapper', () => {
    const inner = snapshotPlainData({ big: [1, 2, 3] }, 'layout payload');
    const wrapped = snapshotPlainData({ inner, extra: 'new' }, 'layout payload');

    expect(wrapped.inner).toBe(inner);
    expect(wrapped.extra).toBe('new');
    expect(Object.isFrozen(wrapped)).toBe(true);
  });

  it('still validates new data next to an already-processed subtree', () => {
    const inner = snapshotPlainData({ ok: 1 }, 'layout payload');

    expect(() => snapshotPlainData({ inner, bad: () => undefined }, 'layout payload'))
      .toThrow(/structured-clone-safe plain data/i);
  });

  it('still validates frozen graphs that were never processed', () => {
    // Frozen alone does not imply validated: deepFreezePlainData never marks a
    // graph as processed, so this graph must still be walked and rejected.
    const frozenInvalid = deepFreezePlainData({ bad: () => undefined });

    expect(() => snapshotPlainData({ frozenInvalid }, 'layout payload'))
      .toThrow(/structured-clone-safe plain data/i);
  });

  it('treats sealed builder-owned data as already processed', () => {
    const sealed = sealPlainData({ nested: { value: 1 } }, 'layout payload');

    expect(snapshotPlainData(sealed, 'layout payload')).toBe(sealed);
  });

  it('preserves an own enumerable __proto__ data property like structuredClone', () => {
    const source = Object.defineProperty({ marker: 1 }, '__proto__', {
      value: { polluted: false },
      enumerable: true,
      writable: true,
      configurable: true,
    });
    const snapshot = snapshotPlainData(source, 'layout payload') as Record<string, unknown>;

    expect(Object.prototype.hasOwnProperty.call(snapshot, '__proto__')).toBe(true);
    expect((snapshot as { __proto__: unknown }).__proto__).toEqual({ polluted: false });
    expect(Object.getPrototypeOf(snapshot)).toBe(Object.prototype);
    expect(({} as Record<string, unknown>).polluted).toBeUndefined();
  });

  it.each([
    ['accessor', (value: object) => Object.defineProperty(value, 'hidden', {
      enumerable: true, get: () => ({ retainedClosure: true }),
    })],
    ['non-enumerable', (value: object) => Object.defineProperty(value, 'hidden', {
      enumerable: false, value: 'secret',
    })],
    ['symbol', (value: object) => Object.defineProperty(value, Symbol('hidden'), {
      enumerable: true, value: 'secret',
    })],
  ])('rejects %s properties before sealing builder-owned data', (_kind, define) => {
    const value = {};
    define(value);
    expect(() => sealPlainData(value, 'layout payload'))
      .toThrow(/enumerable|string|data property/i);
  });
});

import { describe, expect, it } from 'vitest';
import {
  attachBodyLayoutKernel,
  attachDocumentLayoutRuntime,
  attachPaintResourceRegistry,
  attachPrivateResourceLookup,
  createFieldAcquisitionServicesView,
  createImmutableResourceLookup,
  createParagraphAcquisitionCacheServicesView,
  documentLayoutRuntimeOf,
  fieldAcquisitionContextOf,
  paragraphAcquisitionCacheOf,
  paintResourceRegistryOf,
  privateResourceLookupOf,
} from './runtime-state.js';
import { createPaintResourceRegistry } from './paint-resources.js';
import type { BodyLayoutKernel } from './body-layout-kernel.js';
import type { FieldAcquisitionContext } from './runtime-state.js';
import type { LayoutServices } from './types.js';

const fieldAcquisitionContractIsReduced: Exclude<
  keyof FieldAcquisitionContext,
  'totalPages' | 'resolveDestinationPage'
> extends never ? true : false = true;

function attachUnusedKernel(services: LayoutServices): void {
  attachBodyLayoutKernel(services, {
    openBodyLayoutSession() { throw new Error('unused'); },
  } as BodyLayoutKernel);
}

describe('document layout runtime state', () => {
  it('requires explicit deterministic attachment', () => {
    expect(() => documentLayoutRuntimeOf({})).toThrow(/runtime.*not initialized/i);

    const owner = {};
    attachDocumentLayoutRuntime(owner, 123);
    expect(documentLayoutRuntimeOf(owner).defaultCurrentDateMs).toBe(123);
  });

  it('hides mutable handles behind immutable, fixed membership', () => {
    const first = { id: 'first' };
    const entries = new Map<string, object>([['math:a', first]]);
    const lookup = createImmutableResourceLookup(entries);
    entries.set('math:b', { id: 'late' });

    expect(lookup.keys).toEqual(['math:a']);
    expect(Object.isFrozen(lookup.keys)).toBe(true);
    expect(lookup.resolve('math:a')).toBe(first);
    expect(() => lookup.resolve('math:b')).toThrow(/Unknown runtime resource/);
  });

  it('attaches a private lookup once and enforces exact declared membership', () => {
    const owner = {};
    const handle = { id: 'a' };
    attachPrivateResourceLookup(owner, new Map([['a', handle]]), ['a']);
    expect(privateResourceLookupOf(owner)?.resolve('a')).toBe(handle);
    expect(() => attachPrivateResourceLookup(owner, new Map([['a', handle]]), ['a']))
      .toThrow(/already attached/i);

    expect(() => attachPrivateResourceLookup({}, new Map(), ['a']))
      .toThrow(/membership.*missing/i);
    expect(() => attachPrivateResourceLookup({}, new Map([['extra', handle]]), []))
      .toThrow(/membership.*extra/i);
  });

  it('attaches one typed paint resource registry without widening its owner', () => {
    const owner = {};
    const registry = createPaintResourceRegistry([{
      kind: 'math', resourceKey: 'math:a',
    }]);

    attachPaintResourceRegistry(owner, registry);

    expect(paintResourceRegistryOf(owner)).toBe(registry);
    expect(Object.keys(owner)).toEqual([]);
    expect(() => attachPaintResourceRegistry(owner, registry)).toThrow(/already attached/i);
    expect(() => paintResourceRegistryOf({})).toThrow(/not attached/i);
  });

  it('isolates immutable field-acquisition context per service view', () => {
    const services = {
      text: {}, images: {}, math: {},
    } as unknown as LayoutServices;
    attachUnusedKernel(services);
    const first = createFieldAcquisitionServicesView(services, { totalPages: 2 });
    const second = createFieldAcquisitionServicesView(services, { totalPages: 12 });

    expect(first).not.toBe(services);
    expect(second).not.toBe(first);
    expect(first.text).toBe(services.text);
    expect(fieldAcquisitionContextOf(first)).toEqual({ totalPages: 2 });
    expect(fieldAcquisitionContextOf(second)).toEqual({ totalPages: 12 });
    expect(fieldAcquisitionContextOf(services)).toEqual({ totalPages: 1 });
    expect(fieldAcquisitionContractIsReduced).toBe(true);
    expect(Object.isFrozen(fieldAcquisitionContextOf(first))).toBe(true);
    expect(Object.isFrozen(first)).toBe(true);
    expect(Object.keys(services)).toEqual(['text', 'images', 'math']);
    expect(() => createFieldAcquisitionServicesView(services, { totalPages: 0 }))
      .toThrow(/positive integer/i);
  });

  it('shares one private paragraph cache with field views but not another pagination scope', () => {
    const services = {
      text: {}, images: {}, math: {},
    } as unknown as LayoutServices;
    attachUnusedKernel(services);

    const firstScope = createParagraphAcquisitionCacheServicesView(services);
    const fieldView = createFieldAcquisitionServicesView(firstScope, { totalPages: 12 });
    const secondScope = createParagraphAcquisitionCacheServicesView(services);

    expect(paragraphAcquisitionCacheOf(firstScope)).toBeDefined();
    expect(paragraphAcquisitionCacheOf(fieldView))
      .toBe(paragraphAcquisitionCacheOf(firstScope));
    expect(paragraphAcquisitionCacheOf(secondScope))
      .not.toBe(paragraphAcquisitionCacheOf(firstScope));
    expect(paragraphAcquisitionCacheOf(services)).toBeUndefined();
    expect(Object.keys(firstScope)).toEqual(['text', 'images', 'math']);
  });

  it('evicts least-recent paragraph candidates across sources and keys', () => {
    const services = { text: {}, images: {}, math: {} } as unknown as LayoutServices;
    attachUnusedKernel(services);
    const scoped = createParagraphAcquisitionCacheServicesView(services);
    const cache = paragraphAcquisitionCacheOf(scoped)!;
    const sources = Array.from({ length: 129 }, () => ({}));
    const first = {};
    cache.set(sources[0]!, 'original', first);
    // Replacement must not consume another slot.
    cache.set(sources[0]!, 'original', first);
    for (let index = 1; index < 128; index++) cache.set(sources[index]!, 'same-key', index);
    // Replacing an entry at full capacity must not evict another candidate.
    cache.set(sources[0]!, 'original', first);
    expect(cache.get(sources[0]!, 'original')).toBe(first);
    cache.set(sources[128]!, 'same-key', 128);
    expect(cache.get(sources[1]!, 'same-key')).toBeUndefined();
    expect(cache.get(sources[0]!, 'original')).toBe(first);
    expect(cache.get(sources[128]!, 'same-key')).toBe(128);
    expect(cache.get(sources[2]!, 'same-key')).toBe(2);
    // Many candidates for a single source count against the same global bound.
    for (let index = 0; index < 128; index++) cache.set(sources[0]!, `variant:${index}`, index);
    expect(cache.get(sources[0]!, 'original')).toBeUndefined();
    expect(cache.get(sources[128]!, 'same-key')).toBeUndefined();
    expect(cache.get(sources[0]!, 'variant:127')).toBe(127);
  });

  it('keeps destination-page resolution private to its immutable pagination iteration view', () => {
    const services = {
      text: {}, images: {}, math: {},
    } as unknown as LayoutServices;
    attachUnusedKernel(services);
    const first = createFieldAcquisitionServicesView(services, {
      totalPages: 2,
      resolveDestinationPage: (pageIndex) => pageIndex === 1
        ? { pageIndex: 1, displayPageNumber: 50, pageNumberFormat: 'upperRoman' }
        : undefined,
    });

    const context = fieldAcquisitionContextOf(first);
    expect(context.resolveDestinationPage?.(1)).toEqual({
      pageIndex: 1, displayPageNumber: 50, pageNumberFormat: 'upperRoman',
    });
    expect(context.resolveDestinationPage?.(0)).toBeUndefined();
    expect(Object.isFrozen(context)).toBe(true);
    expect(fieldAcquisitionContextOf(services)).toEqual({ totalPages: 1 });
  });

  it('rejects a field-acquisition view without a body layout owner', () => {
    const services = {
      text: {}, images: {}, math: {},
    } as unknown as LayoutServices;

    expect(() => createFieldAcquisitionServicesView(services, { totalPages: 2 }))
      .toThrow(/kernel is not attached/i);
  });
});

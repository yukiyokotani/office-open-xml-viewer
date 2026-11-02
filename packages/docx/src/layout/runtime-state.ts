import { attachNoteOverflowPolicy, noteOverflowPolicyOf, attachNoteRecoveryOwner, noteRecoveryOwnerOf } from './note-overflow-policy.js';
import type { LayoutServices } from './types.js';
import type { PaintResourceRegistry } from './types.js';
import type { KinsokuRules, NumberFormat } from '@silurus/ooxml-core';
import type { BodyLayoutKernel } from './body-layout-kernel.js';
import type { LayoutVariantStore } from './variant-store.js';
import type { VerticalGlyphMeasurementService } from './measurement-capabilities.js';
import type { LayoutSourceStore } from './layout-source-store.js';
import type { LayoutOptions } from './options.js';
import { createPaginationWorkBudget, type PaginationWorkBudget } from './pagination-work-budget.js';
import { createHeaderFooterLayoutCache } from './header-footer-layout-cache.js';
import type { StoryLayoutCache } from './story-layout-cache.js';
import { createFootnoteAcquisitionWorkBudget, type FootnoteAcquisitionWorkBudget } from './resource-budgets.js';

export interface DocumentLayoutRuntimeState {
  services: LayoutServices | null;
  readonly defaultCurrentDateMs: number;
  /**
   * The layout variant this document is currently being VIEWED as.
   *
   * `showTrackedChanges` (and an explicit `currentDate`) select a different
   * retained layout, with a different page count and different page geometry.
   * Document-level accessors used to read the default variant unconditionally,
   * which meant a tracked-changes viewer measured its scrollbar against
   * final-view geometry while painting markup-view pages — and, worse, forced a
   * whole extra synchronous pagination of the variant nobody was looking at.
   *
   * `null` means "the default view", which is what an ordinary load selects.
   */
  activeLayoutOptions: LayoutOptions | null;
}

const documentLayoutRuntime = Symbol('document-layout-runtime');

type RuntimeOwner = object & {
  [documentLayoutRuntime]?: DocumentLayoutRuntimeState;
};

export function attachDocumentLayoutRuntime(
  owner: object,
  defaultCurrentDateMs: number,
): void {
  Object.defineProperty(owner, documentLayoutRuntime, {
    configurable: false,
    enumerable: false,
    writable: false,
    value: { services: null, defaultCurrentDateMs, activeLayoutOptions: null },
  });
}

export function documentLayoutRuntimeOf(owner: object): DocumentLayoutRuntimeState {
  const runtime = (owner as RuntimeOwner)[documentLayoutRuntime];
  if (runtime) return runtime;
  throw new Error('Document layout runtime is not initialized; attach it explicitly');
}

export interface ImmutableResourceLookup<T> {
  readonly keys: readonly string[];
  resolve(resourceKey: string): T;
}

/** Keep browser/DOM handles in a private closure while exposing fixed immutable membership. */
export function createImmutableResourceLookup<T>(entries: ReadonlyMap<string, T>): ImmutableResourceLookup<T> {
  const snapshot = new Map(entries);
  const keys = Object.freeze([...snapshot.keys()].sort());
  return Object.freeze({
    keys,
    resolve(resourceKey: string): T {
      const value = snapshot.get(resourceKey);
      if (value === undefined) throw new Error(`Unknown runtime resource: ${resourceKey}`);
      return value;
    },
  });
}

const privateResourceLookups = new WeakMap<object, ImmutableResourceLookup<unknown>>();
type LayoutServicesRuntimeOwner = object;

const layoutServicesRuntimeOwners = new WeakMap<object, LayoutServicesRuntimeOwner>();
const bodyLayoutKernels = new WeakMap<LayoutServicesRuntimeOwner, BodyLayoutKernel>();
const layoutSourceStores = new WeakMap<LayoutServicesRuntimeOwner, LayoutSourceStore>();
const verticalGlyphMeasurementServices = new WeakMap<
  LayoutServicesRuntimeOwner,
  VerticalGlyphMeasurementService
>();
const layoutVariantStores = new WeakMap<LayoutServices, LayoutVariantStore>();

/** One paginateBodySteps execution owns this memo, whether a synchronous or
 * progressive driver consumes it. It never enters LayoutServices or
 * DocumentLayout; field-acquisition service views inherit the same private
 * handle through createLayoutServicesRuntimeView. */
export interface ParagraphAcquisitionRuntimeCache {
  objectIdentity(value: object): number;
  /** Session-scoped ordinal that stands for one exact service fingerprint
   * string inside acquisition keys (equal ordinals iff equal strings). */
  fingerprintOrdinal(value: string | null): number | null;
  /** Ordinal of the rule set's exact value (enabled flag plus sorted
   * forbidden code points); equal values share one ordinal. */
  kinsokuKey(rules: KinsokuRules): number;
  get(input: object, key: string): unknown;
  set(input: object, key: string, value: unknown): void;
  getLineBreaking(input: object, key: string): unknown;
  setLineBreaking(input: object, key: string, value: unknown): void;
  noteMiss(): void;
  readonly workBudget: PaginationWorkBudget;
}

/** Minimum operational allowance; source-sized pagination scopes may receive more.
 * See pagination-work-budget.ts for the aggregate cap and credit ownership. */
export { PARAGRAPH_ACQUISITION_MISS_BUDGET } from './pagination-work-budget.js';

const headerFooterLayoutCaches = new WeakMap<LayoutServices, StoryLayoutCache>();

export function headerFooterLayoutCacheOf(services: LayoutServices): StoryLayoutCache | undefined {
  return headerFooterLayoutCaches.get(services);
}

const paragraphAcquisitionCaches = new WeakMap<
  LayoutServices,
  ParagraphAcquisitionRuntimeCache
>();
const footnoteAcquisitionWorkBudgets = new WeakMap<LayoutServices, FootnoteAcquisitionWorkBudget>();

export function footnoteAcquisitionWorkBudgetOf(services: LayoutServices): FootnoteAcquisitionWorkBudget | undefined {
  return footnoteAcquisitionWorkBudgets.get(services);
}

function createParagraphAcquisitionRuntimeCache(workBudget: PaginationWorkBudget): ParagraphAcquisitionRuntimeCache {
  const identities = new WeakMap<object, number>();
  const results = new WeakMap<object, Map<string, unknown>>();
  const lineBreaks = new WeakMap<object, Map<string, unknown>>();
  const kinsokuKeys = new WeakMap<KinsokuRules, number>();
  const kinsokuOrdinals = new Map<string, number>();
  // Service fingerprints are exact canonical identities, so they grow with the
  // data they identify: the text-service fingerprint embeds the document's
  // complete font-metric snapshot (tens of KB). Spelling it into every
  // paragraph key made each retained key that large and re-allocated it on
  // every acquisition. A session sees only a handful of distinct services, and
  // this map keeps each string itself (no copy), so ordinals stay one-to-one.
  const fingerprintOrdinals = new Map<string, number>();
  let nextIdentity = 1;

  return Object.freeze({
    objectIdentity(value: object): number {
      let retained = identities.get(value);
      if (retained === undefined) {
        retained = nextIdentity;
        nextIdentity += 1;
        identities.set(value, retained);
      }
      return retained;
    },
    fingerprintOrdinal(value: string | null): number | null {
      if (value === null) return null;
      let ordinal = fingerprintOrdinals.get(value);
      if (ordinal === undefined) {
        ordinal = fingerprintOrdinals.size;
        fingerprintOrdinals.set(value, ordinal);
      }
      return ordinal;
    },
    kinsokuKey(rules: KinsokuRules) {
      let key = kinsokuKeys.get(rules);
      if (key === undefined) {
        // Document layout settings own this §17.3.1.16 / §17.15.1.58-.59
        // rule set for the whole pagination session. Its values are settled
        // before any paragraph context is made. The spelled-out value (about
        // 1 KB for the default East Asian sets) is kept once per session; keys
        // carry its value-exact ordinal.
        const value = JSON.stringify([
          rules.enabled,
          [...rules.lineStartForbidden].sort((left, right) => left - right),
          [...rules.lineEndForbidden].sort((left, right) => left - right),
        ]);
        key = kinsokuOrdinals.get(value);
        if (key === undefined) {
          key = kinsokuOrdinals.size;
          kinsokuOrdinals.set(value, key);
        }
        kinsokuKeys.set(rules, key);
      }
      return key;
    },
    get(input: object, key: string): unknown {
      const byKey = results.get(input);
      if (!byKey?.has(key)) return undefined;
      const value = byKey.get(key);
      byKey.delete(key);
      byKey.set(key, value);
      return value;
    },
    set(input: object, key: string, value: unknown): void {
      let byKey = results.get(input);
      if (!byKey) {
        byKey = new Map();
        results.set(input, byKey);
      }
      byKey.set(key, value);
      // Acquisition cache keys carry the exact placement, so every
      // keep-with-next preflight and every convergence pass adds fresh
      // keys that are almost never re-read. Values stay reachable through
      // the session-lived input objects, so without a bound the cache
      // retains every measurement of the session (unbounded memory on
      // pathological documents). Keep only the most recent placements per
      // paragraph; a miss only costs a re-measurement.
      while (byKey.size > 2) {
        byKey.delete(byKey.keys().next().value!);
      }
    },
    getLineBreaking(input: object, key: string): unknown {
      const byKey = lineBreaks.get(input);
      if (!byKey?.has(key)) return undefined;
      const value = byKey.get(key);
      byKey.delete(key);
      byKey.set(key, value);
      return value;
    },
    setLineBreaking(input: object, key: string, value: unknown): void {
      let byKey = lineBreaks.get(input);
      if (!byKey) {
        byKey = new Map();
        lineBreaks.set(input, byKey);
      }
      byKey.set(key, value);
      // Context/field variants can produce distinct line partitions for one
      // paragraph. Bound them separately from its two retained placements.
      while (byKey.size > 2) byKey.delete(byKey.keys().next().value!);
    },
    workBudget,
    noteMiss: () => workBudget.noteMiss(),
  });
}

/** Service views may replace one component, but mixing components already owned
 * by different documents must fail before acquisition can use a foreign kernel. */
function layoutServicesRuntimeOwner(
  services: LayoutServices,
  create: boolean,
): LayoutServicesRuntimeOwner | undefined {
  const components = [services.text, services.images, services.math];
  const owners = new Set(components.flatMap((component) => {
    const owner = layoutServicesRuntimeOwners.get(component);
    return owner ? [owner] : [];
  }));
  if (owners.size > 1) {
    throw new Error('Layout services combine foreign runtime owners');
  }
  const owner = owners.values().next().value as LayoutServicesRuntimeOwner | undefined;
  const unowned = components.filter((component) => !layoutServicesRuntimeOwners.has(component));
  if (owner && unowned.length > 1) {
    throw new Error('Layout services are missing service lineage for multiple components');
  }
  if (!owner && !create) return undefined;
  const resolved = owner ?? {};
  for (const component of components) {
    const existing = layoutServicesRuntimeOwners.get(component);
    if (existing && existing !== resolved) {
      throw new Error('Layout services combine foreign runtime owners');
    }
    layoutServicesRuntimeOwners.set(component, resolved);
  }
  return resolved;
}

export function attachBodyLayoutKernel(
  services: LayoutServices,
  kernel: BodyLayoutKernel,
): void {
  const owner = layoutServicesRuntimeOwner(services, true)!;
  if (bodyLayoutKernels.has(owner)) throw new Error('Body layout kernel is already attached');
  bodyLayoutKernels.set(owner, kernel);
}

export function bodyLayoutKernelOf(services: LayoutServices): BodyLayoutKernel | undefined {
  const owner = layoutServicesRuntimeOwner(services, false);
  return owner ? bodyLayoutKernels.get(owner) : undefined;
}

/** Bind every runtime service view to the one sealed acquisition source. */
export function attachLayoutSourceStore(
  services: LayoutServices,
  source: LayoutSourceStore,
): void {
  const owner = layoutServicesRuntimeOwner(services, true)!;
  if (layoutSourceStores.has(owner)) throw new Error('Layout source store is already attached');
  layoutSourceStores.set(owner, source);
}

export function layoutSourceStoreOf(services: LayoutServices): LayoutSourceStore | undefined {
  const owner = layoutServicesRuntimeOwner(services, false);
  return owner ? layoutSourceStores.get(owner) : undefined;
}

/** Attach the synchronous vertical-glyph metric authority to the same private
 * document owner as the public layout service components. Runtime service views
 * therefore inherit it without widening the stable LayoutServices contract. */
export function attachVerticalGlyphMeasurementService(
  services: LayoutServices,
  service: VerticalGlyphMeasurementService,
): void {
  const owner = layoutServicesRuntimeOwner(services, true)!;
  if (verticalGlyphMeasurementServices.has(owner)) {
    throw new Error('Vertical glyph measurement service is already attached');
  }
  verticalGlyphMeasurementServices.set(owner, service);
}

export function verticalGlyphMeasurementServiceOf(
  services: LayoutServices,
): VerticalGlyphMeasurementService {
  const owner = layoutServicesRuntimeOwner(services, false);
  const service = owner ? verticalGlyphMeasurementServices.get(owner) : undefined;
  if (!service) throw new Error('Vertical glyph measurement service is not attached');
  return service;
}

export function attachLayoutVariantStore(
  services: LayoutServices,
  store: LayoutVariantStore,
): void {
  if (layoutVariantStores.has(services)) throw new Error('Layout variant store is already attached');
  layoutVariantStores.set(services, store);
}

export function layoutVariantStoreOf(services: LayoutServices): LayoutVariantStore | undefined {
  return layoutVariantStores.get(services);
}

export function attachPrivateResourceLookup<T>(
  owner: object,
  entries: ReadonlyMap<string, T>,
  expectedKeys: Iterable<string> = entries.keys(),
): void {
  if (privateResourceLookups.has(owner)) {
    throw new Error('Private resource lookup is already attached');
  }
  const actual = new Set(entries.keys());
  const expected = new Set(expectedKeys);
  const missing = [...expected].filter((key) => !actual.has(key)).sort();
  const extra = [...actual].filter((key) => !expected.has(key)).sort();
  if (missing.length > 0 || extra.length > 0) {
    throw new Error(
      `Runtime resource membership mismatch: missing [${missing.join(', ')}]; extra [${extra.join(', ')}]`,
    );
  }
  privateResourceLookups.set(owner, createImmutableResourceLookup(entries));
}

export function privateResourceLookupOf<T>(owner: object): ImmutableResourceLookup<T> | undefined {
  return privateResourceLookups.get(owner) as ImmutableResourceLookup<T> | undefined;
}

const paintResourceRegistries = new WeakMap<object, PaintResourceRegistry>();

export interface FieldAcquisitionContext {
  readonly totalPages: number;
  /** Page numbering is finalized only after one pagination iteration. */
  readonly resolveDestinationPage?: (
    physicalPageIndex: number,
  ) => PageFieldAcquisitionContext | undefined;
}

export interface PageFieldAcquisitionContext {
  readonly pageIndex: number;
  readonly displayPageNumber: number;
  readonly pageNumberFormat: NumberFormat;
}

const fieldAcquisitionContexts = new WeakMap<object, FieldAcquisitionContext>();

export function createLayoutServicesRuntimeView(
  services: LayoutServices,
  overrides: Readonly<{ text?: LayoutServices['text']; math?: LayoutServices['math'] }> = {},
): LayoutServices {
  const view = Object.freeze({ ...services, ...overrides });
  const notePolicy = noteOverflowPolicyOf(services);
  if (notePolicy) attachNoteOverflowPolicy(view, notePolicy);
  const noteOwner = noteRecoveryOwnerOf(services);
  if (noteOwner) attachNoteRecoveryOwner(view, noteOwner);
  const kernel = bodyLayoutKernelOf(services);
  if (!kernel) throw new Error('Body layout kernel is not attached to the supplied services');
  if (bodyLayoutKernelOf(view) !== kernel) {
    throw new Error('Layout service view did not retain its body layout kernel owner');
  }
  const lookup = privateResourceLookups.get(services);
  if (lookup) privateResourceLookups.set(view, lookup);
  const registry = paintResourceRegistries.get(services);
  if (registry) paintResourceRegistries.set(view, registry);
  const paragraphCache = paragraphAcquisitionCaches.get(services);
  if (paragraphCache) paragraphAcquisitionCaches.set(view, paragraphCache);
  const storyCache = headerFooterLayoutCaches.get(services);
  if (storyCache) headerFooterLayoutCaches.set(view, storyCache);
  const footnoteBudget = footnoteAcquisitionWorkBudgets.get(services);
  if (footnoteBudget) footnoteAcquisitionWorkBudgets.set(view, footnoteBudget);
  return view;
}

/** Start one pagination acquisition scope. Paragraph memo and full-footnote
 * work ledger are fresh for a later layout/variant, but inherited by every
 * convergence service view inside this execution. */
export function createParagraphAcquisitionCacheServicesView(
  services: LayoutServices,
): LayoutServices {
  const view = createLayoutServicesRuntimeView(services);
  const workBudget = createPaginationWorkBudget(layoutSourceStoreOf(services));
  paragraphAcquisitionCaches.set(view, createParagraphAcquisitionRuntimeCache(workBudget));
  headerFooterLayoutCaches.set(view, createHeaderFooterLayoutCache(() => workBudget.noteMiss()));
  // Endnotes always continue independently of the footnote display option.
  // Every pagination scope owns the same aggregate ledger for both kinds.
  footnoteAcquisitionWorkBudgets.set(view, createFootnoteAcquisitionWorkBudget());
  return view;
}

export function paragraphAcquisitionCacheOf(
  services: LayoutServices,
): ParagraphAcquisitionRuntimeCache | undefined {
  return paragraphAcquisitionCaches.get(services);
}

/** Create one immutable service identity for a pagination-field iteration. */
export function createFieldAcquisitionServicesView(
  services: LayoutServices,
  context: FieldAcquisitionContext,
): LayoutServices {
  if (!Number.isInteger(context.totalPages) || context.totalPages < 1) {
    throw new RangeError('Field acquisition totalPages must be a positive integer');
  }
  const view = createLayoutServicesRuntimeView(services);
  fieldAcquisitionContexts.set(view, Object.freeze({ ...context }));
  return view;
}

export function fieldAcquisitionContextOf(owner: object): FieldAcquisitionContext {
  return fieldAcquisitionContexts.get(owner) ?? Object.freeze({ totalPages: 1 });
}

/** Associate cloneable resource descriptors with their document-scoped owner
 * without widening the stable LayoutServices or public document contracts. */
export function attachPaintResourceRegistry(
  owner: object,
  registry: PaintResourceRegistry,
): void {
  if (paintResourceRegistries.has(owner)) {
    throw new Error('Paint resource registry is already attached');
  }
  paintResourceRegistries.set(owner, registry);
}

export function paintResourceRegistryOf(owner: object): PaintResourceRegistry {
  const registry = paintResourceRegistries.get(owner);
  if (!registry) throw new Error('Paint resource registry is not attached');
  return registry;
}

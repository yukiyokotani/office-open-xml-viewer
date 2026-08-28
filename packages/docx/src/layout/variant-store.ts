import { deepFreezeDocumentLayout } from './invariants.js';
import { layoutOptionsKey, type LayoutOptions } from './options.js';
import type { DeepReadonly, DocumentLayout, LayoutPage, LayoutServices } from './types.js';

export type DocumentLayoutBuilder = (
  options: LayoutOptions,
) => DocumentLayout | DeepReadonly<DocumentLayout>;

export interface DocumentLayoutSelection {
  readonly key: string;
  readonly options: LayoutOptions;
  readonly layout: DeepReadonly<DocumentLayout>;
}

export function requireLayoutPage(
  layout: DeepReadonly<DocumentLayout>,
  pageIndex: number,
): DeepReadonly<LayoutPage> {
  if (!Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= layout.pages.length) {
    throw new RangeError(`Page index ${pageIndex} out of range (count: ${layout.pages.length})`);
  }
  return layout.pages[pageIndex] as DeepReadonly<LayoutPage>;
}

/**
 * Document-scoped layout cache. The key deliberately excludes paint-only facts
 * such as scale, DPR, and color: only acquisition inputs may select geometry.
 */
export class LayoutVariantStore {
  readonly #services: LayoutServices;
  readonly #build: DocumentLayoutBuilder;
  readonly #variants = new Map<string, DeepReadonly<DocumentLayout>>();
  readonly #defaultOptions: LayoutOptions;
  readonly #defaultKey: string;
  #activeNonDefaultKey: string | null = null;

  constructor(
    services: LayoutServices,
    defaultOptions: LayoutOptions,
    build: DocumentLayoutBuilder,
  ) {
    this.#services = services;
    this.#defaultOptions = Object.freeze({ ...defaultOptions });
    this.#defaultKey = layoutOptionsKey(this.#defaultOptions, this.#services);
    this.#build = build;
  }

  get defaultLayout(): DeepReadonly<DocumentLayout> {
    return this.layoutFor(this.#defaultOptions);
  }

  layoutFor(options: LayoutOptions): DeepReadonly<DocumentLayout> {
    return this.select(options).layout;
  }

  select(options: LayoutOptions): DocumentLayoutSelection {
    const normalized = Object.isFrozen(options)
      ? options
      : Object.freeze({ ...options });
    const key = layoutOptionsKey(normalized, this.#services);
    let layout = this.#variants.get(key);
    if (!layout) {
      layout = deepFreezeDocumentLayout(this.#build(normalized) as DocumentLayout);
      if (key !== this.#defaultKey) {
        if (this.#activeNonDefaultKey !== null && this.#activeNonDefaultKey !== key) {
          this.#variants.delete(this.#activeNonDefaultKey);
        }
        this.#activeNonDefaultKey = key;
      }
      this.#variants.set(key, layout);
    }
    return Object.freeze({ key, options: normalized, layout });
  }

  selectPage(
    options: LayoutOptions,
    pageIndex: number,
  ): Readonly<{
    key: string;
    options: LayoutOptions;
    layout: DeepReadonly<DocumentLayout>;
    page: DeepReadonly<LayoutPage>;
  }> {
    const selection = this.select(options);
    return Object.freeze({
      ...selection,
      page: requireLayoutPage(selection.layout, pageIndex),
    });
  }

  /**
   * Deposit a layout built outside this store (by the asynchronous, sliced
   * builder) under its options key, so every later synchronous `select` — the
   * render path included — hits it instead of rebuilding.
   *
   * Retention follows the same one-non-default-variant policy as `select`, so
   * priming cannot grow the cache beyond what building normally would.
   */
  prime(
    options: LayoutOptions,
    layout: DocumentLayout,
  ): DeepReadonly<DocumentLayout> {
    const normalized = Object.isFrozen(options) ? options : Object.freeze({ ...options });
    const key = layoutOptionsKey(normalized, this.#services);
    const existing = this.#variants.get(key);
    if (existing) return existing;
    return this.#store(key, layout);
  }

  /**
   * Atomically replace one exact cached layout. Progressive pagination uses
   * the retained return value from its previous publication as an ownership
   * token: if another consumer evicted or rebuilt the same variant meanwhile,
   * the stale session can no longer overwrite that newer authority.
   *
   * Passing `null` claims an absent key for the first publication. Returning
   * `null` means ownership was not acquired or has been lost.
   */
  replaceIfCurrent(
    options: LayoutOptions,
    expected: DeepReadonly<DocumentLayout> | null,
    layout: DocumentLayout,
  ): DeepReadonly<DocumentLayout> | null {
    const normalized = Object.isFrozen(options) ? options : Object.freeze({ ...options });
    const key = layoutOptionsKey(normalized, this.#services);
    if ((this.#variants.get(key) ?? null) !== expected) return null;
    return this.#store(key, layout);
  }

  #store(key: string, layout: DocumentLayout): DeepReadonly<DocumentLayout> {
    const frozen = deepFreezeDocumentLayout(layout);
    if (key !== this.#defaultKey) {
      if (this.#activeNonDefaultKey !== null && this.#activeNonDefaultKey !== key) {
        this.#variants.delete(this.#activeNonDefaultKey);
      }
      this.#activeNonDefaultKey = key;
    }
    this.#variants.set(key, frozen);
    return frozen;
  }

  /** Whether a layout for these options is already available synchronously. */
  hasLayoutFor(options: LayoutOptions): boolean {
    return this.#variants.has(layoutOptionsKey(options, this.#services));
  }

  isDefault(options: LayoutOptions): boolean {
    return layoutOptionsKey(options, this.#services) === this.#defaultKey;
  }
}

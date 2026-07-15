import { deepFreezeDocumentLayout } from './invariants.js';
import { layoutOptionsKey, type LayoutOptions } from './options.js';
import type { DeepReadonly, DocumentLayout, LayoutServices } from './types.js';

export type DocumentLayoutBuilder = (
  options: LayoutOptions,
) => DocumentLayout | DeepReadonly<DocumentLayout>;

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
    const normalized = Object.freeze({ ...options });
    const key = layoutOptionsKey(normalized, this.#services);
    const cached = this.#variants.get(key);
    if (cached) return cached;
    const layout = deepFreezeDocumentLayout(this.#build(normalized) as DocumentLayout);
    this.#variants.set(key, layout);
    return layout;
  }

  isDefault(options: LayoutOptions): boolean {
    return layoutOptionsKey(options, this.#services) === this.#defaultKey;
  }
}

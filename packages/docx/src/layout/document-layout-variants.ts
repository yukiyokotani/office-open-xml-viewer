import {
  resolvedFromWire,
  type ResolvedErrorPlaceholder,
  type WireErrorPlaceholder,
} from '@silurus/ooxml-core/internal/part-failures';
import { layoutParseErrorPage } from './error-page.js';
import { deepFreezeDocumentLayout } from './invariants.js';
import type { LayoutSourceStore } from './layout-source-store.js';
import {
  layoutOptionsForRender,
  type LayoutOptions,
  type LayoutRenderSelectionInput,
} from './options.js';
import {
  attachLayoutVariantStore,
  layoutVariantStoreOf,
} from './runtime-state.js';
import type { FatalParseFact } from './layout-source-store.js';
import type { DeepReadonly, DocumentLayout, LayoutPage, LayoutServices } from './types.js';
import { LayoutVariantStore, type DocumentLayoutBuilder } from './variant-store.js';

/**
 * The caller-text error page of a fatally unparseable document (#1974). It is
 * laid out at acquisition, like the built-in page, and at most one is retained
 * per document: a different title/detail pair replaces it. The built-in page
 * stays in the variant store untouched.
 */
class CustomErrorPageSlot {
  #key: string | null = null;
  #layout: DeepReadonly<DocumentLayout> | null = null;

  constructor(
    private readonly fatalParse: FatalParseFact,
    private readonly services: LayoutServices,
  ) {}

  layoutFor(placeholder: ResolvedErrorPlaceholder): DeepReadonly<DocumentLayout> {
    const key = JSON.stringify([placeholder.title, placeholder.detail ?? 0]);
    if (this.#key !== key || this.#layout === null) {
      this.#layout = deepFreezeDocumentLayout(layoutParseErrorPage(
        this.fatalParse.message,
        this.fatalParse.pageSize,
        this.services.text,
        placeholder,
      ));
      this.#key = key;
    }
    return this.#layout;
  }
}

const customErrorPages = new WeakMap<LayoutServices, CustomErrorPageSlot>();

export type DocumentLayoutVariantFactoryInput = Readonly<{
  readonly source: LayoutSourceStore;
  readonly services: LayoutServices;
  readonly defaultCurrentDateMs: number;
  readonly buildLayout: DocumentLayoutBuilder;
}>;

export interface AttachedDocumentLayoutVariants {
  readonly store: LayoutVariantStore;
  readonly defaultOptions: LayoutOptions;
}

export function attachDocumentLayoutVariants(
  input: DocumentLayoutVariantFactoryInput,
): AttachedDocumentLayoutVariants {
  const { services, defaultCurrentDateMs, buildLayout } = input;
  const defaultOptions = layoutOptionsForRender({ defaultCurrentDateMs });
  const fatalParse = input.source.fatalParse;
  const parseError = fatalParse === null
    ? null
    : layoutParseErrorPage(
        fatalParse.message,
        fatalParse.pageSize,
        services.text,
      );
  const store = new LayoutVariantStore(
    services,
    defaultOptions,
    parseError === null ? buildLayout : () => parseError,
  );
  attachLayoutVariantStore(services, store);
  if (fatalParse !== null) customErrorPages.set(services, new CustomErrorPageSlot(fatalParse, services));
  return Object.freeze({ store, defaultOptions });
}

export function selectDocumentLayoutPage(
  services: LayoutServices,
  input: LayoutRenderSelectionInput & {
    /** Caller text for a fatally unparseable document's error page, already
     * resolved on the main thread. Absent means the built-in text. */
    readonly errorPlaceholderText?: WireErrorPlaceholder;
  },
  pageIndex: number,
): Readonly<{
  key: string;
  options: LayoutOptions;
  layout: DeepReadonly<DocumentLayout>;
  page: DeepReadonly<LayoutPage>;
}> {
  const store = layoutVariantStoreOf(services);
  if (!store) throw new Error('Document layout variant store is not attached to the supplied services');
  const selection = store.selectPage(layoutOptionsForRender(input), pageIndex);
  const custom = input.errorPlaceholderText ? customErrorPages.get(services) : undefined;
  if (!custom) return selection;
  const placeholder = resolvedFromWire(input.errorPlaceholderText);
  if (placeholder === null) return selection;
  const layout = custom.layoutFor(placeholder);
  const page = layout.pages[pageIndex];
  if (!page) return selection;
  return Object.freeze({ key: `${selection.key}#error-placeholder`, options: selection.options, layout, page });
}

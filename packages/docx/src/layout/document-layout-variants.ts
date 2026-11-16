import {
  ERROR_PLACEHOLDER_TEXT_HOST,
  onDemandModule,
  type ResolvedErrorPlaceholder,
  type WireErrorPlaceholder,
} from '@silurus/ooxml-core/internal/part-failures';
import { deepFreezeDocumentLayout } from './invariants.js';
import { createPageLayers } from './page-graph.js';
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
 * The synthetic page of a fatally unparseable document is code no healthy
 * document needs, so it is loaded with the first such document (#2248). A
 * self-contained worker bundle inlines it.
 */
const errorPage = /* @__PURE__ */ onDemandModule(async () => {
  const { parseErrorPageLayout, resolvedFromWire } = await import('./error-page.js');
  return {
    layoutParseErrorPage: parseErrorPageLayout({ ...ERROR_PLACEHOLDER_TEXT_HOST, createPageLayers }),
    resolvedFromWire,
  };
});

/**
 * Layout is synchronous: the error page's code has to be here before a
 * fatally unparseable source gets its layout variants. Whatever hands out such
 * a source asynchronously awaits this first: the pull materializers
 * (`document-pull-client.ts`), the render worker, the model-source loaders and
 * a direct `renderLayoutSourceToCanvas()`. `undefined` when there is nothing
 * to wait for, which is always the case for a source that parsed.
 */
export function errorPageLayoutReady(source: Pick<LayoutSourceStore, 'fatalParse'>): Promise<unknown> | undefined {
  return source.fatalParse === null || errorPage.loaded ? undefined : errorPage.load();
}

function loadedErrorPage() {
  const loaded = errorPage.loaded;
  if (!loaded) throw new Error('DOCX error page layout was not loaded before its document was laid out');
  return loaded;
}

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
    // Omitted (built-in) and hidden (`null`) detail are different pages.
    const key = JSON.stringify([
      placeholder.title,
      placeholder.detail === undefined ? 'builtin' : placeholder.detail === null ? 'hidden' : ['text', placeholder.detail],
    ]);
    if (this.#key !== key || this.#layout === null) {
      this.#layout = deepFreezeDocumentLayout(loadedErrorPage().layoutParseErrorPage(
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
    : loadedErrorPage().layoutParseErrorPage(
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
  input: LayoutRenderSelectionInput,
  pageIndex: number,
): Readonly<{
  key: string;
  options: LayoutOptions;
  layout: DeepReadonly<DocumentLayout>;
  page: DeepReadonly<LayoutPage>;
}> {
  const store = layoutVariantStoreOf(services);
  if (!store) throw new Error('Document layout variant store is not attached to the supplied services');
  return store.selectPage(layoutOptionsForRender(input), pageIndex);
}

/**
 * Swap a fatally unparseable document's selected error page for the one laid
 * out with caller text (#1974), acquired here before paint. Any other
 * selection, and the built-in text, are returned unchanged.
 */
export function withErrorPlaceholderPage<
  Selection extends Readonly<{
    key: string;
    options: LayoutOptions;
    layout: DeepReadonly<DocumentLayout>;
    page: DeepReadonly<LayoutPage>;
  }>,
>(
  services: LayoutServices,
  selection: Selection,
  errorPlaceholderText: WireErrorPlaceholder | undefined,
  pageIndex: number,
): Selection {
  const custom = errorPlaceholderText ? customErrorPages.get(services) : undefined;
  if (!custom) return selection;
  const placeholder = loadedErrorPage().resolvedFromWire(errorPlaceholderText);
  if (placeholder === null) return selection;
  const layout = custom.layoutFor(placeholder);
  const page = layout.pages[pageIndex];
  if (!page) return selection;
  return Object.freeze({ ...selection, key: `${selection.key}#error-placeholder`, layout, page });
}

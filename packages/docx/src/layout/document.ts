import type { BodyLayoutInput } from './body-layout-input.js';
import type { LayoutSourceStore } from './layout-source-store.js';
import { paginateBody, paginateBodySteps } from './body-paginator.js';
import {
  drainPaginationAsync,
  type PaginationSchedulerOptions,
} from './pagination-scheduler.js';
import {
  attachDocumentLayoutVariants,
} from './document-layout-variants.js';
import { layoutVariantStoreOf } from './runtime-state.js';
import { normalizeLayoutOptions, type LayoutOptions } from './options.js';
import type { DocumentLayout, LayoutServices } from './types.js';

export function layoutDocumentInput(
  input: BodyLayoutInput,
  services: LayoutServices,
  options: LayoutOptions = normalizeLayoutOptions(undefined, Date.now()),
): DocumentLayout {
  return paginateBody(input, services, options);
}

/**
 * {@link layoutDocumentInput} spread across event-loop turns.
 *
 * Same generator, same result — only the driving differs — so an awaited layout
 * is byte-identical to the blocking one.
 */
export function layoutDocumentInputAsync(
  input: BodyLayoutInput,
  services: LayoutServices,
  options: LayoutOptions = normalizeLayoutOptions(undefined, Date.now()),
  scheduler?: PaginationSchedulerOptions,
): Promise<DocumentLayout> {
  return drainPaginationAsync(paginateBodySteps(input, services, options), scheduler);
}

export function ensureDocumentLayoutVariants(
  services: LayoutServices,
  defaultCurrentDateMs: number,
  resolveSource: () => LayoutSourceStore,
): void {
  if (layoutVariantStoreOf(services)) return;
  const source = resolveSource();
  attachDocumentLayoutVariants({
    source,
    services,
    defaultCurrentDateMs,
    buildLayout: (options) => layoutDocumentInput(source.bodyLayoutInput, services, options),
  });
}

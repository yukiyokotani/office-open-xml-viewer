import type { FontDemandCollector } from './presentation-preflight.js';

/** Analysis/data is optional: ordinary parser/viewer graphs do not execute or
 * fetch it. The builder's ACK transaction stays synchronous after injection. */
export async function loadFontDemandCollector(enabled: boolean): Promise<FontDemandCollector | undefined> {
  return enabled ? (await import('./font-preload-demand.js')).collectSlideFontDemand : undefined;
}

/** The collector in a parse worker, imported from the sidecar module whose
 * absolute URL the window resolved (font-demand-collector-url.ts). An inline
 * worker cannot resolve relative chunks, so a relative dynamic import would
 * inline the collector into the worker source of the static graph.
 *
 * A bundler that does not emit `new URL(…, import.meta.url)` assets (esbuild,
 * Angular CLI; README) leaves the sidecar unserved. The load then demands
 * every subset of the same families, as before demand was collected: the
 * same faces paint, only more of their subsets load. */
export async function importFontDemandCollector(url: string): Promise<FontDemandCollector> {
  if (!url) return demandEverySubset;
  try {
    return ((await import(/* @vite-ignore */ url)) as typeof import('./internal/font-demand-sidecar.js')).collectSlideFontDemand;
  } catch {
    return demandEverySubset;
  }
}

const demandEverySubset: FontDemandCollector = () => 'all';

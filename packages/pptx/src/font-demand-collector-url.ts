import { resolveBundledAssetUrl } from '@silurus/ooxml-core/internal/bundled-asset-url';

/** Absolute URL of the font-demand collector sidecar (`pptx-font-demand.mjs`,
 * scripts/build-model-source-sidecars.mjs) that the parse worker imports.
 * Window-only: the worker receives the resolved URL in its parse request. */
export function fontDemandCollectorUrl(): string {
  return resolveBundledAssetUrl(new URL(
    import.meta.env.DEV ? './font-preload-demand.ts' : './pptx-font-demand.mjs',
    import.meta.url,
  ).href, import.meta.url);
}

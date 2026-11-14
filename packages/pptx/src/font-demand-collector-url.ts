import { resolveBundledAssetUrl } from '@silurus/ooxml-core/internal/bundled-asset-url';

/** Absolute URL of the font-demand collector sidecar (`pptx-font-demand.mjs`,
 * scripts/build-model-source-sidecars.mjs) that the parse worker imports, or
 * `''` when this bundle has no module URL to resolve it against (an esbuild
 * IIFE leaves `import.meta.url` undefined); the worker then demands every
 * subset (font-demand-collector.ts). Window-only: the worker receives the
 * resolved URL in its parse request. */
export function fontDemandCollectorUrl(): string {
  try {
    return resolveBundledAssetUrl(new URL(
      import.meta.env.DEV ? './font-preload-demand.ts' : './pptx-font-demand.mjs',
      import.meta.url,
    ).href, import.meta.url);
  } catch {
    return '';
  }
}

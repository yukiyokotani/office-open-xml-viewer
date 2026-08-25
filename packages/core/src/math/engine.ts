/// <reference types="vite/client" />

// MathML → SVG via a pre-bundled MathJax v4 + STIX Two Math converter.
//
// This module is the *heavy* half of the math feature: it references the
// pre-built engine asset `assets/mathjax-stix2.js` (~3 MB; MathJax core + the
// statically-baked STIX2 font), so anything that statically imports it drags
// that asset into the bundle.
//
// To keep the asset OUT of the docx/pptx initial bundles, the renderers do NOT
// import this module. Instead it is published as a *separate* entry point
// (`@silurus/ooxml/math`) that consumers explicitly import and pass to a viewer
// (`new DocxViewer(canvas, { math })`). When they don't, the whole asset
// is not loaded or evaluated unless supplied. See `src/math.ts` (root) and the
// `MathRenderer` interface.
//
// The asset itself is self-contained: DOM-free internally, zero network, zero
// cross-origin requests. It exposes `globalThis.__ooxmlStix2`.

// `?url` (not a bare `new URL(..., import.meta.url)`) so the same `wasmAssetUrl`
// build plugin that keeps the WASM parsers out of the base64 data-URL trap emits
// this ~3 MB engine as a real asset too. In Vite **library mode** a bare
// `new URL` is force-inlined as a `data:text/javascript;base64,…` string, which
// turned the opt-in `math.mjs` chunk into a 4.1 MB base64 blob; the `?url` form
// is intercepted by the plugin, `emitFile`d as a real asset next to the chunk,
// and handed back as a plain URL the `<script>` loader below can fetch.
//
// The engine URL is not otherwise configurable: a consumer that needs to serve
// it from elsewhere injects their own `MathRenderer` via the viewer `math`
// option (the whole engine is already a swappable dependency), so no dedicated
// `mathUrl`/asset-override option is warranted.
import mathjaxAssetUrl from '../../assets/mathjax-stix2.js?url';
import type { MathSvg } from './mathjax.js';
import {
  loadMathJaxFromResolvedAsset,
  mathMLToSvgFromResolvedAsset,
} from './engine-runtime.js';

export function resolveMathJaxAssetUrl(): string {
  // `?url` yields the asset href directly — an absolute URL at build time, the
  // dev-served path in dev. Resolve against the module URL so a bare relative
  // dev value still becomes an absolute href fetchable from any realm (matches
  // the `new URL(wasmAssetUrl, …)` pattern in the format handles).
  return new URL(mathjaxAssetUrl, import.meta.url).href;
}

function normalizedAssetUrl(assetUrl?: string): string {
  return assetUrl ? new URL(assetUrl, import.meta.url).href : resolveMathJaxAssetUrl();
}

/** Preload the math engine. Call once before rendering equations. */
export async function loadMathJax(): Promise<void> {
  await loadMathJaxFromResolvedAsset(resolveMathJaxAssetUrl());
}

/** Internal worker entry: use the asset URL resolved by the consumer bundler
 * in the main realm instead of resolving relative to an opaque worker asset. */
export async function loadMathJaxFromAsset(assetUrl: string): Promise<void> {
  await loadMathJaxFromResolvedAsset(normalizedAssetUrl(assetUrl));
}

/** Convert a MathML string to a standalone SVG + its baseline-relative extents. */
export async function mathMLToSvg(mathml: string): Promise<MathSvg> {
  return mathMLToSvgFromResolvedAsset(mathml, resolveMathJaxAssetUrl());
}

/** Internal worker entry paired with {@link loadMathJaxFromAsset}. */
export async function mathMLToSvgFromAsset(
  mathml: string,
  assetUrl: string,
): Promise<MathSvg> {
  return mathMLToSvgFromResolvedAsset(mathml, normalizedAssetUrl(assetUrl));
}

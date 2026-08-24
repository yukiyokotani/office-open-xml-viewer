import { type MathSvg, svgExtents } from './mathjax.js';

/* eslint-disable @typescript-eslint/no-explicit-any */

interface Stix2Engine {
  /** MathML string → standalone `<svg>…</svg>` (currentColor fill, viewBox in 1000-units/em). */
  mathml2svg(mathml: string): string;
}

let enginePromise: Promise<Stix2Engine> | null = null;

function loadScript(src: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = src;
    script.async = true;
    script.onload = () => resolve();
    script.onerror = () => reject(new Error(`Failed to load math engine from ${src}`));
    document.head.appendChild(script);
  });
}

/** Module workers have no document and cannot use the classic-worker
 * importScripts API. The prebuilt engine is a strict IIFE that is also valid as
 * a side-effect-only ES module, so dynamic import evaluates it in the worker's
 * own global realm and installs `globalThis.__ooxmlStix2`. */
async function loadWorkerModule(src: string): Promise<void> {
  await import(/* @vite-ignore */ src);
}

function ensureEngine(assetUrl: string): Promise<Stix2Engine> {
  if (enginePromise) return enginePromise;
  enginePromise = (async () => {
    const existing = (globalThis as any).__ooxmlStix2 as Stix2Engine | undefined;
    if (existing) return existing;
    // The main realm resolves this URL before its descriptor crosses the worker
    // boundary. Keeping that resolution out of this runtime prevents opaque
    // worker assets from retaining import.meta solely for optional MathJax.
    const resolvedAssetUrl = new URL(assetUrl).href;
    if (typeof document === 'undefined') await loadWorkerModule(resolvedAssetUrl);
    else await loadScript(resolvedAssetUrl);
    const engine = (globalThis as any).__ooxmlStix2 as Stix2Engine | undefined;
    if (!engine) throw new Error('Math engine failed to initialize');
    return engine;
  })();
  return enginePromise;
}

/** Preload MathJax from an absolute URL resolved outside an opaque worker asset. */
export async function loadMathJaxFromResolvedAsset(assetUrl: string): Promise<void> {
  await ensureEngine(assetUrl);
}

/** Convert MathML using an absolute URL resolved outside an opaque worker asset. */
export async function mathMLToSvgFromResolvedAsset(
  mathml: string,
  assetUrl: string,
): Promise<MathSvg> {
  const engine = await ensureEngine(assetUrl);
  const svg = engine.mathml2svg(mathml);
  return { svg, ...svgExtents(svg) };
}

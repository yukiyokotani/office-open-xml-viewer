/**
 * Format-generic loader for a worker-realm sidecar module: one self-contained
 * ES module the owning realm emitted beside its chunks and resolved to an
 * absolute URL (scripts/build-model-source-sidecars.mjs), imported where an
 * inline (Blob) worker cannot resolve a relative chunk. Used by the PPTX
 * parse workers for the font-demand collector.
 *
 * The URL is resolved by library code from `import.meta.url`; document bytes
 * never choose it. Only an absolute http(s), file or blob URL is accepted, so
 * a bare package specifier (another package, such as a legacy reader) cannot
 * be loaded through it. This is the module's one computed import
 * (scripts/check-core-legacy-boundary.mjs ALLOWED_COMPUTED_IMPORTS).
 */
const SIDECAR_URL_PROTOCOLS: ReadonlySet<string> = new Set(['http:', 'https:', 'file:', 'blob:']);

/** Import the sidecar at `url` (absolute http(s), file or blob URL). Rejects
 * with a `TypeError` for any other specifier, before importing. */
export async function importWorkerSidecar(url: string): Promise<unknown> {
  if (typeof url !== 'string' || url.length === 0 || url.trim() !== url) {
    throw new TypeError('worker sidecar URL must be a nonempty absolute URL');
  }
  let parsed: URL;
  try {
    parsed = new URL(url);
  } catch {
    throw new TypeError('worker sidecar URL must be a nonempty absolute URL');
  }
  if (!SIDECAR_URL_PROTOCOLS.has(parsed.protocol)) throw new TypeError('worker sidecar URL protocol is unsupported');
  return import(/* @vite-ignore */ url);
}

/**
 * The calling realm's base for a relative application URL: the document base
 * URL in a window (it honours `<base href>`, as `fetch()` and `<script src>`
 * do), the worker location in a worker, and undefined in a realm without
 * either (Node). Resolve a relative URL with this before posting it to a
 * worker, whose own location would otherwise be the base.
 */
export function realmBaseUrl(): string | undefined {
  const scope = globalThis as { document?: { baseURI?: unknown }; location?: { href?: unknown } };
  const base = scope.document?.baseURI ?? scope.location?.href;
  return typeof base === 'string' ? base : undefined;
}

/** Resolve a bundler-emitted asset before passing it to an inline worker. */
export function resolveBundledAssetUrl(assetUrl: string, ownerUrl: string): string {
  // Bundlers may emit root-relative assets while retaining a build-time file:
  // import.meta.url (e.g. Turbopack). Use the calling page for these paths only
  // when the owner is file: or cannot base a root path (WHATWG URL semantics).
  // A valid HTTP/CDN owner keeps its own origin, including for root paths.
  const pageBase = realmBaseUrl();
  if (assetUrl.startsWith('/') && pageBase) {
    let ownerRoot: URL | undefined;
    try { ownerRoot = new URL('/', ownerUrl); } catch { /* Opaque build owner. */ }
    if (!ownerRoot || ownerRoot.protocol === 'file:') {
      try {
        return new URL(assetUrl, pageBase).href;
      } catch {
        // A blob:/data: worker or about:blank page cannot base a root path.
        // Retain module-owner resolution (or its ordinary invalid-URL error).
      }
    }
  }
  // Module-relative and explicit absolute assets, including Node file: URLs,
  // retain their original semantics. Do not reinterpret an absolute file URL.
  return new URL(assetUrl, ownerUrl).href;
}

/**
 * Fetch a bundled parser WASM from the webview document and return a Blob URL
 * for the viewer's `wasmUrl` option.
 *
 * The parser runs in a Blob-URL worker, and the worker is what fetches
 * `wasmUrl`. VS Code serves extension resources through the webview's service
 * worker, which asks the workbench for the bytes on behalf of the webview a
 * request came from. It identifies that webview from the `id` parameter of the
 * requesting client's URL. A worker client has no such parameter, so nobody is
 * asked and the service worker answers 408 after its 30 s timeout (#2212), in
 * local and remote windows alike. A request from the document is served
 * normally, and a Blob URL is read by the worker without the service worker.
 *
 * The Blob URL is never revoked: the parser re-instantiates from it after a
 * trap.
 */
export async function fetchParserWasmBlobUrl(
  assetUrl: string,
  baseUrl: string,
  fetchImpl: typeof fetch = fetch,
): Promise<string> {
  const response = await fetchImpl(new URL(assetUrl, baseUrl));
  if (!response.ok) {
    throw new Error(`the parser WebAssembly could not be loaded (HTTP ${response.status})`);
  }
  const bytes = await response.arrayBuffer();
  return URL.createObjectURL(new Blob([bytes], { type: 'application/wasm' }));
}

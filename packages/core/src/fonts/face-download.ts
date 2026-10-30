/**
 * Library downloads of `url` faces. Library resource policy, not OOXML
 * behaviour.
 *
 * Face files are normally left to `FontFace.load()` in the painting realm.
 * The library downloads them itself, and hands the bytes to `FontFace` as a
 * `data` face, in two cases only:
 *
 * - a source configured with `credentials: 'omit'` (`fontCssApi`, `fontCss`,
 *   `fontFiles`): CSS font fetching always sends same-origin credentials, so
 *   omitting them needs a request of the library's own, in the painting realm;
 * - a {@link fontResolver} answer bridged to a render worker: the page
 *   downloads every http(s) / blob: URL and transfers the bytes, so cookies,
 *   HTTP authentication and Referer checks see the page exactly as when the
 *   page paints (`resolver-bridge-host.ts`).
 *
 * Requests are CORS requests, as CSS font fetching is. A body is bounded by
 * `HARD_MAX_FONT_SOURCE_FACE_BYTES`, the ceiling every `data` face obeys, and
 * a breach rejects with `OoxmlError('font-source-limit')`. A network error or
 * a non-OK status fails that face only (its family keeps its other faces).
 * `data:` URLs are kept as they are: they are inline and fetch nothing.
 */
import { fontSourceLimitError } from '../errors/font-source-limit.js';
import { OoxmlError } from '../errors/ooxml-error.js';
import { HARD_MAX_FONT_SOURCE_FACE_BYTES } from '../worker/resource-policy.generated.js';
import type { FontFaceDeclaration, FontSourceCredentials } from './source.js';

/** Whether the library downloads `declaration` (an http(s) or blob: URL). */
export function isDownloadableFace(declaration: FontFaceDeclaration): boolean {
  return declaration.url !== undefined && /^(?:https?|blob):/i.test(declaration.url);
}

interface FaceResponse {
  readonly ok: boolean;
  readonly status?: number;
  readonly headers?: { get(name: string): string | null };
  readonly body?: { getReader?(): ReadableStreamDefaultReader<Uint8Array>; cancel?(): Promise<void> } | null;
  arrayBuffer(): Promise<ArrayBuffer>;
}

async function readBoundedBytes(response: FaceResponse): Promise<ArrayBuffer> {
  const limit = HARD_MAX_FONT_SOURCE_FACE_BYTES;
  const declared = Number(response.headers?.get('content-length') ?? NaN);
  if (Number.isSafeInteger(declared) && declared > limit) {
    response.body?.cancel?.().catch(() => undefined);
    throw fontSourceLimitError('face', limit, declared);
  }
  const reader = typeof response.body?.getReader === 'function' ? response.body.getReader() : null;
  if (!reader) {
    const buffer = await response.arrayBuffer();
    if (buffer.byteLength > limit) throw fontSourceLimitError('face', limit, buffer.byteLength);
    return buffer;
  }
  const chunks: Uint8Array[] = [];
  let size = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    if (!value) continue;
    size += value.byteLength;
    if (size > limit) {
      reader.cancel().catch(() => undefined);
      throw fontSourceLimitError('face', limit, size);
    }
    chunks.push(value);
  }
  const bytes = new Uint8Array(size);
  let offset = 0;
  for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.byteLength; }
  return bytes.buffer;
}

/** Download one face; a failure reason, or the owned bytes. */
async function downloadFace(
  url: string,
  credentials: FontSourceCredentials,
  signal: AbortSignal,
): Promise<{ readonly data: ArrayBuffer } | { readonly reason: string }> {
  if (typeof fetch !== 'function') return { reason: 'fetch is unavailable' };
  const request = new AbortController();
  const onAbort = () => request.abort();
  signal.addEventListener('abort', onAbort, { once: true });
  if (signal.aborted) request.abort();
  let complete = false;
  try {
    const response = await fetch(url, { mode: 'cors', credentials, signal: request.signal }) as FaceResponse;
    if (!response.ok) return { reason: `HTTP ${response.status ?? 'error'}` };
    const data = await readBoundedBytes(response);
    complete = true;
    if (data.byteLength === 0) return { reason: 'empty font file' };
    return { data };
  } catch (error) {
    if (error instanceof OoxmlError) throw error;
    return { reason: signal.aborted ? 'timed out' : error instanceof Error ? error.message : String(error) };
  } finally {
    signal.removeEventListener('abort', onAbort);
    if (!complete) request.abort();
  }
}

/** Downloaded replacements for a set of declarations. */
export interface FaceDownloads {
  /** `declaration` with its downloaded bytes in place of the URL; the
   * declaration itself when it is not downloaded; undefined when its
   * download failed. */
  readonly replace: (declaration: FontFaceDeclaration) => FontFaceDeclaration | undefined;
  /** Failure reason per failed URL. */
  readonly failures: ReadonlyMap<string, string>;
  /** Every downloaded buffer (fresh, owned by the caller). */
  readonly buffers: readonly ArrayBuffer[];
}

/**
 * Download every downloadable face of `declarations` once per URL, all
 * concurrently, with `credentials`. Rejects with
 * `OoxmlError('font-source-limit')` when a body exceeds the face ceiling;
 * any other failure (network, status, abort) fails only the faces of that URL.
 */
export async function downloadFontFaces(
  declarations: Iterable<FontFaceDeclaration>,
  credentials: FontSourceCredentials,
  signal: AbortSignal,
): Promise<FaceDownloads> {
  const urls = new Set<string>();
  for (const declaration of declarations) if (isDownloadableFace(declaration)) urls.add(declaration.url as string);
  const results = new Map(await Promise.all([...urls].map(async (url) =>
    [url, await downloadFace(url, credentials, signal)] as const)));
  const failures = new Map<string, string>();
  const buffers: ArrayBuffer[] = [];
  for (const [url, result] of results) {
    if ('reason' in result) failures.set(url, result.reason);
    else buffers.push(result.data);
  }
  return {
    failures,
    buffers,
    replace: (declaration) => {
      if (!isDownloadableFace(declaration)) return declaration;
      const result = results.get(declaration.url as string);
      if (!result || 'reason' in result) return undefined;
      const { family, weight, style, stretch, unicodeRange } = declaration;
      return Object.freeze({
        family,
        data: result.data,
        ...(weight === undefined ? {} : { weight }),
        ...(style === undefined ? {} : { style }),
        ...(stretch === undefined ? {} : { stretch }),
        ...(unicodeRange === undefined ? {} : { unicodeRange }),
      });
    },
  };
}

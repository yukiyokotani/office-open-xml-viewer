import {
  DELIMITED_TEXT_MAX_SOURCE_BYTES,
  assertDelimitedTextSourceBytes,
} from './delimited-text.js';

/** Read a fetched delimited-text body without ever retaining an unbounded
 * response. `Content-Length` rejects early when available; the stream count is
 * authoritative because responses may be compressed or misdeclare length. */
export async function readDelimitedTextResponse(response: Response): Promise<ArrayBuffer> {
  const contentEncoding = response.headers.get('content-encoding')?.trim().toLowerCase();
  const rawContentLength = response.headers.get('content-length');
  // Fetch exposes decoded body chunks while Content-Length describes the
  // transported representation. It is an early bound only without compression.
  if (
    (!contentEncoding || contentEncoding === 'identity')
    && rawContentLength !== null
    && /^\d+$/.test(rawContentLength.trim())
  ) {
    const declaredBytes = Number(rawContentLength);
    if (declaredBytes > DELIMITED_TEXT_MAX_SOURCE_BYTES) {
      await response.body?.cancel().catch(() => undefined);
      assertDelimitedTextSourceBytes(declaredBytes);
    }
  }

  const body = response.body;
  if (!body) return new ArrayBuffer(0);

  const reader = body.getReader();
  const chunks: Uint8Array[] = [];
  let byteLength = 0;
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      if (!value || value.byteLength === 0) continue;
      if (value.byteLength > DELIMITED_TEXT_MAX_SOURCE_BYTES - byteLength) {
        await reader.cancel().catch(() => undefined);
        assertDelimitedTextSourceBytes(DELIMITED_TEXT_MAX_SOURCE_BYTES + 1);
      }
      chunks.push(value);
      byteLength += value.byteLength;
    }
  } finally {
    reader.releaseLock();
  }

  const bytes = new Uint8Array(byteLength);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return bytes.buffer;
}

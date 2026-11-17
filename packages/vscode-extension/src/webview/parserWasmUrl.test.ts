import { afterEach, describe, expect, it, vi } from 'vitest';
import { fetchParserWasmBlobUrl } from './parserWasmUrl';

describe('fetchParserWasmBlobUrl', () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it('fetches the asset against the document base and returns a WASM Blob URL', async () => {
    const bytes = new Uint8Array([0, 0x61, 0x73, 0x6d]);
    const fetchImpl = vi.fn(async (_input: URL) => new Response(bytes));
    let blob: Blob | undefined;
    vi.spyOn(URL, 'createObjectURL').mockImplementation((value) => {
      blob = value as Blob;
      return 'blob:parser';
    });

    const url = await fetchParserWasmBlobUrl(
      './assets/docx_parser_bg-HASH.wasm',
      'https://file+.vscode-resource.vscode-cdn.net/ext/dist/',
      fetchImpl as unknown as typeof fetch,
    );

    expect(url).toBe('blob:parser');
    expect(String(fetchImpl.mock.calls[0]?.[0])).toBe(
      'https://file+.vscode-resource.vscode-cdn.net/ext/dist/assets/docx_parser_bg-HASH.wasm',
    );
    expect(blob?.type).toBe('application/wasm');
    expect(new Uint8Array(await blob!.arrayBuffer())).toEqual(bytes);
  });

  it('reports a failed response instead of handing the worker an error body', async () => {
    const fetchImpl = vi.fn(async () => new Response('Request Timeout', { status: 408 }));
    await expect(
      fetchParserWasmBlobUrl('./assets/x.wasm', 'https://host/dist/', fetchImpl as unknown as typeof fetch),
    ).rejects.toThrow('HTTP 408');
  });
});

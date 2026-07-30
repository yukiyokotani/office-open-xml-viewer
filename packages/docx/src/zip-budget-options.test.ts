import { describe, expect, it, vi } from 'vitest';
import { DocxDocument } from './document.js';
import { attachDocumentLayoutRuntime } from './layout/runtime-state.js';

describe('DocxDocument ZIP budget wiring', () => {
  it('forwards aggregate and entry-count limits to worker parsing', async () => {
    let request: Record<string, unknown> | undefined;
    const instance = Object.create(DocxDocument.prototype) as Record<string, unknown>;
    attachDocumentLayoutRuntime(instance, 0);
    instance._mode = 'worker';
    instance._bridge = {
      request: vi.fn(async (createRequest: (id: number) => Record<string, unknown>) => {
        request = createRequest(7);
        return { type: 'parsedMeta', id: 7, meta: {} };
      }),
    };

    await (
      instance as unknown as {
        _parse(
          data: ArrayBuffer,
          maxEntry: number,
          maxTotal: number,
          maxEntries: number,
          useGoogleFonts: boolean,
          timeout: number,
        ): Promise<void>;
      }
    )._parse(new ArrayBuffer(1), 64, 512, 20_000, false, 30_000);

    expect(request).toMatchObject({
      type: 'parse',
      id: 7,
      maxZipEntryBytes: 64,
      maxZipTotalBytes: 512,
      maxZipEntries: 20_000,
    });
  });
});

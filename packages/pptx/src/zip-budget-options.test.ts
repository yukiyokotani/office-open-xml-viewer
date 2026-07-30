import { describe, expect, it, vi } from 'vitest';
import { PptxPresentation } from './presentation.js';

describe('PptxPresentation ZIP budget wiring', () => {
  it('forwards aggregate and entry-count limits to worker parsing', async () => {
    let request: Record<string, unknown> | undefined;
    const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
    instance._mode = 'worker';
    instance._bridge = {
      request: vi.fn(async (createRequest: (id: number) => Record<string, unknown>) => {
        request = createRequest(11);
        return { kind: 'parsedMeta', id: 11, meta: {} };
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
      kind: 'parse',
      id: 11,
      maxZipEntryBytes: 64,
      maxZipTotalBytes: 512,
      maxZipEntries: 20_000,
    });
  });
});

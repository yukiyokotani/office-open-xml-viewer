import { describe, expect, it, vi } from 'vitest';
import { DocxDocument } from './document.js';
import { attachDocumentLayoutRuntime } from './layout/runtime-state.js';

describe('DocxDocument resource-policy wiring', () => {
  it('forwards one normalized policy object to worker parsing', async () => {
    const onUsage = vi.fn();
    let request: Record<string, unknown> | undefined;
    const instance = Object.create(DocxDocument.prototype) as Record<string, unknown>;
    attachDocumentLayoutRuntime(instance, 0);
    instance._mode = 'worker';
    instance._bridge = {
      request: vi.fn(async (createRequest: (id: number) => Record<string, unknown>) => {
        request = createRequest(7);
        return {
          type: 'parsedMeta',
          id: 7,
          meta: {
            pageCount: 0,
            revisions: [],
            comments: [],
            footnotes: [],
            endnotes: [],
            pageSizes: [],
            bookmarkPages: [],
          },
          usage: {
            archiveEntryCount: 3,
            declaredInflatedBytes: 4,
            largestInflatedEntryBytes: 5,
            distinctInflatedBytes: 6,
            operationInflatedBytes: 7,
          },
        };
      }),
    };
    const policy = {
      maxArchiveEntryBytes: 64,
      maxTotalInflatedBytes: 512,
      maxArchiveEntries: 17,
    } as const;

    await (
      instance as unknown as {
        _parse(
          data: ArrayBuffer,
          resourcePolicy: typeof policy,
          useGoogleFonts: boolean,
          timeout: number,
          onUsage: (usage: unknown) => void,
        ): Promise<void>;
      }
    )._parse(new ArrayBuffer(1), policy, false, 30_000, onUsage);

    expect(request).toMatchObject({
      type: 'parse',
      id: 7,
      resourcePolicy: policy,
    });
    expect(request).not.toHaveProperty('maxZipEntryBytes');
    expect(request).not.toHaveProperty('parserResourceLimits');
    expect(onUsage).toHaveBeenCalledWith(expect.objectContaining({
      largestInflatedEntryBytes: 5,
      distinctInflatedBytes: 6,
    }));
  });
});

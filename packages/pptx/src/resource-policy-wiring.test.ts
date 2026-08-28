import { describe, expect, it, vi } from 'vitest';
import { PptxPresentation } from './presentation.js';

describe('PptxPresentation resource-policy wiring', () => {
  it('forwards one normalized policy object to worker parsing', async () => {
    const onUsage = vi.fn();
    let request: Record<string, unknown> | undefined;
    const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
    instance._mode = 'worker';
    instance._bridge = {
      request: vi.fn(async (createRequest: (id: number) => Record<string, unknown>) => {
        request = createRequest(3);
        return {
          kind: 'presentationReady',
          id: 3,
          preflight: {
            slideCount: 0,
            slideWidth: 914400,
            slideHeight: 914400,
            defaultTextColor: null,
            majorFont: null,
            minorFont: null,
            hlinkColor: null,
            folHlinkColor: null,
            embeddedFonts: [],
            slides: [],
            fontPreloadNames: [],
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
      maxArchiveEntryBytes: null,
      maxTotalInflatedBytes: 512,
      maxArchiveEntries: 18,
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

    expect(request).toMatchObject({ kind: 'parse', id: 3, resourcePolicy: policy });
    expect(request).not.toHaveProperty('maxZipEntryBytes');
    expect(request).not.toHaveProperty('parserResourceLimits');
    expect(onUsage).toHaveBeenCalledWith(expect.objectContaining({
      largestInflatedEntryBytes: 5,
      distinctInflatedBytes: 6,
    }));
  });
});

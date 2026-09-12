import { describe, expect, it, vi } from 'vitest';
import { XlsxWorkbook } from './workbook.js';
import { BoundedRawPartCache } from '@silurus/ooxml-core/internal/bounded-raw-part-cache';

describe('XlsxWorkbook resource-policy wiring', () => {
  it('sends the normalized policy once and reuses the retained session for sheets', async () => {
    const onUsage = vi.fn();
    const requests: Record<string, unknown>[] = [];
    const instance = Object.create(XlsxWorkbook.prototype) as Record<string, unknown>;
    instance._mode = 'worker';
    instance.sheetCache = new Map();
    instance.sheetLoads = new Map();
    instance.imageCache = new Map();
    instance.rawParts = new BoundedRawPartCache({ maxEntries: 4, maxBytes: 1024 });
    instance.googleFontNames = [];
    instance.retainedFontSets = new Map();
    instance.fontsDestroyed = false;
    const bridge = {
      request: vi.fn(async (createRequest: (id: number) => Record<string, unknown>) => {
        const request = createRequest(requests.length + 1);
        requests.push(request);
        if (request.type === 'parse') {
          return {
            type: 'parsed',
            id: 1,
            workbook: {
              workbook: { sheets: [{ name: 'Sheet1' }] },
              styles: {},
              sharedStrings: [],
            },
            usage: {
              archiveEntryCount: 3,
              declaredInflatedBytes: 4,
              largestInflatedEntryBytes: 5,
              distinctInflatedBytes: 6,
              operationInflatedBytes: 7,
            },
          };
        }
        if (request.type === 'openSheetSession') {
          return { type: 'sheetSessionOpened', id: request.id };
        }
        if (request.kind === 'pull') {
          const payload = new TextEncoder().encode(JSON.stringify({
            kind: 'finished',
            worksheet: {
            name: 'Sheet1',
            rows: [],
            colWidths: {},
            rowHeights: {},
            defaultColWidth: 8,
            defaultRowHeight: 15,
            mergeCells: [],
            freezeRows: 0,
            freezeCols: 0,
            conditionalFormats: [],
            images: [],
            charts: [],
            },
          })).buffer;
          return { ...request, kind: 'chunk', done: true, byteLength: payload.byteLength, payload };
        }
        return { ...request, kind: 'accepted', command: request.kind };
      }),
      transport: () => bridge,
      forgetOrphaned: vi.fn(),
      terminate: vi.fn(),
    };
    instance.bridge = bridge;
    const policy = {
      maxArchiveEntryBytes: 64,
      maxTotalInflatedBytes: null,
      maxArchiveEntries: 19,
    } as const;

    await (
      instance as unknown as {
        _load(
          data: ArrayBuffer,
          options: Record<string, unknown>,
          resourcePolicy: typeof policy,
          onUsage: (usage: unknown) => void,
        ): Promise<void>;
        getWorksheet(index: number): Promise<unknown>;
      }
    )._load(new ArrayBuffer(1), {}, policy, onUsage);
    await (
      instance as unknown as { getWorksheet(index: number): Promise<unknown> }
    ).getWorksheet(0);

    expect(requests).toHaveLength(4);
    expect(requests[0]).toMatchObject({ type: 'parse', resourcePolicy: policy });
    expect(requests[1]).toMatchObject({ type: 'openSheetSession', sheetIndex: 0 });
    expect(requests[1]).not.toHaveProperty('resourcePolicy');
    expect(requests[1]).not.toHaveProperty('maxZipEntryBytes');
    expect(requests[1]).not.toHaveProperty('parserResourceLimits');
    expect(onUsage).toHaveBeenCalledWith(expect.objectContaining({
      largestInflatedEntryBytes: 5,
      distinctInflatedBytes: 6,
    }));
  });
});

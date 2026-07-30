import { describe, expect, it, vi } from 'vitest';
import { XlsxWorkbook } from './workbook.js';

describe('XlsxWorkbook ZIP budget wiring', () => {
  it('forwards the same limits to the initial parse and retained-archive sheet parse', async () => {
    const requests: Record<string, unknown>[] = [];
    const instance = Object.create(XlsxWorkbook.prototype) as Record<string, unknown>;
    instance._mode = 'worker';
    instance.sheetCache = new Map();
    instance.imageCache = new Map();
    instance.imageBlobCache = new Map();
    instance.googleFontFaces = [];
    instance.bridge = {
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
          };
        }
        return {
          type: 'parsedSheet',
          id: 2,
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
        };
      }),
    };

    await (
      instance as unknown as {
        _load(data: ArrayBuffer, options: Record<string, unknown>): Promise<void>;
        getWorksheet(index: number): Promise<unknown>;
      }
    )._load(new ArrayBuffer(1), {
      maxZipEntryBytes: 64,
      maxZipTotalBytes: 512,
      maxZipEntries: 20_000,
    });
    await (
      instance as unknown as { getWorksheet(index: number): Promise<unknown> }
    ).getWorksheet(0);

    expect(requests).toHaveLength(2);
    expect(requests[0]).toMatchObject({
      type: 'parse',
      maxZipEntryBytes: 64,
      maxZipTotalBytes: 512,
      maxZipEntries: 20_000,
    });
    expect(requests[1]).toMatchObject({
      type: 'parseSheet',
      maxZipEntryBytes: 64,
      maxZipTotalBytes: 512,
      maxZipEntries: 20_000,
    });
  });
});

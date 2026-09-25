import { afterEach, describe, expect, it, vi } from 'vitest';
import type { ModelSource } from '@silurus/ooxml-core';
import { XlsxViewer } from './viewer.js';
import { XlsxWorkbook } from './workbook.js';
import { installDom, makeContainer } from './viewer-destroy-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

// The viewer owns the load call: a model source it drops would silently send
// the input down the OOXML path instead.
describe('XlsxViewer load options', () => {
  it('forwards modelSources to XlsxWorkbook.load only when configured', async () => {
    installDom();
    const source = { target: 'xlsx', claim: () => false, beginLoad: vi.fn() } as unknown as ModelSource<'xlsx'>;
    const workbook = {
      sheetNames: ['Sheet1'], tabColors: {}, destroy: vi.fn(),
      getWorksheet: vi.fn().mockResolvedValue(undefined),
    } as unknown as XlsxWorkbook;
    const load = vi.spyOn(XlsxWorkbook, 'load').mockResolvedValue(workbook);
    for (const [options, expected] of [[{ modelSources: [source] }, [source]], [{}, undefined]] as const) {
      const viewer = new XlsxViewer(makeContainer() as unknown as HTMLElement, options);
      vi.spyOn(viewer as unknown as { showSheet(index: number): Promise<void> }, 'showSheet')
        .mockResolvedValue(undefined);
      await viewer.load(new ArrayBuffer(1));
      expect(load.mock.lastCall?.[1]?.modelSources).toEqual(expected);
      viewer.destroy();
    }
  });
});

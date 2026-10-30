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
  it('forwards legacyXls to XlsxWorkbook.load only when configured', async () => {
    installDom();
    const source = { target: 'xlsx', claim: () => false, beginLoad: vi.fn() } as unknown as ModelSource<'xlsx'>;
    const workbook = {
      sheetNames: ['Sheet1'], tabColors: {}, destroy: vi.fn(),
      getWorksheet: vi.fn().mockResolvedValue(undefined),
    } as unknown as XlsxWorkbook;
    const load = vi.spyOn(XlsxWorkbook, 'load').mockResolvedValue(workbook);
    const ods = { ...source } as typeof source;
    for (const [options, legacyXls, expectedOds] of [
      [{ legacyXls: source }, source, undefined],
      [{ ods }, undefined, ods],
      [{ legacyXls: source, ods }, source, ods],
      [{}, undefined, undefined],
    ] as const) {
      const viewer = new XlsxViewer(makeContainer() as unknown as HTMLElement, options);
      vi.spyOn(viewer as unknown as { showSheet(index: number): Promise<void> }, 'showSheet')
        .mockResolvedValue(undefined);
      await viewer.load(new ArrayBuffer(1));
      expect(load.mock.lastCall?.[1]?.legacyXls).toEqual(legacyXls);
      expect(load.mock.lastCall?.[1]?.ods).toEqual(expectedOds);
      viewer.destroy();
    }
  });
});

// Each surface takes only a source for its own format, under its own key.
describe('XLSX source option typing', () => {
  it('rejects another format\'s source and the removed modelSources option', () => {
    const docx = { target: 'docx', claim: () => false, beginLoad: vi.fn() } as unknown as ModelSource<'docx'>;
    const pptx = { target: 'pptx', claim: () => false, beginLoad: vi.fn() } as unknown as ModelSource<'pptx'>;
    // @ts-expect-error A DOCX source cannot configure a XLSX surface.
    const xlsxViewerDocx: ConstructorParameters<typeof XlsxViewer>[1] = { legacyXls: docx };
    // @ts-expect-error A PPTX source cannot configure a XLSX surface.
    const xlsxViewerPptx: ConstructorParameters<typeof XlsxViewer>[1] = { legacyXls: pptx };
    // @ts-expect-error The generic list option was replaced by `legacyXls`.
    const xlsxViewerList: ConstructorParameters<typeof XlsxViewer>[1] = { modelSources: [] };
    // @ts-expect-error A DOCX source cannot load XLSX input.
    const loadDocx: NonNullable<Parameters<typeof XlsxWorkbook.load>[1]> = { legacyXls: docx };
    // @ts-expect-error A PPTX source cannot load XLSX input.
    const loadPptx: NonNullable<Parameters<typeof XlsxWorkbook.load>[1]> = { legacyXls: pptx };
    expect([xlsxViewerDocx, xlsxViewerPptx, xlsxViewerList, loadDocx, loadPptx]).toHaveLength(5);
  });
});

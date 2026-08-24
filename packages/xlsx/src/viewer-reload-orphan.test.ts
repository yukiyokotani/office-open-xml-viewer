import { describe, it, expect, afterEach, vi } from 'vitest';
import { XlsxViewer, type XlsxViewerOptions } from './viewer.js';
import { XlsxWorkbook } from './workbook.js';
import type { Worksheet } from './types.js';
import { installDom, makeContainer } from './viewer-destroy-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/** A minimal XlsxWorkbook stand-in covering exactly the surface the viewer's
 *  load() → buildTabs() path touches, plus a `destroy` spy so SC20 can assert the
 *  previous workbook (its worker + WASM) is torn down on a re-load. */
function fakeWorkbook() {
  const destroy = vi.fn();
  const wb = {
    sheetNames: ['Sheet1'],
    tabColors: {} as Record<number, string>,
    destroy,
    getWorksheet: vi.fn().mockResolvedValue(undefined),
  };
  return { wb: wb as unknown as XlsxWorkbook, destroy };
}

function worksheet(name: string): Worksheet {
  return {
    name,
    rows: [],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 64,
    defaultRowHeight: 20,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    charts: [],
    images: [],
    shapeGroups: [],
  } as unknown as Worksheet;
}

/**
 * SC20: a second `load()` must not orphan the previous workbook (its worker +
 * pinned WASM). The fix is an atomic swap — load the new workbook first, destroy
 * the old one only on success.
 */
describe('XlsxViewer.load() — no orphaned workbook on re-load (SC20)', () => {
  function build(opts: XlsxViewerOptions = {}) {
    installDom();
    const container = makeContainer();
    const v = new XlsxViewer(container as unknown as HTMLElement, opts);
    // Isolate SC20 from the sheet-render path: showSheet needs a full worksheet
    // model to lay out, which is out of scope here. The engine-swap happens in
    // load() BEFORE showSheet, so a resolved no-op keeps this test on the leak.
    const showSheet = vi.spyOn(
      v as unknown as { showSheet: (i: number) => Promise<void> },
      'showSheet',
    ).mockResolvedValue(undefined);
    return { v, showSheet };
  }

  it('destroys the previous workbook when load() is called again', async () => {
    const { v } = build();
    const a = fakeWorkbook();
    const b = fakeWorkbook();
    const loadSpy = vi
      .spyOn(XlsxWorkbook, 'load')
      .mockResolvedValueOnce(a.wb)
      .mockResolvedValueOnce(b.wb);

    await v.load('one.xlsx');
    expect(a.destroy).not.toHaveBeenCalled();

    await v.load('two.xlsx');
    expect(loadSpy).toHaveBeenCalledTimes(2);
    expect(a.destroy).toHaveBeenCalledTimes(1);
    expect(b.destroy).not.toHaveBeenCalled();

    v.destroy();
    expect(b.destroy).toHaveBeenCalledTimes(1);
  });

  it('forwards the opt-in ChartEx renderer to the workbook load', async () => {
    const chartEx = { render: vi.fn() };
    const { v } = build({ chartEx });
    const loaded = fakeWorkbook();
    const loadSpy = vi.spyOn(XlsxWorkbook, 'load').mockResolvedValueOnce(loaded.wb);

    await v.load('chartex.xlsx');

    expect(loadSpy).toHaveBeenCalledWith(
      'chartex.xlsx',
      expect.objectContaining({ chartEx }),
    );
    v.destroy();
  });

  it('does not report an old worksheet acquisition rejected by a successful reload', async () => {
    installDom();
    const onError = vi.fn();
    const v = new XlsxViewer(makeContainer() as unknown as HTMLElement, { onError });
    const internals = v as unknown as Record<string, unknown> & {
      showSheet(index: number): Promise<void>;
    };
    for (const method of [
      'hideCommentPopup', 'hideValidationPanel', 'updateSelectionOverlay', 'updateTabActive',
      'buildCommentMap', 'buildHyperlinkMap', 'buildOutline', 'layoutGutters', 'updateSpacerSize',
      'resetHorizontalScroll', 'updateFindOverlay', 'emitViewportChange',
    ]) internals[method] = vi.fn();
    internals.renderCurrentSheet = vi.fn(async () => undefined);

    let rejectOldRequest!: (error: Error) => void;
    const oldRequest = new Promise<Worksheet>((_resolve, reject) => { rejectOldRequest = reject; });
    const old = fakeWorkbook();
    const next = fakeWorkbook();
    const oldGetWorksheet = old.wb.getWorksheet as ReturnType<typeof vi.fn>;
    oldGetWorksheet.mockResolvedValueOnce(worksheet('Old'));
    const nextGetWorksheet = next.wb.getWorksheet as ReturnType<typeof vi.fn>;
    nextGetWorksheet.mockResolvedValue(worksheet('Next'));
    vi.spyOn(XlsxWorkbook, 'load')
      .mockResolvedValueOnce(old.wb)
      .mockResolvedValueOnce(next.wb);

    await v.load('old.xlsx');
    oldGetWorksheet.mockImplementationOnce(() => oldRequest);
    const staleNavigation = internals.showSheet(0);

    await v.load('next.xlsx');
    rejectOldRequest(new Error('old worker terminated'));
    await staleNavigation;

    expect(old.destroy).toHaveBeenCalledOnce();
    expect(onError).not.toHaveBeenCalled();
    v.destroy();
  });

  it('keeps the current workbook when the re-load fails (atomic swap)', async () => {
    const onError = vi.fn();
    const { v } = build({ onError });
    const a = fakeWorkbook();
    vi.spyOn(XlsxWorkbook, 'load')
      .mockResolvedValueOnce(a.wb)
      .mockRejectedValueOnce(new Error('boom'));

    await v.load('one.xlsx');

    await expect(v.load('bad.xlsx')).rejects.toThrow('boom');
    expect(onError).not.toHaveBeenCalled();
    expect(a.destroy).not.toHaveBeenCalled();

    v.destroy();
    expect(a.destroy).toHaveBeenCalledTimes(1);
  });

  it('rejects an initial sheet-render failure without also calling onError', async () => {
    const onError = vi.fn();
    const { v, showSheet } = build({ onError });
    const a = fakeWorkbook();
    vi.spyOn(XlsxWorkbook, 'load').mockResolvedValueOnce(a.wb);
    showSheet.mockRejectedValueOnce(new Error('initial render boom'));

    await expect(v.load('one.xlsx')).rejects.toThrow('initial render boom');
    expect(onError).not.toHaveBeenCalled();

    v.destroy();
    expect(a.destroy).toHaveBeenCalledTimes(1);
  });

  it('maps an in-flight loader rejection after destroy to the terminal viewer error', async () => {
    const { v } = build();
    let rejectLoad: ((reason: Error) => void) | undefined;
    vi.spyOn(XlsxWorkbook, 'load').mockImplementation(() =>
      new Promise<XlsxWorkbook>((_resolve, reject) => { rejectLoad = reject; }));

    const pending = v.load('pending.xlsx');
    v.destroy();
    rejectLoad?.(new Error('late parser failure'));

    await expect(pending).rejects.toThrow('XlsxViewer is destroyed');
  });

  it('disposes an in-flight workbook that resolves after destroy and reports terminal close', async () => {
    const { v } = build();
    const late = fakeWorkbook();
    let resolveLoad: ((workbook: XlsxWorkbook) => void) | undefined;
    vi.spyOn(XlsxWorkbook, 'load').mockImplementation(() =>
      new Promise<XlsxWorkbook>((resolve) => { resolveLoad = resolve; }));

    const pending = v.load('pending.xlsx');
    v.destroy();
    resolveLoad?.(late.wb);

    await expect(pending).rejects.toThrow('XlsxViewer is destroyed');
    expect(late.destroy).toHaveBeenCalledOnce();
  });
});

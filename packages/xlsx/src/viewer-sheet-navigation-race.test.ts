import { afterEach, describe, expect, it, vi } from 'vitest';
import { XlsxViewer } from './viewer.js';
import { XlsxWorkbook } from './workbook.js';
import { WorksheetPreview } from './internal/worksheet-preview.js';
import type { Worksheet } from './types.js';
import { installDom, makeContainer, type FakeEl } from './viewer-destroy-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

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

function deferred<T>(): { readonly promise: Promise<T>; readonly resolve: (value: T) => void } {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((accept) => { resolve = accept; });
  return { promise, resolve };
}

function buildViewer(onSheetChange = vi.fn(), names = ['A', 'B']) {
  installDom();
  const container = makeContainer();
  const viewer = new XlsxViewer(container as unknown as HTMLElement, { onSheetChange });
  const requests = names.map(() => deferred<Worksheet>());
  const workbook = {
    sheetNames: names,
    sheetCount: names.length,
    getWorksheet: vi.fn((index: number) => requests[index].promise),
    acquireWorksheetLease: vi.fn(async (index: number) => ({
      worksheet: await requests[index].promise,
      release: vi.fn(),
    })),
    retainWorksheetReference: vi.fn(async (_index: number) => vi.fn()),
    destroy: vi.fn(),
  };
  const engine = viewer as unknown as Record<string, unknown> & {
    wb: unknown;
    showSheet(index: number): Promise<void>;
    currentSheet: number;
    currentWorksheet: Worksheet | null;
  };
  engine.wb = workbook;
  for (const method of [
    'hideCommentPopup', 'hideValidationPanel', 'updateSelectionOverlay', 'updateTabActive',
    'buildCommentMap', 'buildHyperlinkMap', 'buildOutline', 'layoutGutters', 'updateSpacerSize',
    'resetHorizontalScroll', 'updateFindOverlay', 'emitViewportChange',
  ]) engine[method] = vi.fn();
  engine.renderCurrentSheet = vi.fn(async () => {
    // The render double represents a committed first frame.
    engine.firstPreviewRender = false;
    engine.committedFrameCount = Number(engine.committedFrameCount) + 1;
  });
  return { viewer, engine, workbook, requests, onSheetChange, container };
}

describe('XlsxViewer sheet acquisition generation', () => {
  it('covers the viewport after saved hidden rows are restored while the pull continues', async () => {
    const { viewer, engine, workbook } = buildViewer(vi.fn(), ['A']);
    delete engine.renderCurrentSheet;
    const model = worksheet('A');
    const progress = new WorksheetPreview([]);
    progress.preview(model, null, 1_000, 1);
    progress.append(Array.from({ length: 128 }, (_, i) => ({ index: i + 1, height: null, cells: [] })));
    const completion = deferred<Worksheet>();
    const sizes = (engine.viewEdits as { sizeOverrideStore: Map<number, unknown> }).sizeOverrideStore;
    sizes.set(0, {
      rows: new Map(Array.from({ length: 200 }, (_, i) => [i + 1, 0])),
      automaticRows: new Map(), cols: new Map(), colCss: new Map(), revision: 1,
    });
    const area = engine.canvasArea as { clientWidth: number; clientHeight: number };
    area.clientWidth = 800;
    area.clientHeight = 600;
    let requestedRow = 0;
    let renderRow = 0;
    Object.assign(workbook, {
      acquireWorksheetPreviewLease: async () => ({
        worksheet: model, release: vi.fn(), partial: true, completion: completion.promise,
        waitForRows: (row: number) => { requestedRow = row; return progress.waitFor(row); },
      }),
      renderViewport: async (_target: unknown, _index: number, range: { row: number; rows: number }) => {
        renderRow = range.row + range.rows - 1;
        await progress.waitFor(renderRow);
      },
    });

    const shown = engine.showSheet(0);
    // Subsequent row chunks do not wait for the provisional render. In the
    // old handshake, restoring rows 1–200 as hidden made paint wait for rows
    // beyond the paused chunk and neither side could resume.
    progress.append(Array.from({ length: 128 }, (_, i) => ({ index: i + 129, height: null, cells: [] })));
    await shown;
    expect(requestedRow).toBeGreaterThan(128);
    expect(renderRow).toBeLessThanOrEqual(progress.coveredThrough);
    progress.finish(model);
    completion.resolve(model);
    viewer.destroy();
  });

  it.each([
    ['frozen row', 1, 0, 1, 2],
    ['frozen column', 0, 1, 2, 1],
    ['frozen corner', 1, 1, 1, 1],
  ] as const)('checks conditional formatting in the %s before provisional paint',
    async (_region, freezeRows, freezeCols, row, col) => {
    const { viewer, engine, workbook } = buildViewer(vi.fn(), ['A']);
    const model = worksheet('A');
    model.freezeRows = freezeRows;
    model.freezeCols = freezeCols;
    model.conditionalFormats = [{
      sqref: [{ top: row, left: col, bottom: row, right: col }],
      rules: [{ type: 'cellIs', operator: 'greaterThan', formulas: ['A500'], dxfId: 0, priority: 1 }],
    }];
    Object.assign(workbook, {
      acquireWorksheetPreviewLease: async () => ({
        worksheet: model, release: vi.fn(), partial: true,
        completion: Promise.resolve(model), waitForRows: async () => undefined,
      }),
    });
    await engine.showSheet(0);
    expect(engine.previewFallbackReason).toBe('conditional-format-range');
    viewer.destroy();
  });

  it('waits for a committed frame when completion supersedes the first paint', async () => {
    const { viewer, engine, workbook } = buildViewer(vi.fn(), ['A']);
    const completion = deferred<Worksheet>();
    const model = worksheet('A');
    Object.assign(workbook, {
      acquireWorksheetPreviewLease: vi.fn(async () => ({
        worksheet: model, release: vi.fn(), partial: true,
        completion: completion.promise, waitForRows: vi.fn(async () => undefined),
        releaseFirstPaint: vi.fn(),
      })),
    });
    engine.scheduleRender = vi.fn();
    let attempts = 0;
    engine.renderCurrentSheet = vi.fn(async () => {
      attempts++;
      if (attempts === 1) {
        completion.resolve(model);
        await Promise.resolve();
        // The completion callback has invalidated this in-flight frame.
        return;
      }
      engine.committedFrameCount = Number(engine.committedFrameCount) + 1;
    });

    await engine.showSheet(0);
    expect(engine.renderCurrentSheet).toHaveBeenCalledTimes(2);
    expect(attempts).toBe(2);
    viewer.destroy();
  });

  it('clears a provisional frame and reports a pull failure after first paint', async () => {
    const { viewer, engine, workbook } = buildViewer(vi.fn(), ['A']);
    const onError = vi.fn();
    (engine.opts as { onError?: (error: Error) => void }).onError = onError;
    let rejectCompletion!: (error: Error) => void;
    const completion = new Promise<Worksheet>((_, reject) => { rejectCompletion = reject; });
    const release = vi.fn();
    Object.assign(workbook, {
      acquireWorksheetPreviewLease: vi.fn(async () => ({
        worksheet: worksheet('A'), release, partial: true, completion,
        waitForRows: vi.fn(async () => undefined),
      })),
    });

    await engine.showSheet(0);
    expect(engine.currentWorksheet).not.toBeNull();
    const error = new Error('row pull failed');
    rejectCompletion(error);
    await Promise.resolve();
    await Promise.resolve();
    expect(engine.currentWorksheet).toBeNull();
    expect(release).toHaveBeenCalledOnce();
    expect(onError).toHaveBeenCalledWith(error);
    viewer.destroy();
  });

  it('replaces a provisional sheet with the parser placeholder after a late row error', async () => {
    const { viewer, engine, workbook } = buildViewer(vi.fn(), ['A']);
    const completion = deferred<Worksheet>();
    Object.assign(workbook, {
      acquireWorksheetPreviewLease: vi.fn(async () => ({
        worksheet: worksheet('A'), release: vi.fn(), partial: true,
        completion: completion.promise, waitForRows: vi.fn(async () => undefined),
        releaseFirstPaint: vi.fn(),
      })),
    });
    engine.scheduleRender = vi.fn();

    await engine.showSheet(0);
    completion.resolve({ ...worksheet('A'), parseError: 'invalid later cell' });
    await Promise.resolve();
    await Promise.resolve();
    expect(engine.currentWorksheet?.parseError).toBe('invalid later cell');
    expect(engine.scheduleRender).toHaveBeenCalled();
    viewer.destroy();
  });

  it('ignores an old provisional completion after sheet navigation', async () => {
    const { viewer, engine, workbook } = buildViewer();
    const first = deferred<Worksheet>();
    const firstSheet = worksheet('A');
    const secondSheet = worksheet('B');
    Object.assign(workbook, {
      acquireWorksheetPreviewLease: vi.fn(async (index: number) => index === 0
        ? { worksheet: firstSheet, release: vi.fn(), partial: true,
            completion: first.promise, waitForRows: vi.fn(async () => undefined),
            releaseFirstPaint: vi.fn() }
        : { worksheet: secondSheet, release: vi.fn(), partial: false,
            completion: Promise.resolve(secondSheet) }),
    });

    await engine.showSheet(0);
    await engine.showSheet(1);
    first.resolve(firstSheet);
    await Promise.resolve();
    expect(engine.currentSheet).toBe(1);
    expect(engine.currentWorksheet).toMatchObject(secondSheet);
    viewer.destroy();
  });

  it('releases the outgoing active model so a one-sheet cache can navigate', async () => {
    const { viewer, engine, workbook } = buildViewer();
    let held: number | null = null;
    workbook.acquireWorksheetLease.mockImplementation(async (index: number) => {
      if (held !== null && held !== index) throw new Error('outgoing sheet still leased');
      held = index;
      return {
        worksheet: worksheet(index === 0 ? 'A' : 'B'),
        release: vi.fn(() => { if (held === index) held = null; }),
      };
    });

    await engine.showSheet(0);
    await engine.showSheet(1);

    expect(engine.currentSheet).toBe(1);
    expect(held).toBe(1);
    viewer.destroy();
    expect(held).toBeNull();
  });

  it('commits the newest worksheet and index atomically when an older request resolves late', async () => {
    const { viewer, engine, requests, onSheetChange } = buildViewer();
    const a = worksheet('A');
    const b = worksheet('B');

    const first = engine.showSheet(0);
    const second = engine.showSheet(1);
    requests[1].resolve(b);
    await second;
    requests[0].resolve(a);
    await first;

    expect(engine.currentSheet).toBe(1);
    expect(engine.currentWorksheet).not.toBe(b);
    expect(engine.currentWorksheet).toMatchObject(b);
    expect(onSheetChange).toHaveBeenCalledOnce();
    expect(onSheetChange).toHaveBeenCalledWith(1, 2);
    viewer.destroy();
  });

  it('drops a worksheet that resolves after destroy without callbacks or model commit', async () => {
    const { viewer, engine, requests, onSheetChange } = buildViewer();
    const pending = engine.showSheet(0);
    viewer.destroy();
    requests[0].resolve(worksheet('A'));
    await pending;

    expect(engine.currentWorksheet).toBeNull();
    expect(onSheetChange).not.toHaveBeenCalled();
  });

  it('keeps the displayed sheet and its interaction state when replacement loading fails', async () => {
    const { viewer, engine, workbook, requests, onSheetChange } = buildViewer();
    requests[0].resolve(worksheet('A'));
    let held = 0;
    workbook.acquireWorksheetLease.mockImplementation(async (index: number) => {
      if (index === 1) throw new Error('replacement load failed');
      held++;
      return { worksheet: await requests[0].promise, release: vi.fn(() => { held--; }) };
    });
    workbook.retainWorksheetReference.mockImplementation(async () => {
      held++;
      return vi.fn(() => { held--; });
    });
    await engine.showSheet(0);
    viewer.setSelection('A1');
    const displayed = engine.currentWorksheet;

    await expect(engine.showSheet(1)).rejects.toThrow('replacement load failed');
    expect(engine.currentSheet).toBe(0);
    expect(engine.currentWorksheet).toBe(displayed);
    expect(held).toBe(1);
    expect(viewer.getSelectionContext()).toMatchObject({ sheetIndex: 0, sheetName: 'A' });
    expect(onSheetChange).toHaveBeenCalledTimes(1);
    viewer.destroy();
    expect(held).toBe(0);
  });

  it('keeps selection and comments interactive during a slow sheet load', async () => {
    const { viewer, engine, requests } = buildViewer();
    const current = { ...worksheet('A'), comments: [{ cellRef: 'A1', text: 'note' }] };
    requests[0].resolve(current);
    await engine.showSheet(0);

    const loading = engine.showSheet(1);
    await Promise.resolve();
    viewer.setSelection('A1');
    expect(viewer.getSelectionContext()).toMatchObject({ sheetIndex: 0, sheetName: 'A' });
    expect(viewer.getComments()).toEqual([{ cellRef: 'A1', text: 'note' }]);
    requests[1].resolve(worksheet('B'));
    await loading;
    viewer.destroy();
  });

  it('preserves the old sheet when a pending navigation is superseded', async () => {
    const { viewer, engine, requests, onSheetChange } = buildViewer(vi.fn(), ['A', 'B', 'C']);
    requests[0].resolve(worksheet('A'));
    await engine.showSheet(0);
    viewer.setSelection('A1');
    const displayed = engine.currentWorksheet;
    const slow = engine.showSheet(1);
    const newest = engine.showSheet(2);
    requests[1].resolve(worksheet('B'));
    await slow;
    expect(engine.currentWorksheet).toBe(displayed);
    expect(viewer.getSelectionContext()).toMatchObject({ sheetIndex: 0, sheetName: 'A' });
    requests[2].resolve(worksheet('C'));
    await newest;
    expect(engine.currentSheet).toBe(2);
    expect(onSheetChange).toHaveBeenCalledTimes(2);
    viewer.destroy();
  });

  it('completes teardown after its workbook has already destroyed itself', () => {
    const { viewer, engine, container } = buildViewer();
    const WorkbookConstructor = XlsxWorkbook as unknown as new (
      worker: Worker | null, mode: 'worker', wasmUrl?: string | URL,
    ) => XlsxWorkbook;
    const workbook = new WorkbookConstructor(null, 'worker');
    engine.wb = workbook;
    workbook.destroy();

    expect(() => viewer.destroy()).not.toThrow();
    expect(container.children).toHaveLength(0);
    expect(engine.currentWorksheet).toBeNull();
  });

  it('retains resized row geometry across A → B → A', async () => {
    const { viewer, engine, requests } = buildViewer();
    requests[0].resolve(worksheet('A'));
    requests[1].resolve(worksheet('B'));
    await engine.showSheet(0);
    engine.scheduleRender = vi.fn();
    const area = engine.canvasArea as { clientWidth: number; clientHeight: number };
    area.clientWidth = 800; area.clientHeight = 600;
    const input = engine.scrollHost as FakeEl;
    input.clientWidth = 800; input.clientHeight = 600;
    const cell = viewer.getCellViewportRect('A1')!;
    const pointer = (y: number) => ({ button: 0, pointerId: 1, pointerType: 'mouse',
      clientX: 5, clientY: y, preventDefault() {}, shiftKey: false, ctrlKey: false, metaKey: false });
    input.dispatch('pointerdown', pointer(cell.y + cell.height + 0.5));
    input.dispatch('pointermove', pointer(cell.y + 120));
    input.dispatch('pointerup', pointer(cell.y + 120));
    const resized = (engine.currentWorksheet as Worksheet).rowHeights[1];
    expect(resized).toBeGreaterThan(20);

    await engine.showSheet(1);
    await engine.showSheet(0);
    expect((engine.currentWorksheet as Worksheet).rowHeights[1]).toBe(resized);
    expect((engine.wireSizeOverrides as () => { overrides: { rows: Record<number, number> } })()
      .overrides.rows[1]).toBe(resized);
    viewer.destroy();
  });

  it('retains resized column geometry across A → B → A', async () => {
    const { viewer, engine, requests } = buildViewer();
    requests[0].resolve({ ...worksheet('A'), defaultColWidth: 8.43 });
    requests[1].resolve(worksheet('B'));
    await engine.showSheet(0);
    engine.scheduleRender = vi.fn();
    const area = engine.canvasArea as { clientWidth: number; clientHeight: number };
    area.clientWidth = 800; area.clientHeight = 600;
    const input = engine.scrollHost as FakeEl;
    input.clientWidth = 800; input.clientHeight = 600;
    const cell = viewer.getCellViewportRect('B1')!;
    expect(cell.x + cell.width).toBeLessThan(800);
    const pointer = (x: number) => ({ button: 0, pointerId: 1, pointerType: 'mouse',
      clientX: x, clientY: 5, preventDefault() {}, shiftKey: false, ctrlKey: false, metaKey: false });
    input.dispatch('pointerdown', pointer(cell.x + cell.width + 0.5));
    input.dispatch('pointermove', pointer(cell.x + 120));
    input.dispatch('pointerup', pointer(cell.x + 120));
    const resized = (engine.currentWorksheet as Worksheet).colWidths[2];
    expect(resized).toBeGreaterThan(0);

    await engine.showSheet(1);
    await engine.showSheet(0);
    expect((engine.currentWorksheet as Worksheet).colWidths[2]).toBe(resized);
    expect((engine.wireSizeOverrides as () => { overrides: { cols: Record<number, number> } })()
      .overrides.cols[2]).toBe(resized);
    viewer.destroy();
  });

  it('retains outline collapse and pre-collapse sizes across A → B → A', async () => {
    const { viewer, engine, requests } = buildViewer();
    requests[0].resolve({
      ...worksheet('A'),
      rows: [
        { index: 1, height: null, cells: [], outlineLevel: 1 },
        { index: 2, height: null, cells: [], collapsed: false },
      ],
      rowHeights: { 1: 31 },
      colWidths: { 1: 12 },
      colOutlineLevels: { 1: 1 },
      colCollapsed: {},
    } as Worksheet);
    requests[1].resolve(worksheet('B'));
    delete engine.buildOutline;
    await engine.showSheet(0);
    (engine.setBandHidden as (axis: 'row' | 'col', index: number, hidden: boolean) => void)('row', 1, true);
    (engine.setBandHidden as (axis: 'row' | 'col', index: number, hidden: boolean) => void)('col', 1, true);
    (engine.setBandCollapsed as (axis: 'row' | 'col', index: number, collapsed: boolean) => void)('row', 2, true);
    (engine.setBandCollapsed as (axis: 'row' | 'col', index: number, collapsed: boolean) => void)('col', 2, true);

    await engine.showSheet(1);
    await engine.showSheet(0);
    const restored = engine.currentWorksheet as Worksheet;
    expect(restored.rowHeights[1]).toBe(0);
    expect(restored.colWidths[1]).toBe(0);
    expect(restored.rows.find((row) => row.index === 2)?.collapsed).toBe(true);
    expect(restored.colCollapsed?.[2]).toBe(true);
    (engine.setBandHidden as (axis: 'row' | 'col', index: number, hidden: boolean) => void)('row', 1, false);
    (engine.setBandHidden as (axis: 'row' | 'col', index: number, hidden: boolean) => void)('col', 1, false);
    expect(restored.rowHeights[1]).toBe(31);
    expect(restored.colWidths[1]).toBe(12);
    viewer.destroy();
  });
});

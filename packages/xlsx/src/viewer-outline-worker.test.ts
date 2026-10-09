import { describe, it, expect, afterEach, vi } from 'vitest';
import { XlsxViewer } from './viewer.js';
import { installDom, makeContainer, type FakeEl } from './viewer-destroy-test-dom.js';
import {
  applySizeOverrides,
  createSizeOverriddenWorksheet,
  WorksheetViewProjectionCache,
  type WireSizeOverrides,
} from './worker-protocol.js';
import { getSheetRenderCache, inheritSheetRenderCache } from './renderer.js';
import type { Worksheet } from './types.js';
import type { OutlineLayout } from './outline.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * Worker-mode model-sync for view-only size mutations.
 *
 * The render worker draws from its own worker-local parsed-sheet cache
 * (the render worker's worksheet-cursor cache), so a main-thread Worksheet mutation —
 * outline collapse/expand mapping bands to the size-0 hidden encoding, or
 * drag-to-resize (#567) — never reaches the worker on its own: the gutter and
 * overlays would update while the grid bitmap kept the file's original sizes
 * (rows stayed hidden after an expand click). These tests pin the override
 * channel that closes the gap:
 *
 * 1. every worker `renderViewport` request issued after a mutation carries
 *    `opts.sizeOverrides` describing the touched bands' CURRENT model sizes, and
 * 2. applying that wire payload to a pristine copy of the sheet (exactly what
 *    the worker does before drawing) converges its size maps to the main-thread
 *    model — the main↔worker equivalence that makes both modes lay out the
 *    same grid.
 */

/** The synthetic outline fixture's row model: rows 2-9 grouped (3 nested
 *  levels), detail rows 4-7 collapsed-hidden (height 0), row 8 the collapsed
 *  level-2 summary. Mirrors outline-fixture.xlsx. */
function outlineWorksheet(): Worksheet {
  const row = (
    index: number,
    outlineLevel = 0,
    opts: { collapsed?: boolean; hidden?: boolean } = {},
  ) => ({ index, height: opts.hidden ? 0 : null, cells: [], outlineLevel, ...opts });
  return {
    name: 'Outline',
    rows: [
      row(1),
      row(2, 1),
      row(3, 2),
      row(4, 3, { hidden: true }),
      row(5, 3, { hidden: true }),
      row(6, 3, { hidden: true }),
      row(7, 3, { hidden: true }),
      row(8, 2, { collapsed: true }),
      row(9, 1),
      row(10),
    ],
    colWidths: {},
    rowHeights: { 4: 0, 5: 0, 6: 0, 7: 0 },
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    charts: [],
    images: [],
    shapeGroups: [],
    outlinePr: { summaryBelow: true, summaryRight: true },
  } as unknown as Worksheet;
}

/** Private-surface shape of the viewer these tests drive. */
interface ViewerPriv {
  wb: unknown;
  currentWorksheet: Worksheet;
  currentSheet: number;
  canvasArea: FakeEl;
  scrollHost: FakeEl;
  outlineGutter: {
    rowOutline: OutlineLayout | null;
    applyGroupToggle(group: OutlineLayout['groups'][number], axis: 'row' | 'col'): void;
  };
  scrollOutlineSummaryToStart(axis: 'row' | 'col', summary: number): void;
  selectionInput: {
    resizeDrag: { kind: 'col' | 'row'; index: number; originScaled: number; mdw: number; indices: readonly number[]; worksheet: Worksheet } | null;
    applyResize(clientX: number, clientY: number): void;
  };
  buildOutline(ws: Worksheet): void;
  renderCurrentSheet(): Promise<void>;
}

/** Worker-mode viewer over the outline worksheet, with a controllable first
 *  bitmap so the latest-only render queue can be advanced deterministically. */
function buildWorker(onError?: (error: Error) => void) {
  installDom();
  const container = makeContainer();
  const v = new XlsxViewer(container as unknown as HTMLElement, { mode: 'worker', onError });
  const renderGate = deferred<ImageBitmap>();
  const renderViewportToBitmap = vi.fn(() => renderGate.promise);
  const fakeWb = {
    renderViewportToBitmap,
    sheetNames: ['Outline'],
    sheetCount: 1,
    destroy: vi.fn(),
  };
  const priv = v as unknown as ViewerPriv;
  priv.wb = fakeWb;
  priv.currentWorksheet = outlineWorksheet();
  priv.currentSheet = 0;
  priv.canvasArea.clientWidth = 800;
  priv.canvasArea.clientHeight = 600;
  priv.buildOutline(priv.currentWorksheet);
  return { v, priv, renderViewportToBitmap, completeRender: renderGate.resolve };
}

function deferred<T>(): { promise: Promise<T>; resolve(value: T): void } {
  let resolvePromise: (value: T) => void = () => undefined;
  const promise = new Promise<T>((resolve) => { resolvePromise = resolve; });
  return { promise, resolve: resolvePromise };
}

async function settleRenders(): Promise<void> {
  for (let index = 0; index < 8; index++) await Promise.resolve();
}

/** The `sizeOverrides` of the most recent render request, if any. */
function lastOverrides(fn: ReturnType<typeof vi.fn>): WireSizeOverrides | undefined {
  const call = fn.mock.calls.at(-1) as unknown[] | undefined;
  return (call?.[2] as { sizeOverrides?: WireSizeOverrides } | undefined)?.sizeOverrides;
}

describe('worker-mode outline collapse/expand reaches the grid bitmap', () => {
  it('expanding the collapsed group sends row overrides that reveal rows 4-7', () => {
    const { priv, renderViewportToBitmap } = buildWorker();
    const l3 = priv.outlineGutter.rowOutline?.groups.find((g) => g.level === 3);
    expect(l3?.collapsed).toBe(true);

    // Expand (the sync-render fallback fires renderCurrentSheet immediately).
    priv.outlineGutter.applyGroupToggle(l3 as OutlineLayout['groups'][number], 'row');
    expect(renderViewportToBitmap).toHaveBeenCalled();
    const o = lastOverrides(renderViewportToBitmap);
    // Rows 4-7 were revealed: the model has NO entry (default height) ⇒ null.
    expect(o?.rows).toMatchObject({ 4: null, 5: null, 6: null, 7: null });

    // Worker-side application converges a pristine sheet to the main model —
    // the property that makes the worker's next bitmap actually re-lay the
    // revealed rows (main↔worker equivalence at the model level).
    const workerSheet = outlineWorksheet();
    applySizeOverrides(workerSheet, o);
    expect(workerSheet.rowHeights).toEqual(priv.currentWorksheet.rowHeights);
    expect(workerSheet.rowHeights[4]).toBeUndefined(); // was 0 (hidden) pre-override
  });

  it('re-collapsing sends rows back as 0-height overrides', async () => {
    const { priv, renderViewportToBitmap, completeRender } = buildWorker();
    const expand = priv.outlineGutter.rowOutline?.groups.find((g) => g.level === 3);
    priv.outlineGutter.applyGroupToggle(expand as OutlineLayout['groups'][number], 'row');
    // The layout was rebuilt after the expand — fetch the group's new object.
    const collapse = priv.outlineGutter.rowOutline?.groups.find((g) => g.level === 3);
    expect(collapse?.collapsed).toBe(false);
    priv.outlineGutter.applyGroupToggle(collapse as OutlineLayout['groups'][number], 'row');
    completeRender({ close: vi.fn() } as unknown as ImageBitmap);
    await settleRenders();

    const o = lastOverrides(renderViewportToBitmap);
    expect(o?.rows).toMatchObject({ 4: 0, 5: 0, 6: 0, 7: 0 });

    const workerSheet = outlineWorksheet();
    applySizeOverrides(workerSheet, o);
    expect(workerSheet.rowHeights).toEqual(priv.currentWorksheet.rowHeights);
  });

  it('anchors the surviving summary row after collapsing its detail group', () => {
    const { priv } = buildWorker();
    priv.canvasArea.clientHeight = 60;
    priv.scrollHost.clientWidth = 800;
    priv.scrollHost.clientHeight = 60;
    priv.scrollHost.scrollWidth = 1000;
    priv.scrollHost.scrollHeight = 1000;
    const expand = priv.outlineGutter.rowOutline?.groups.find((g) => g.level === 3);
    priv.outlineGutter.applyGroupToggle(expand as OutlineLayout['groups'][number], 'row');
    const collapse = priv.outlineGutter.rowOutline?.groups.find((g) => g.level === 3);

    priv.outlineGutter.applyGroupToggle(collapse as OutlineLayout['groups'][number], 'row');

    // Rows 1-3 remain at the 15pt default (= 20 CSS px) and rows 4-7 collapse
    // to zero, so summary row 8 starts exactly 60px into the scrollable axis.
    expect(priv.scrollHost.scrollTop).toBe(60);
  });
});

describe('worker-mode drag-to-resize reaches the grid bitmap (#567 hole)', () => {
  it('reports a failed worker frame after a complete batch, and accepts the next gesture', async () => {
    const onError = vi.fn();
    const { v, priv, renderViewportToBitmap, completeRender } = buildWorker(onError);
    v.setSelection('B:D');
    const selection = v.selectionState;
    const drag = (pixels: number) => {
      const cell = v.getCellViewportRect('B1')!;
      const event = (x: number) => ({ button: 0, pointerId: 1, pointerType: 'mouse',
        clientX: x, clientY: 5, preventDefault() {}, shiftKey: false, ctrlKey: false, metaKey: false });
      priv.scrollHost.dispatch('pointerdown', event(cell.x + cell.width + 0.5));
      priv.scrollHost.dispatch('pointermove', event(cell.x + pixels));
      priv.scrollHost.dispatch('pointercancel', event(cell.x + pixels));
    };
    renderViewportToBitmap.mockImplementation(() => Promise.reject(new Error('synthetic worker frame failure')));
    drag(84);
    completeRender({ close: vi.fn() } as unknown as ImageBitmap);
    await settleRenders();
    expect(onError).toHaveBeenCalledExactlyOnceWith(expect.objectContaining({ message: 'synthetic worker frame failure' }));
    expect(lastOverrides(renderViewportToBitmap)?.columnCssWidths).toEqual({ 2: 84, 3: 84, 4: 84 });
    expect(priv.currentWorksheet.colWidths[4]).toBe(priv.currentWorksheet.colWidths[2]);
    expect(v.selectionState).toEqual(selection);
    renderViewportToBitmap.mockImplementation(() => Promise.resolve({ close: vi.fn() } as unknown as ImageBitmap));
    drag(90);await settleRenders();
    expect(lastOverrides(renderViewportToBitmap)?.columnCssWidths).toEqual({ 2: 90, 3: 90, 4: 90 });
    expect(onError).toHaveBeenCalledTimes(1);
    v.destroy();
  });

  it.each(['col', 'row'] as const)('transfers every selected %s resize through the actual override channel', async (axis) => {
    const { v, priv, renderViewportToBitmap, completeRender } = buildWorker();
    v.setSelection(axis === 'col' ? 'B:D' : '2:9');
    const cell = v.getCellViewportRect(axis === 'col' ? 'B1' : 'A2');
    if (!cell) throw new Error('Expected visible header geometry');
    const event = (clientX: number, clientY: number) => ({ button: 0, pointerId: 1, pointerType: 'mouse',
      clientX, clientY, preventDefault() {}, shiftKey: false, ctrlKey: false, metaKey: false });
    priv.scrollHost.dispatch('pointerdown', event(axis === 'col' ? cell.x + cell.width + 0.5 : 5,
      axis === 'row' ? cell.y + cell.height + 0.5 : 5));
    const end = event(axis === 'col' ? cell.x + 84 : 5, axis === 'row' ? cell.y + 84 : 5);
    priv.scrollHost.dispatch('pointermove', end);
    priv.scrollHost.dispatch('pointerup', end);
    completeRender({ close: vi.fn() } as unknown as ImageBitmap);
    await settleRenders();
    const workerSheet = outlineWorksheet();
    applySizeOverrides(workerSheet, structuredClone(lastOverrides(renderViewportToBitmap)));
    const actual = axis === 'col' ? workerSheet.colWidths : workerSheet.rowHeights;
    const viewerSizes = axis === 'col' ? priv.currentWorksheet.colWidths : priv.currentWorksheet.rowHeights;
    expect(actual).toEqual(viewerSizes);
    const indices = axis === 'col' ? [2, 3, 4] : [2, 3, 8, 9];
    for (const index of indices) expect(actual[index]).toBe(actual[2]);
    if (axis === 'row') expect([4, 5, 6, 7].map((index) => actual[index])).toEqual([0, 0, 0, 0]);
    v.destroy();
  });

  it('a column resize sends the new width as a col override', () => {
    const { priv, renderViewportToBitmap } = buildWorker();
    priv.selectionInput.resizeDrag = { kind: 'col', index: 2, originScaled: 0, mdw: 7, indices: [2], worksheet: priv.currentWorksheet };
    priv.selectionInput.applyResize(100, 0); // drag column B's right border to x=100

    const o = lastOverrides(renderViewportToBitmap);
    const newWidth = priv.currentWorksheet.colWidths[2];
    expect(newWidth).toBeGreaterThan(0);
    expect(o?.cols?.[2]).toBe(newWidth);

    const workerSheet = outlineWorksheet();
    applySizeOverrides(workerSheet, o);
    expect(workerSheet.colWidths[2]).toBe(newWidth);
  });

  it('a row resize sends the new height as a row override', () => {
    const { priv, renderViewportToBitmap } = buildWorker();
    priv.selectionInput.resizeDrag = { kind: 'row', index: 9, originScaled: 0, mdw: 7, indices: [9], worksheet: priv.currentWorksheet };
    priv.selectionInput.applyResize(0, 60);

    const o = lastOverrides(renderViewportToBitmap);
    const newHeight = priv.currentWorksheet.rowHeights[9];
    expect(newHeight).toBeGreaterThan(0);
    expect(o?.rows?.[9]).toBe(newHeight);
  });
});

describe('applySizeOverrides wire semantics', () => {
  it('sets numeric values and deletes null entries on both axes', () => {
    const ws = outlineWorksheet();
    ws.colWidths[3] = 20;
    applySizeOverrides(ws, { rows: { 4: null, 9: 30 }, cols: { 3: null, 5: 12 } });
    expect(ws.rowHeights[4]).toBeUndefined();
    expect(ws.rowHeights[9]).toBe(30);
    expect(ws.colWidths[3]).toBeUndefined();
    expect(ws.colWidths[5]).toBe(12);
  });

  it('is idempotent (re-applying the same map is a no-op) and undefined-safe', () => {
    const ws = outlineWorksheet();
    const o: WireSizeOverrides = { rows: { 4: null, 5: 22 } };
    applySizeOverrides(ws, o);
    const snapshot = JSON.stringify(ws.rowHeights);
    applySizeOverrides(ws, o);
    expect(JSON.stringify(ws.rowHeights)).toBe(snapshot);
    applySizeOverrides(ws, undefined); // absent overrides: no mutation
    expect(JSON.stringify(ws.rowHeights)).toBe(snapshot);
  });

  it('creates an isolated render projection for each viewer override set', () => {
    const cached = outlineWorksheet();
    const first = createSizeOverriddenWorksheet(cached, { rows: { 4: null }, cols: { 2: 12 } });
    const second = createSizeOverriddenWorksheet(cached, { rows: { 4: 30 }, cols: { 2: 18 } });

    expect(first.rowHeights[4]).toBeUndefined();
    expect(first.colWidths[2]).toBe(12);
    expect(second.rowHeights[4]).toBe(30);
    expect(second.colWidths[2]).toBe(18);
    expect(cached.rowHeights[4]).toBe(0);
    expect(cached.colWidths[2]).toBeUndefined();
  });

  it('reuses viewport-independent render indexes across size projections', () => {
    const cached = outlineWorksheet();
    const first = createSizeOverriddenWorksheet(cached, { rows: { 4: null } });
    const second = createSizeOverriddenWorksheet(cached, { cols: { 2: 18 } });

    inheritSheetRenderCache(cached, first);
    inheritSheetRenderCache(cached, second);

    expect(getSheetRenderCache(first)).toBe(getSheetRenderCache(cached));
    expect(getSheetRenderCache(second)).toBe(getSheetRenderCache(cached));
  });

  it('reuses one worksheet projection across viewport frames until its revision changes', () => {
    const source = outlineWorksheet();
    const cache = new WorksheetViewProjectionCache();
    const first = cache.resolve(source, 0, { id: 7, revision: 1 }, { cols: { 2: 18 } });
    const nextFrame = cache.resolve(source, 0, { id: 7, revision: 1 }, { cols: { 2: 18 } });
    expect(first.created).toBe(true);
    expect(nextFrame.created).toBe(false);
    expect(nextFrame.worksheet).toBe(first.worksheet);

    const resized = cache.resolve(source, 0, { id: 7, revision: 2 }, { cols: { 2: 21 } });
    expect(resized.created).toBe(true);
    expect(resized.worksheet).not.toBe(first.worksheet);
    expect(resized.worksheet.colWidths[2]).toBe(21);

    cache.release(7);
    const reopened = cache.resolve(source, 0, { id: 7, revision: 2 }, { cols: { 2: 21 } });
    expect(reopened.created).toBe(true);
    expect(reopened.worksheet).not.toBe(resized.worksheet);
    const lateFrame = cache.resolve(source, 0, { id: 7, revision: 2 }, { cols: { 2: 21 } });
    expect(lateFrame.worksheet).not.toBe(reopened.worksheet);

    cache.clear();
    const nextWorkbook = cache.resolve(source, 0, { id: 7, revision: 2 }, { cols: { 2: 21 } });
    const nextWorkbookFrame = cache.resolve(source, 0, { id: 7, revision: 2 }, { cols: { 2: 21 } });
    expect(nextWorkbookFrame.worksheet).toBe(nextWorkbook.worksheet);
  });
});

describe('override plumbing stays silent when nothing was mutated', () => {
  it('a plain render request carries NO sizeOverrides key', async () => {
    const { priv, renderViewportToBitmap } = buildWorker();
    void priv.renderCurrentSheet();
    expect(renderViewportToBitmap).toHaveBeenCalledTimes(1);
    const call = renderViewportToBitmap.mock.calls[0] as unknown[] | undefined;
    const opts = call?.[2] as Record<string, unknown> | undefined;
    expect(opts).toBeDefined();
    expect(opts !== undefined && 'sizeOverrides' in opts).toBe(false);
  });
});

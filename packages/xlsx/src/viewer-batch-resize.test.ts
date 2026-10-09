import { afterEach, describe, expect, it, vi } from 'vitest';
import { XlsxSheetViewer } from './viewer.js';
import { colWidthToPx, rowHeightToPx, bindXlsxOfficeFontRoutes, getGridGeometryForWorksheet } from './renderer.js';
import { GridGeometry } from './internal/grid-geometry.js';
import { WorksheetViewProjectionCache, type WireSizeOverrides } from './worker-protocol.js';
import type { Worksheet } from './types.js';
import type { XlsxSelectionInput } from './selection.js';
import { installDom, makeEl, type FakeEl } from './viewer-destroy-test-dom.js';

afterEach(() => { vi.unstubAllGlobals(); vi.restoreAllMocks(); });

function fixture(selection: XlsxSelectionInput, rtl = false) {
  installDom();
  const onError = vi.fn();
  const viewer = new XlsxSheetViewer(makeEl('canvas') as unknown as HTMLCanvasElement, { cellScale: 1.25, onError });
  const engine = (viewer as unknown as { engine: {
    currentWorksheet: Worksheet; canvasArea: FakeEl; scrollHost: FakeEl;
    viewEdits: { wireSizeOverrides(index: number): { overrides: WireSizeOverrides; revision: number } | undefined };
    _cellRect(row: number, col: number): { x: number; y: number; w: number; h: number };
    updateSelectionOverlay(): void;
  } }).engine;
  engine.currentWorksheet = {
    name: 'Synthetic resize', rows: [], colWidths: { 1: 10, 2: 11, 3: 0, 4: 13 },
    rowHeights: { 1: 15, 2: 21, 3: 0, 4: 27 }, defaultColWidth: 8, defaultRowHeight: 15,
    mergeCells: [], freezeRows: 1, freezeCols: 1, rightToLeft: rtl,
    conditionalFormats: [], charts: [], images: [], shapeGroups: [],
  } as unknown as Worksheet;
  engine.canvasArea.clientWidth = engine.scrollHost.clientWidth = 900;
  engine.canvasArea.clientHeight = engine.scrollHost.clientHeight = 600;
  viewer.setSelection(selection);
  const drag = (axis: 'col' | 'row', index: number, pixels: number, during?: () => void, cancel = false,
    afterMove?: () => void) => {
    const cell = engine._cellRect(axis === 'row' ? index : 1, axis === 'col' ? index : 1);
    const logicalX = cell.x + cell.w + 0.5;
    const x = rtl ? 900 - logicalX : logicalX;
    const event = (clientX: number, clientY: number) => ({ button: 0, pointerId: 1, pointerType: 'mouse',
      clientX, clientY, preventDefault() {}, shiftKey: false, ctrlKey: false, metaKey: false });
    engine.scrollHost.dispatch('pointerdown', event(axis === 'col' ? x : (rtl ? 895 : 5), axis === 'row' ? cell.y + cell.h + 0.5 : 5));
    during?.();
    const endX = cell.x + pixels * 1.25;
    const end = event(axis === 'col' ? (rtl ? 900 - endX : endX) : (rtl ? 895 : 5), axis === 'row' ? cell.y + pixels * 1.25 : 5);
    engine.scrollHost.dispatch('pointermove', end);
    afterMove?.();
    engine.scrollHost.dispatch(cancel ? 'pointercancel' : 'pointerup', end);
  };
  return { viewer, engine, drag, onError };
}

describe('selected whole-band drag resizing', () => {
  it.each(['col', 'row'] as const)('resizes selected %ss identically, preserving hidden bands, selection and frozen siblings', (axis) => {
    const { viewer, engine, drag } = fixture(axis === 'col' ? 'B:D' : '2:4', axis === 'col');
    const selection = viewer.selectionState;
    drag(axis, 2, 84);
    const ws = engine.currentWorksheet;
    const sizes = axis === 'col' ? ws.colWidths : ws.rowHeights;
    const px = axis === 'col' ? (size: number) => colWidthToPx(size, 8) : rowHeightToPx;
    expect(px(sizes[2])).toBe(84);
    expect(sizes[4]).toBe(sizes[2]);
    expect(sizes[3]).toBe(0);
    expect(sizes[1]).toBe(axis === 'col' ? 10 : 15);
    expect(viewer.selectionState).toEqual(selection);
    viewer.destroy();
  });

  it('unions discontiguous same-axis selections and snapshots the targets before later API changes', () => {
    const { viewer, engine, drag } = fixture({ areas: [
      { kind: 'columns', firstColumn: 2, lastColumn: 4 },
      { kind: 'columns', firstColumn: 6, lastColumn: 7 },
      { kind: 'cells', top: 1, bottom: 2, left: 8, right: 9 },
    ], activeAreaIndex: 0, activeCell: { row: 1, col: 2 }, extensionAnchor: { row: 1, col: 2 } });
    drag('col', 2, 84, () => viewer.setSelection('J:L'));
    const widths = engine.currentWorksheet.colWidths;
    expect([2, 4, 6, 7].map((index) => colWidthToPx(widths[index], 8))).toEqual([84, 84, 84, 84]);
    expect(widths[3]).toBe(0);
    expect([widths[5], widths[8], widths[9], widths[10]]).toEqual([undefined, undefined, undefined, undefined]);
    viewer.destroy();
  });

  it('preserves a hidden column encoded as a width range rather than an explicit size', () => {
    const { viewer, engine, drag } = fixture('B:D');
    delete engine.currentWorksheet.colWidths[3];
    engine.currentWorksheet.colWidthRanges = [{ min: 3, max: 3, width: 0 }];
    GridGeometry.invalidate(engine.currentWorksheet);
    drag('col', 2, 84);
    expect(getGridGeometryForWorksheet(engine.currentWorksheet).col.sizeOf(3)).toBe(0);
    expect(engine.currentWorksheet.colWidths[3]).toBeUndefined();
    expect(engine.viewEdits.wireSizeOverrides(0)!.overrides.columnCssWidths).toEqual({ 2: 84, 4: 84 });
    viewer.destroy();
  });

  it.each(['B:D', 'B2:D4'])('keeps an outside boundary or a cell selection single (%s)', (selection) => {
    const { viewer, engine, drag } = fixture(selection);
    const index = selection === 'B:D' ? 5 : 4;
    drag('col', index, 84);
    expect(colWidthToPx(engine.currentWorksheet.colWidths[index], 8)).toBe(84);
    expect(engine.currentWorksheet.colWidths[2]).toBe(11);
    if (index === 5) expect(engine.currentWorksheet.colWidths[4]).toBe(13);
    viewer.destroy();
  });

  it('uses the existing minimum for every selected row', () => {
    const { viewer, engine, drag } = fixture('2:4');
    drag('row', 2, -100);
    expect(rowHeightToPx(engine.currentWorksheet.rowHeights[2])).toBe(5);
    expect(engine.currentWorksheet.rowHeights[4]).toBe(engine.currentWorksheet.rowHeights[2]);
    expect(engine.currentWorksheet.rowHeights[3]).toBe(0);
    viewer.destroy();
  });

  it('resizes a million rows with a compact wire, preserving hidden rows and the authored sparse map', () => {
    const { viewer, engine, drag, onError } = fixture('1:1048576');
    const before = { ...engine.currentWorksheet.rowHeights };
    drag('row', 2, 84);
    expect(onError).not.toHaveBeenCalled();
    const row = getGridGeometryForWorksheet(engine.currentWorksheet).row;
    expect([1, 2, 1048576].map(index => row.sizeOf(index))).toEqual([84, 84, 84]);
    expect(row.sizeOf(3)).toBe(0);
    expect(row.offsetOf(1048577)).toBe(1048575 * 84);
    expect(engine._cellRect(1048576, 1).h).toBe(105);
    expect(engine.currentWorksheet.rowHeights).toEqual(before);
    const wire = structuredClone(engine.viewEdits.wireSizeOverrides(0)!.overrides);
    expect(JSON.stringify(wire).length).toBeLessThan(1024);
    const source = { ...engine.currentWorksheet, rowHeights: before };
    const cache = new WorksheetViewProjectionCache();
    const worker = cache.resolve(source, 0, { id: 1, revision: 1 }, wire).worksheet;
    expect(getGridGeometryForWorksheet(worker).row.sizeOf(1048576)).toBe(84);
    expect(getGridGeometryForWorksheet(worker).row.sizeOf(3)).toBe(0);
    viewer.destroy();
  });

  it('rolls a cancelled huge-row preview back to the complete previous interval projection', () => {
    const { viewer, engine, drag, onError } = fixture('1:1048576');
    drag('row', 2, 84);
    const beforeWire = structuredClone(engine.viewEdits.wireSizeOverrides(0)!.overrides);
    const beforeRows = { ...engine.currentWorksheet.rowHeights };
    drag('row', 2, 120, undefined, true);
    expect(onError).not.toHaveBeenCalled();
    expect(engine.currentWorksheet.rowHeights).toEqual(beforeRows);
    expect(engine.viewEdits.wireSizeOverrides(0)!.overrides).toEqual(beforeWire);
    expect(getGridGeometryForWorksheet(engine.currentWorksheet).row.sizeOf(1048576)).toBe(84);
    drag('row', 2, 90);
    expect(getGridGeometryForWorksheet(engine.currentWorksheet).row.sizeOf(1048576)).toBe(90);
    viewer.destroy();
  });

  it('keeps blank zero-default rows hidden while resizing only sparse positive rows', () => {
    const { viewer, engine, drag, onError } = fixture('1:1048576');
    engine.currentWorksheet.defaultRowHeight = 0;
    GridGeometry.invalidate(engine.currentWorksheet);
    drag('row', 2, 84);
    expect(onError).not.toHaveBeenCalled();
    const row = getGridGeometryForWorksheet(engine.currentWorksheet).row;
    expect([1, 2, 4].map(index => row.sizeOf(index))).toEqual([84, 84, 84]);
    expect([3, 5, 1048576].map(index => row.sizeOf(index))).toEqual([0, 0, 0]);
    expect(row.offsetOf(1048577)).toBe(252);
    viewer.destroy();
  });

  it('overlays discontiguous row ranges without altering earlier outside dimensions', () => {
    const { viewer, engine, drag } = fixture('1:1048576');
    drag('row', 2, 84);
    viewer.setSelection({ areas: [{ kind: 'rows', firstRow: 1, lastRow: 2 },
      { kind: 'rows', firstRow: 100000, lastRow: 200000 }], activeAreaIndex: 0,
      activeCell: { row: 2, col: 1 }, extensionAnchor: { row: 2, col: 1 } });
    drag('row', 2, 90);
    const row = getGridGeometryForWorksheet(engine.currentWorksheet).row;
    expect([1, 2, 100000, 200000].map(index => row.sizeOf(index))).toEqual([90, 90, 90, 90]);
    expect([4, 99999, 200001, 1048576].map(index => row.sizeOf(index))).toEqual([84, 84, 84, 84]);
    expect(row.sizeOf(3)).toBe(0);
    viewer.destroy();
  });

  it.each(['lostpointercapture', 'Escape', 'destroy'])('restores compact preview dimensions on %s', action => {
    const { viewer, engine, drag } = fixture('1:1048576');
    const ws = engine.currentWorksheet;
    const before = getGridGeometryForWorksheet(ws).row.offsetOf(1048577);
    drag('row', 2, 84, undefined, false, () => {
      expect(getGridGeometryForWorksheet(ws).row.sizeOf(1048576)).toBe(84);
      if (action === 'destroy') viewer.destroy();
      else if (action === 'Escape') engine.scrollHost.dispatch('keydown', { key: 'Escape', preventDefault() {} });
      else engine.scrollHost.dispatch(action, { pointerId: 1 });
    });
    expect(getGridGeometryForWorksheet(ws).row.offsetOf(1048577)).toBe(before);
    viewer.destroy();
  });

  it('rolls back a complete preview when a later UI update fails, then permits a fresh gesture', () => {
    const { viewer, engine, drag, onError } = fixture('1:1048576');
    const before = getGridGeometryForWorksheet(engine.currentWorksheet).row.offsetOf(1048577);
    const original = engine.updateSelectionOverlay.bind(engine);
    engine.updateSelectionOverlay = vi.fn().mockImplementationOnce(() => { throw new Error('overlay unavailable'); })
      .mockImplementation(original);
    drag('row', 2, 84);
    expect(onError).toHaveBeenCalledExactlyOnceWith(expect.objectContaining({ message: 'overlay unavailable' }));
    expect(getGridGeometryForWorksheet(engine.currentWorksheet).row.offsetOf(1048577)).toBe(before);
    expect(engine.viewEdits.wireSizeOverrides(0)).toBeUndefined();
    drag('row', 2, 90);
    expect(getGridGeometryForWorksheet(engine.currentWorksheet).row.sizeOf(1048576)).toBe(90);
    viewer.destroy();
  });

  it('retains the last complete batch on cancellation and releases ownership even if release fails', () => {
    const { viewer, engine, drag, onError } = fixture('B:D');
    engine.scrollHost.releasePointerCapture = () => { throw new Error('capture already lost'); };
    drag('col', 2, 84, undefined, true);
    const before = { ...engine.currentWorksheet.colWidths };
    engine.scrollHost.dispatch('pointermove', { pointerId: 1, pointerType: 'mouse', clientX: 800, clientY: 5, preventDefault() {} });
    expect(engine.currentWorksheet.colWidths).toEqual(before);
    expect(engine.currentWorksheet.colWidths[4]).toBe(engine.currentWorksheet.colWidths[2]);
    expect(onError).toHaveBeenCalledExactlyOnceWith(expect.objectContaining({ message: 'capture already lost' }));
    // A fresh gesture can still acquire ownership after the cancellation.
    engine.scrollHost.releasePointerCapture = () => undefined;
    drag('col', 2, 90);
    expect(colWidthToPx(engine.currentWorksheet.colWidths[4], 8)).toBe(90);
    viewer.destroy();
  });

  it('makes a failed capture atomic and ignores subsequent moves', () => {
    const { viewer, engine, drag, onError } = fixture('2:4');
    const before = { ...engine.currentWorksheet.rowHeights };
    engine.scrollHost.setPointerCapture = () => { throw new Error('no active pointer'); };
    drag('row', 2, 84);
    expect(onError).toHaveBeenCalledExactlyOnceWith(expect.objectContaining({ message: 'no active pointer' }));
    expect(engine.currentWorksheet.rowHeights).toEqual(before);
    expect(engine.viewEdits.wireSizeOverrides(0)).toBeUndefined();
    viewer.destroy();
  });

  it('ignores another pointer trying to steal or cancel an in-flight batch', () => {
    const { viewer, engine, drag } = fixture('B:D');
    drag('col', 2, 84, () => {
      const cell = engine._cellRect(1, 4);
      engine.scrollHost.dispatch('pointerdown', { button: 0, pointerId: 2, pointerType: 'mouse',
        clientX: cell.x + cell.w + 0.5, clientY: 5, preventDefault() {} });
      engine.scrollHost.dispatch('pointercancel', { pointerId: 2 });
    });
    expect([2, 4].map(index => colWidthToPx(engine.currentWorksheet.colWidths[index], 8))).toEqual([84, 84]);
    viewer.destroy();
  });

  it('never applies a delayed drag to a replacement sheet projection', () => {
    const { viewer, engine, drag, onError } = fixture('B:D');
    const previous = engine.currentWorksheet;
    const before = { ...previous.colWidths };
    drag('col', 2, 84, () => {
      engine.currentWorksheet = { ...previous, colWidths: { ...before }, rowHeights: { ...previous.rowHeights } };
    });
    expect(previous.colWidths).toEqual(before);
    expect(engine.currentWorksheet.colWidths).toEqual(before);
    expect(engine.viewEdits.wireSizeOverrides(0)).toBeUndefined();
    expect(onError).not.toHaveBeenCalled();
    viewer.destroy();
  });

  it('preserves every batched column CSS intent across MDW7 to8 and worker projection replay', () => {
    const { viewer, engine, drag } = fixture('B:D');
    engine.currentWorksheet.colWidths[5] = 12; // same raw value as the future MDW7 drag
    const source = structuredClone(engine.currentWorksheet);
    engine.currentWorksheet.defaultFontFamily = 'Metric Test';
    engine.currentWorksheet.defaultFontSize = 11;
    let digit = 7;
    const ctx = { canvas: {}, font: '', save() {}, restore() {}, measureText: () => ({ width: digit }) } as unknown as CanvasRenderingContext2D;
    bindXlsxOfficeFontRoutes(ctx, engine.currentWorksheet);
    drag('col', 2, 84);
    digit = 8;bindXlsxOfficeFontRoutes(ctx, engine.currentWorksheet);
    const geometry = getGridGeometryForWorksheet(engine.currentWorksheet);
    expect([2, 4].map(index => geometry.col.sizeOf(index))).toEqual([84, 84]);
    expect(geometry.col.sizeOf(1)).toBe(80); // untouched authored width10 at MDW8
    expect(geometry.col.sizeOf(3)).toBe(0);
    expect(geometry.col.sizeOf(5)).toBe(96); // untouched authored12 follows MDW8
    const wire = engine.viewEdits.wireSizeOverrides(0)!;
    expect(wire.overrides.columnCssWidths).toEqual({ 2: 84, 4: 84 });
    const replayed = new WorksheetViewProjectionCache().resolve(source, 0, { id: 9, revision: wire.revision }, structuredClone(wire.overrides)).worksheet;
    const workerGeometry = GridGeometry.forWorksheet(replayed, 8);
    expect([1, 2, 3, 4, 5].map(index => workerGeometry.col.sizeOf(index))).toEqual([80, 84, 0, 84, 96]);
    viewer.destroy();
  });

  it('allows all worksheet columns with a bounded complete CSS/raw wire payload', () => {
    const { viewer, engine, drag } = fixture('A:XFD');
    drag('col', 2, 84);
    const wire = engine.viewEdits.wireSizeOverrides(0)!;
    expect(Object.keys(wire.overrides.cols!)).toHaveLength(16_383); // one hidden band remains untouched
    expect(Object.keys(wire.overrides.columnCssWidths!)).toHaveLength(16_383);
    expect(wire.overrides.columnCssWidths![16_384]).toBe(84);
    expect(JSON.stringify(wire.overrides).length).toBeLessThan(1_048_576);
    viewer.destroy();
  });

  it('continues cumulative row edits through compact intervals and permits later overlapping edits', () => {
    const { viewer, engine, drag, onError } = fixture('1:16384');
    drag('row', 2, 84);
    viewer.setSelection('16385:16385');drag('row', 16_385, 84);
    const before = engine.viewEdits.wireSizeOverrides(0)!;
    expect(Object.keys(before.overrides.rows!)).toHaveLength(16_384);
    viewer.setSelection('16386:16386');drag('row', 16_386, 90);
    expect(onError).not.toHaveBeenCalled();
    expect(engine.currentWorksheet.rowHeights[16_386]).toBeUndefined();
    expect(getGridGeometryForWorksheet(engine.currentWorksheet).row.sizeOf(16_386)).toBe(90);
    expect(Object.keys(engine.viewEdits.wireSizeOverrides(0)!.overrides.rows!)).toHaveLength(16_384);
    viewer.setSelection('2:4');drag('row', 2, 90);
    expect(getGridGeometryForWorksheet(engine.currentWorksheet).row.sizeOf(4)).toBe(90);
    expect(onError).not.toHaveBeenCalled();
    viewer.destroy();
  });
});

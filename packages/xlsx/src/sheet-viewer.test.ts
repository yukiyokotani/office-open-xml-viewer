import { afterEach, describe, expect, it, vi } from 'vitest';
import { XlsxSheetViewer } from './viewer.js';
import { XlsxWorkbook, loadXlsxSheetSource } from './workbook.js';
import type { OoxmlResourceMetrics } from '@silurus/ooxml-core';
import type { Worksheet, XlsxComment } from './types.js';
import {
  installDom,
  makeContainer,
  makeDocument,
  makeEl,
  type FakeEl,
} from './viewer-destroy-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

function descendants(root: FakeEl): FakeEl[] {
  return root.children.flatMap((child) => [child, ...descendants(child)]);
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

function worksheetWithChart(name: string): Worksheet {
  return {
    ...worksheet(name),
    charts: [{
      fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
      toCol: 4, toColOff: 0, toRow: 8, toRowOff: 0,
      chart: {
        chartType: 'bar', title: 'Revenue', categories: ['Q1', 'Q2'],
        series: [{ name: 'Actual', values: [10, 20] }],
      },
    }],
  } as unknown as Worksheet;
}

function pointerEvent(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    button: 0,
    pointerId: 1,
    pointerType: 'mouse',
    clientX: 100,
    clientY: 80,
    shiftKey: false,
    ctrlKey: false,
    metaKey: false,
    preventDefault: () => undefined,
    ...overrides,
  };
}

describe('XlsxSheetViewer canvas mount', () => {
  it('forwards delimited source options through the ordinary reload lifecycle', async () => {
    installDom();
    const firstDestroy = vi.fn();
    const secondDestroy = vi.fn();
    const first = {
      mode: 'main',
      sheetNames: ['CSV'],
      tabColors: [null],
      destroy: firstDestroy,
    } as unknown as XlsxWorkbook;
    const second = {
      mode: 'main',
      sheetNames: ['TSV'],
      tabColors: [null],
      destroy: secondDestroy,
    } as unknown as XlsxWorkbook;
    const load = vi.spyOn(XlsxWorkbook, loadXlsxSheetSource)
      .mockResolvedValueOnce(first)
      .mockResolvedValueOnce(second);
    const canvas = makeEl('canvas');
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const engine = (viewer as unknown as {
      engine: { showSheet(index: number): Promise<void> };
    }).engine;
    vi.spyOn(engine, 'showSheet').mockResolvedValue(undefined);
    const csv = new TextEncoder().encode('a;b').buffer as ArrayBuffer;

    await viewer.load(csv, {
      format: 'delimited-text',
      delimiter: ';',
      sheetName: 'CSV',
    });
    expect(load).toHaveBeenNthCalledWith(
      1,
      csv,
      expect.objectContaining({ mode: 'main' }),
      { format: 'delimited-text', delimiter: ';', sheetName: 'CSV' },
    );
    expect(firstDestroy).not.toHaveBeenCalled();

    await viewer.load('/table.tsv', { format: 'tsv' });
    expect(load).toHaveBeenNthCalledWith(
      2,
      '/table.tsv',
      expect.objectContaining({ mode: 'main' }),
      { format: 'tsv' },
    );
    expect(firstDestroy).toHaveBeenCalledOnce();
    expect(secondDestroy).not.toHaveBeenCalled();

    viewer.destroy();
    expect(secondDestroy).toHaveBeenCalledOnce();
  });

  it('routes Ctrl/Cmd areas through the canonical state, context, copy, and notification APIs', async () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const notifications: unknown[] = [];
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement, {
      onSelectionStateChange(selection) { notifications.push(selection); },
    });
    const engine = (viewer as unknown as { engine: {
      currentWorksheet: Worksheet;
      canvasArea: FakeEl;
      scrollHost: FakeEl;
      getCellAt(clientX: number, clientY: number): { row: number; col: number } | null;
    } }).engine;
    engine.currentWorksheet = worksheet('Multi selection');
    engine.canvasArea.clientWidth = 800;
    engine.canvasArea.clientHeight = 600;
    engine.scrollHost.clientWidth = 800;
    engine.scrollHost.clientHeight = 600;

    const first = engine.getCellAt(100, 80);
    const secondStart = engine.getCellAt(300, 180);
    const secondEnd = engine.getCellAt(430, 260);
    const third = engine.getCellAt(560, 340);
    expect(first).not.toBeNull();
    expect(secondStart).not.toBeNull();
    expect(secondEnd).not.toBeNull();
    expect(third).not.toBeNull();
    if (!first || !secondStart || !secondEnd || !third) throw new Error('Expected cell hits');

    engine.scrollHost.dispatch('pointerdown', pointerEvent({ clientX: 100, clientY: 80 }));
    engine.scrollHost.dispatch('pointerup', pointerEvent({ clientX: 100, clientY: 80 }));

    engine.scrollHost.dispatch('pointerdown', pointerEvent({
      clientX: 300, clientY: 180, ctrlKey: true,
    }));
    engine.scrollHost.dispatch('pointermove', pointerEvent({
      clientX: 430, clientY: 260, ctrlKey: true,
    }));
    engine.scrollHost.dispatch('pointerup', pointerEvent({
      clientX: 430, clientY: 260, ctrlKey: true,
    }));

    engine.scrollHost.dispatch('pointerdown', pointerEvent({
      clientX: 560, clientY: 340, metaKey: true,
    }));
    engine.scrollHost.dispatch('pointerup', pointerEvent({
      clientX: 560, clientY: 340, metaKey: true,
    }));

    const expectedSelection = {
      areas: [
        { kind: 'cells', top: first.row, left: first.col, bottom: first.row, right: first.col },
        {
          kind: 'cells',
          top: Math.min(secondStart.row, secondEnd.row),
          left: Math.min(secondStart.col, secondEnd.col),
          bottom: Math.max(secondStart.row, secondEnd.row),
          right: Math.max(secondStart.col, secondEnd.col),
        },
        { kind: 'cells', top: third.row, left: third.col, bottom: third.row, right: third.col },
      ],
      activeAreaIndex: 2,
      activeCell: third,
      extensionAnchor: third,
    } as const;
    expect(viewer.selectionState).toEqual(expectedSelection);
    expect(notifications.at(-1)).toEqual(expectedSelection);
    const context = viewer.getSelectionContext();
    expect(context?.kind).toBe('range');
    if (!context || context.kind !== 'range') throw new Error('Expected range context');
    expect(context.selection).toEqual(expectedSelection);
    await expect(viewer.copySelection()).resolves.toEqual({ status: 'unsupported-multiple-areas' });
    viewer.destroy();
  });

  it('keeps an existing range on contextmenu inside it and selects an outside target', async () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const received: Array<{ originalEvent: MouseEvent; getContext(): Promise<unknown> }> = [];
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement, {
      onContextMenu(event) { received.push(event); },
    });
    const engine = (viewer as unknown as { engine: {
      currentWorksheet: Worksheet;
      canvasArea: FakeEl;
      scrollHost: FakeEl;
      getCellAt(clientX: number, clientY: number): { row: number; col: number } | null;
    } }).engine;
    engine.currentWorksheet = worksheet('Context menu');
    engine.canvasArea.clientWidth = 800;
    engine.canvasArea.clientHeight = 600;
    engine.scrollHost.clientWidth = 800;
    engine.scrollHost.clientHeight = 600;
    viewer.setSelection({
      areas: [
        { kind: 'cells', top: 1, left: 1, bottom: 5, right: 2 },
        { kind: 'cells', top: 10, left: 4, bottom: 11, right: 5 },
      ],
      activeAreaIndex: 1,
      activeCell: { row: 10, col: 4 },
      extensionAnchor: { row: 11, col: 5 },
    });
    const before = viewer.selectionState;
    const insideEvent = {
      button: 2, clientX: 100, clientY: 80, defaultPrevented: false,
    } as unknown as MouseEvent;

    engine.scrollHost.dispatch('contextmenu', insideEvent);

    expect(received[0].originalEvent).toBe(insideEvent);
    expect(viewer.selectionState).toEqual(before);
    await expect(received[0].getContext()).resolves.toMatchObject({ kind: 'range', selection: before });

    const outsideEvent = {
      button: 2, clientX: 300, clientY: 200, defaultPrevented: false,
    } as unknown as MouseEvent;
    const outsideCell = engine.getCellAt(300, 200);
    engine.scrollHost.dispatch('contextmenu', outsideEvent);

    await expect(received[1].getContext()).resolves.toMatchObject({ kind: 'range' });
    expect(viewer.selectionState).toMatchObject({
      areas: [{ kind: 'cells', top: outsideCell?.row, left: outsideCell?.col,
        bottom: outsideCell?.row, right: outsideCell?.col }],
      activeCell: outsideCell,
    });
    viewer.destroy();
  });

  it('selects and outlines a chart from contextmenu when element selection is enabled', async () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const contexts: Array<Promise<unknown>> = [];
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement, {
      enableElementSelection: true,
      onContextMenu(event) { contexts.push(event.getContext()); },
    });
    const engine = (viewer as unknown as { engine: {
      currentWorksheet: Worksheet;
      canvasArea: FakeEl;
      scrollHost: FakeEl;
      selectionOverlay: FakeEl;
    } }).engine;
    engine.currentWorksheet = worksheetWithChart('Objects');
    engine.canvasArea.clientWidth = 800;
    engine.canvasArea.clientHeight = 600;
    engine.scrollHost.clientWidth = 800;
    engine.scrollHost.clientHeight = 600;

    engine.scrollHost.dispatch('contextmenu', {
      button: 2, clientX: 100, clientY: 80, defaultPrevented: false,
    });

    await expect(contexts[0]).resolves.toMatchObject({
      format: 'xlsx', kind: 'element', elementType: 'chart',
    });
    expect(engine.selectionOverlay.querySelector('[data-xlsx-element-context-outline]')).not.toBeNull();
    viewer.destroy();
  });

  it('preserves a selected row from its header contextmenu and selects an unselected row', async () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const contexts: Array<Promise<unknown>> = [];
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement, {
      onContextMenu(event) { contexts.push(event.getContext()); },
      onError: vi.fn(),
    });
    const engine = (viewer as unknown as { engine: {
      currentWorksheet: Worksheet;
      canvasArea: FakeEl;
      scrollHost: FakeEl;
    } }).engine;
    engine.currentWorksheet = worksheet('Headers');
    engine.canvasArea.clientWidth = 800;
    engine.canvasArea.clientHeight = 600;
    engine.scrollHost.clientWidth = 800;
    engine.scrollHost.clientHeight = 600;
    viewer.setSelection('2:4');

    engine.scrollHost.dispatch('contextmenu', {
      button: 2, clientX: 25, clientY: 80, defaultPrevented: false,
    });
    expect(viewer.selectionState?.areas).toEqual([{ kind: 'rows', firstRow: 2, lastRow: 4 }]);
    await expect(contexts[0]).resolves.toMatchObject({ kind: 'range' });

    engine.scrollHost.dispatch('contextmenu', {
      button: 2, clientX: 25, clientY: 200, defaultPrevented: false,
    });
    const outsideArea = viewer.selectionState?.areas[0];
    expect(outsideArea?.kind).toBe('rows');
    if (outsideArea?.kind === 'rows') {
      expect(outsideArea.firstRow).toBe(outsideArea.lastRow);
      expect(outsideArea.firstRow < 2 || outsideArea.firstRow > 4).toBe(true);
    }
    await expect(contexts[1]).resolves.toMatchObject({ kind: 'range' });
    viewer.destroy();
  });

  it('selects an A1 range as geometry without inventing ActiveCell direction', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);

    viewer.setSelection('D5:B2');

    expect(viewer.selectionState).toEqual({
      areas: [{ kind: 'cells', top: 2, left: 2, bottom: 5, right: 4 }],
      activeAreaIndex: 0,
      activeCell: { row: 2, col: 2 },
      extensionAnchor: { row: 2, col: 2 },
    });
    viewer.destroy();
  });

  it('keeps bounded cell rectangles distinct from row and column selections', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);

    viewer.setSelection('A2:XFD4');
    expect(viewer.selectionState?.areas[0]).toEqual({
      kind: 'cells', top: 2, left: 1, bottom: 4, right: 16_384,
    });
    viewer.setSelection('2:4');
    expect(viewer.selectionState?.areas[0]).toEqual({ kind: 'rows', firstRow: 2, lastRow: 4 });
    viewer.setSelection('B:D');
    expect(viewer.selectionState?.areas[0]).toEqual({
      kind: 'columns', firstColumn: 2, lastColumn: 4,
    });

    viewer.destroy();
  });

  it('represents ActiveCell, extension anchor, and multiple areas independently', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);

    viewer.setSelection({
      areas: [
        { kind: 'cells', top: 1, left: 1, bottom: 4, right: 4 },
        { kind: 'cells', top: 8, left: 2, bottom: 9, right: 3 },
      ],
      activeAreaIndex: 0,
      activeCell: { row: 2, col: 2 },
      extensionAnchor: { row: 4, col: 4 },
    });

    expect(viewer.selectionState?.activeCell).toEqual({ row: 2, col: 2 });
    expect(viewer.selectionState?.extensionAnchor).toEqual({ row: 4, col: 4 });
    expect(viewer.selectionState?.areas).toHaveLength(2);
    viewer.destroy();
  });

  it('emits canonical state only for semantic changes and validates invariants', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const onSelectionStateChange = vi.fn();
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement, {
      onSelectionStateChange,
    });

    viewer.setSelection('B2:D5');
    viewer.setSelection('D5:B2');
    expect(onSelectionStateChange).toHaveBeenCalledOnce();
    expect(() => viewer.setSelection({
      areas: [{ kind: 'cells', top: 1, left: 1, bottom: 2, right: 2 }],
      activeAreaIndex: 0,
      activeCell: { row: 3, col: 1 },
      extensionAnchor: { row: 1, col: 1 },
    })).toThrow(/activeCell/);
    viewer.destroy();
  });

  it('coalesces bounded context notifications to the latest selection per frame', () => {
    const document = installDom();
    const frames: FrameRequestCallback[] = [];
    Object.assign(document.defaultView, {
      requestAnimationFrame(callback: FrameRequestCallback) {
        frames.push(callback);
        return frames.length;
      },
      cancelAnimationFrame: vi.fn(),
    });
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const contexts: unknown[] = [];
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement, {
      onSelectionContextChange(context) {
        contexts.push(context);
      },
    });
    const engine = (viewer as unknown as { engine: { currentWorksheet: Worksheet } }).engine;
    engine.currentWorksheet = {
      ...worksheet('AI context'),
      rows: [
        { index: 1, cells: [{ row: 1, col: 1, value: { type: 'text', text: 'old' } }] },
        { index: 2, cells: [{ row: 2, col: 2, value: { type: 'text', text: 'latest' } }] },
      ],
    } as unknown as Worksheet;

    viewer.setSelection('A1');
    viewer.setSelection('B2');
    expect(contexts).toEqual([]);
    expect(frames).toHaveLength(1);
    frames.shift()?.(0);

    expect(contexts).toHaveLength(1);
    expect(contexts[0]).toMatchObject({
      format: 'xlsx',
      kind: 'range',
      maxCells: 1_000,
      maxTextCharacters: 65_536,
      cells: [{ address: { row: 2, col: 2 }, value: 'latest' }],
    });
    viewer.destroy();
  });

  it('keeps element context disabled when only the context callback is supplied', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement, {
      onSelectionContextChange: vi.fn(),
    });
    const engine = (viewer as unknown as { engine: {
      currentWorksheet: Worksheet;
      canvasArea: FakeEl;
      scrollHost: FakeEl;
      selectionOverlay: FakeEl;
    } }).engine;
    engine.currentWorksheet = worksheetWithChart('Objects');
    engine.canvasArea.clientWidth = 800;
    engine.canvasArea.clientHeight = 600;
    engine.scrollHost.clientWidth = 800;
    engine.scrollHost.clientHeight = 600;
    expect(engine.scrollHost._listeners.get('contextmenu') ?? []).toHaveLength(0);

    engine.scrollHost.dispatch('pointerdown', pointerEvent({ clientX: 100, clientY: 80 }));
    engine.scrollHost.dispatch('pointerup', pointerEvent({ clientX: 100, clientY: 80 }));

    expect(viewer.getSelectionContext()?.kind).toBe('range');
    viewer.destroy();
  });

  it('supports getter-only chart context when explicitly enabled', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement, {
      enableElementSelection: true,
    });
    const engine = (viewer as unknown as { engine: {
      currentWorksheet: Worksheet;
      canvasArea: FakeEl;
      scrollHost: FakeEl;
      selectionOverlay: FakeEl;
    } }).engine;
    engine.currentWorksheet = worksheetWithChart('Objects');
    engine.canvasArea.clientWidth = 800;
    engine.canvasArea.clientHeight = 600;
    engine.scrollHost.clientWidth = 800;
    engine.scrollHost.clientHeight = 600;

    engine.scrollHost.dispatch('pointerdown', pointerEvent({ clientX: 100, clientY: 80 }));
    engine.scrollHost.dispatch('pointerup', pointerEvent({ clientX: 100, clientY: 80 }));

    expect(viewer.getSelectionContext()).toMatchObject({
      format: 'xlsx', kind: 'element', sheetName: 'Objects', elementType: 'chart',
      text: expect.stringContaining('Revenue'),
    });
    expect(engine.selectionOverlay.querySelector('[data-xlsx-element-context-outline]')).not.toBeNull();

    engine.scrollHost.dispatch('pointerdown', pointerEvent({ clientX: 700, clientY: 500 }));
    engine.scrollHost.dispatch('pointerup', pointerEvent({ clientX: 700, clientY: 500 }));
    expect(viewer.getSelectionContext()?.kind).toBe('range');
    expect(engine.selectionOverlay.querySelector('[data-xlsx-element-context-outline]')).toBeNull();
    viewer.destroy();
  });

  it('serializes reentrant selection callbacks in semantic order', async () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const activeCells: Array<{ row: number; col: number }> = [];
    let viewer: XlsxSheetViewer;
    viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement, {
      onSelectionStateChange(selection) {
        if (!selection) return;
        activeCells.push(selection.activeCell);
        if (selection.activeCell.row === 2) viewer.setSelection('C3');
      },
    });

    viewer.setSelection('B2');
    await Promise.resolve();

    expect(activeCells).toEqual([{ row: 2, col: 2 }, { row: 3, col: 3 }]);
    expect(viewer.selectionState?.activeCell).toEqual({ row: 3, col: 3 });
    viewer.destroy();
  });

  it('bounds reentrant callback cycles without synchronously hanging', async () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    let calls = 0;
    let viewer: XlsxSheetViewer;
    viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement, {
      onSelectionStateChange(selection) {
        if (!selection) return;
        calls++;
        viewer.setSelection(selection.activeCell.row === 1 ? 'B2' : 'A1');
      },
    });

    viewer.setSelection('A1');
    await new Promise((resolve) => setTimeout(resolve, 0));

    expect(calls).toBe(100);
    expect(viewer.selectionState?.activeCell).toEqual({ row: 1, col: 1 });
    viewer.destroy();
  });

  it('delivers a repeated state when it is a distinct semantic transition', async () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const activeCells: string[] = [];
    let viewer: XlsxSheetViewer;
    viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement, {
      onSelectionStateChange(selection) {
        if (!selection) return;
        activeCells.push(`${selection.activeCell.row},${selection.activeCell.col}`);
        if (activeCells.length === 1) viewer.setSelection('B2');
        else if (activeCells.length === 2) viewer.setSelection('A1');
        else if (activeCells.length === 3) viewer.setSelection('C3');
      },
    });

    viewer.setSelection('A1');
    await new Promise((resolve) => setTimeout(resolve, 0));

    expect(activeCells).toEqual(['1,1', '2,2', '1,1', '3,3']);
    expect(viewer.selectionState?.activeCell).toEqual({ row: 3, col: 3 });
    viewer.destroy();
  });

  it('delivers a reentrant state even when the initiating callback throws', async () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const activeRows: number[] = [];
    let viewer: XlsxSheetViewer;
    viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement, {
      onSelectionStateChange(selection) {
        if (!selection) return;
        activeRows.push(selection.activeCell.row);
        if (selection.activeCell.row === 1) {
          viewer.setSelection('A2');
          throw new Error('consumer failure');
        }
      },
    });

    expect(() => viewer.setSelection('A1')).toThrow('consumer failure');
    await Promise.resolve();
    expect(activeRows).toEqual([1, 2]);
    viewer.destroy();
  });

  it('projects a logical selection into frozen-pane fragments without fake pane edges', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const engine = (viewer as unknown as { engine: {
      currentWorksheet: Worksheet;
      canvasArea: FakeEl;
      overlayHost: { selection: FakeEl };
    } }).engine;
    engine.currentWorksheet = { ...worksheet('Frozen'), freezeRows: 1, freezeCols: 1 };
    engine.canvasArea.clientWidth = 640;
    engine.canvasArea.clientHeight = 360;

    viewer.setSelection('A1:C3');

    const borders = descendants(engine.overlayHost.selection).filter(
      (element) => element.getAttribute('data-xlsx-selection-border') !== null,
    );
    expect(borders).toHaveLength(1);
    expect(borders[0].getAttribute('d')?.match(/[MHV]/g)).toHaveLength(6);
    viewer.destroy();
  });

  it('bounds overlay work and geometry when legal freeze counts cover the sheet', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const engine = (viewer as unknown as { engine: {
      currentWorksheet: Worksheet;
      canvasArea: FakeEl;
      overlayHost: { selection: FakeEl };
      getCellRect(row: number, col: number): { x: number; y: number; w: number; h: number } | null;
    } }).engine;
    engine.currentWorksheet = {
      ...worksheet('Fully frozen'),
      freezeRows: 1_048_576,
      freezeCols: 16_384,
      rightToLeft: true,
    };
    engine.canvasArea.clientWidth = 640;
    engine.canvasArea.clientHeight = 360;

    viewer.setSelection('A1:XFD1048576');

    const borders = descendants(engine.overlayHost.selection).filter(
      (element) => element.getAttribute('data-xlsx-selection-border') !== null,
    );
    expect(borders).toHaveLength(1);
    expect(borders[0].getAttribute('d')?.length).toBeLessThan(256);
    viewer.destroy();
  });

  it('mirrors one-sided logical selection borders in RTL and suppresses duplicate fragments', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const engine = (viewer as unknown as { engine: {
      currentWorksheet: Worksheet;
      canvasArea: FakeEl;
      overlayHost: { selection: FakeEl };
    } }).engine;
    engine.currentWorksheet = { ...worksheet('RTL'), rightToLeft: true };
    engine.canvasArea.clientWidth = 240;
    engine.canvasArea.clientHeight = 120;
    viewer.setSelection({
      areas: [
        { kind: 'cells', top: 1, left: 1, bottom: 1, right: 16_384 },
        { kind: 'cells', top: 1, left: 1, bottom: 1, right: 16_384 },
      ],
      activeAreaIndex: 0,
      activeCell: { row: 1, col: 1 },
      extensionAnchor: { row: 1, col: 1 },
    });

    const borders = descendants(engine.overlayHost.selection).filter(
      (element) => element.getAttribute('data-xlsx-selection-border') !== null,
    );
    expect(borders).toHaveLength(1);
    expect(borders[0].getAttribute('d')).toContain('M0 ');
    viewer.destroy();
  });

  it('paints every multiple-selection area in the configured color and only outlines the active cell', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement, {
      selectionColor: '#765432',
    });
    const engine = (viewer as unknown as { engine: {
      currentWorksheet: Worksheet;
      canvasArea: FakeEl;
      overlayHost: { selection: FakeEl };
    } }).engine;
    engine.currentWorksheet = worksheet('Overlapping');
    engine.canvasArea.clientWidth = 640;
    engine.canvasArea.clientHeight = 360;

    viewer.setSelection({
      areas: [
        { kind: 'cells', top: 1, left: 1, bottom: 2, right: 2 },
        { kind: 'cells', top: 2, left: 2, bottom: 3, right: 3 },
      ],
      activeAreaIndex: 0,
      activeCell: { row: 1, col: 1 },
      extensionAnchor: { row: 1, col: 1 },
    });

    const overlayDescendants = descendants(engine.overlayHost.selection);
    const fills = overlayDescendants.filter(
      (element) => element.getAttribute('data-xlsx-selection-fill') !== null,
    );
    const borders = overlayDescendants.filter(
      (element) => element.getAttribute('data-xlsx-selection-border') !== null,
    );
    const activeBorders = overlayDescendants.filter(
      (element) => element.getAttribute('data-xlsx-active-cell-border') !== null,
    );
    const activeCutouts = overlayDescendants.filter(
      (element) => element.getAttribute('data-xlsx-active-cell-cutout') !== null,
    );
    const coloredPaths = overlayDescendants.filter(
      (element) => element.getAttribute('fill') ===
        'color-mix(in srgb, #765432 8%, transparent)',
    );
    expect(fills).toHaveLength(1);
    expect(borders).toHaveLength(0);
    expect(activeBorders).toHaveLength(1);
    expect(activeCutouts).toHaveLength(1);
    expect(activeBorders[0].getAttribute('stroke-width')).toBe('1');
    expect(activeBorders[0].getAttribute('stroke')).toBe('#765432');
    expect(coloredPaths).toHaveLength(1);
    viewer.destroy();
  });

  it('does not outline or divide touching areas in a multiple selection', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const engine = (viewer as unknown as { engine: {
      currentWorksheet: Worksheet;
      canvasArea: FakeEl;
      overlayHost: { selection: FakeEl };
    } }).engine;
    engine.currentWorksheet = worksheet('Touching');
    engine.canvasArea.clientWidth = 640;
    engine.canvasArea.clientHeight = 360;

    viewer.setSelection({
      areas: [
        { kind: 'cells', top: 1, left: 1, bottom: 1, right: 1 },
        { kind: 'cells', top: 1, left: 2, bottom: 1, right: 2 },
      ],
      activeAreaIndex: 1,
      activeCell: { row: 1, col: 2 },
      extensionAnchor: { row: 1, col: 2 },
    });

    const overlayDescendants = descendants(engine.overlayHost.selection);
    const border = overlayDescendants.find(
      (element) => element.getAttribute('data-xlsx-selection-border') !== null,
    );
    const activeBorder = overlayDescendants.find(
      (element) => element.getAttribute('data-xlsx-active-cell-border') !== null,
    );
    expect(border).toBeUndefined();
    expect(activeBorder).toBeDefined();
    expect(activeBorder?.getAttribute('stroke-width')).toBe('1');
    viewer.destroy();
  });

  it('does not materialize a programmatic clipboard range above the cell limit', async () => {
    const document = installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const writeText = vi.fn(() => Promise.resolve());
    Object.assign(document.defaultView, { navigator: { clipboard: { writeText } } });
    const engine = (viewer as unknown as {
      engine: {
        currentWorksheet: Worksheet;
        copySelection(): Promise<unknown>;
      };
    }).engine;
    engine.currentWorksheet = {
      ...worksheet('Sparse'),
      rows: [
        { index: 2, cells: [{ row: 2, col: 2, value: { type: 'text', text: 'first' } }] },
        {
          index: 1_048_575,
          cells: [{ row: 1_048_575, col: 16_383, value: { type: 'text', text: 'last' } }],
        },
      ],
    } as unknown as Worksheet;

    viewer.setSelection('B2:XFC1048575');
    const result = await engine.copySelection();

    expect(writeText).not.toHaveBeenCalled();
    expect(result).toEqual({ status: 'too-large', limit: 'cells' });
    viewer.destroy();
  });

  it('applies the clipboard cell limit to pointer-created selections too', async () => {
    const document = installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const writeText = vi.fn(() => Promise.resolve());
    Object.assign(document.defaultView, { navigator: { clipboard: { writeText } } });
    const engine = (viewer as unknown as {
      engine: {
        currentWorksheet: Worksheet;
        selectionController: {
          select(cell: { row: number; col: number }): void;
          extend(cell: { row: number; col: number }): void;
        };
        copySelection(): Promise<unknown>;
      };
    }).engine;
    engine.currentWorksheet = worksheet('Large selection');
    engine.selectionController.select({ row: 1, col: 1 });
    engine.selectionController.extend({ row: 501, col: 500 });

    const result = await engine.copySelection();

    expect(writeText).not.toHaveBeenCalled();
    expect(result).toEqual({ status: 'too-large', limit: 'cells' });
    viewer.destroy();
  });

  it('reports unsupported multi-area copy without flattening selection semantics', async () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const engine = (viewer as unknown as { engine: { currentWorksheet: Worksheet } }).engine;
    engine.currentWorksheet = worksheet('Areas');
    viewer.setSelection({
      areas: [
        { kind: 'cells', top: 1, left: 1, bottom: 1, right: 1 },
        { kind: 'cells', top: 3, left: 3, bottom: 3, right: 3 },
      ],
      activeAreaIndex: 0,
      activeCell: { row: 1, col: 1 },
      extensionAnchor: { row: 1, col: 1 },
    });

    await expect(viewer.copySelection()).resolves.toEqual({ status: 'unsupported-multiple-areas' });
    viewer.destroy();
  });

  it('extracts bounded, serializable selection context without workbook mutation APIs', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const engine = (viewer as unknown as { engine: { currentWorksheet: Worksheet } }).engine;
    engine.currentWorksheet = {
      ...worksheet('Analysis'),
      rows: [
        { index: 1, cells: [
          { row: 1, col: 1, value: { type: 'text', text: 'Revenue' } },
          { row: 1, col: 2, value: { type: 'number', number: 42 }, formula: 'SUM(B2:B4)' },
        ] },
        { index: 2, cells: [
          { row: 2, col: 1, value: { type: 'bool', bool: true } },
        ] },
      ],
    } as unknown as Worksheet;
    viewer.setSelection('A1:B2');

    const context = viewer.getSelectionContext({ maxCells: 2 });

    expect(context).toMatchObject({
      format: 'xlsx',
      kind: 'range',
      sheetName: 'Analysis',
      coordinateCountUpperBound: 4,
      truncated: true,
      maxCells: 2,
    });
    expect(context?.kind).toBe('range');
    if (!context || context.kind !== 'range') throw new Error('Expected range context');
    expect(context.cells).toEqual([
      { address: { row: 1, col: 1 }, displayText: '', valueType: 'text', value: 'Revenue' },
      {
        address: { row: 1, col: 2 }, displayText: '', valueType: 'number', value: 42,
        formula: 'SUM(B2:B4)',
      },
    ]);
    viewer.destroy();
  });

  it('does not let explicit empty cells consume the populated-cell limit', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const engine = (viewer as unknown as { engine: { currentWorksheet: Worksheet } }).engine;
    engine.currentWorksheet = {
      ...worksheet('Sparse'),
      rows: [{ index: 1, cells: [
        { row: 1, col: 1, value: { type: 'empty' } },
        { row: 1, col: 2, value: { type: 'number', number: 42 } },
      ] }],
    } as unknown as Worksheet;
    viewer.setSelection('A1:B1');

    const context = viewer.getSelectionContext({ maxCells: 1 });

    expect(context?.kind).toBe('range');
    if (!context || context.kind !== 'range') throw new Error('Expected range context');
    expect(context.cells).toEqual([
      { address: { row: 1, col: 2 }, displayText: '', valueType: 'number', value: 42 },
    ]);
    expect(context.truncated).toBe(false);
    viewer.destroy();
  });

  it('includes an authored comment when its cell is selected', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const engine = (viewer as unknown as { engine: {
      currentWorksheet: Worksheet;
      sourceCommentMap: Map<string, XlsxComment>;
    } }).engine;
    const comment = {
      kind: 'thread' as const,
      id: 'root',
      cellRef: 'A1',
      author: 'Ada',
      rootText: 'Review this',
      text: 'Review this\nDone',
      replies: [{
        id: 'reply', parentId: 'root', personId: 'person', author: 'Linus', text: 'Done',
      }],
    };
    engine.currentWorksheet = {
      ...worksheet('Comments'),
      comments: [comment],
      rows: [{ index: 1, cells: [{ row: 1, col: 1, value: { type: 'empty' } }] }],
    } as unknown as Worksheet;
    engine.sourceCommentMap = new Map([['1:1', comment]]);
    viewer.setSelection('A1');

    const context = viewer.getSelectionContext();
    expect(context?.kind).toBe('range');
    if (!context || context.kind !== 'range') throw new Error('Expected range context');
    expect(context.cells[0]).toMatchObject({
      address: { row: 1, col: 1 },
      comment: {
        root: { id: 'root', author: 'Ada', text: 'Review this' },
        replies: [{ id: 'reply', author: 'Linus', text: 'Done' }],
      },
    });
    viewer.destroy();
  });

  it('preserves the published 1 MiB default text budget for direct range queries', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const engine = (viewer as unknown as { engine: { currentWorksheet: Worksheet } }).engine;
    const value = 'x'.repeat(40_000);
    engine.currentWorksheet = {
      ...worksheet('Large context'),
      rows: [{ index: 1, cells: [
        { row: 1, col: 1, value: { type: 'text', text: value } },
        { row: 1, col: 2, value: { type: 'text', text: value } },
      ] }],
    } as unknown as Worksheet;
    viewer.setSelection('A1:B1');

    const context = viewer.getSelectionContext();

    expect(context?.kind).toBe('range');
    if (!context || context.kind !== 'range') throw new Error('Expected range context');
    expect(context.maxTextCharacters).toBe(1_048_576);
    expect(context.textCharacters).toBe(80_000);
    expect(context.truncated).toBe(false);
    expect(context.cells).toHaveLength(2);
    viewer.destroy();
  });

  it('skips rows and cells outside a small selected interval', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const engine = (viewer as unknown as { engine: { currentWorksheet: Worksheet } }).engine;
    let rowCellReads = 0;
    const rows = Array.from({ length: 10_000 }, (_, index) => {
      const row = index + 1;
      return {
        index: row,
        get cells() {
          rowCellReads++;
          return [{ row, col: 10_000, value: { type: 'number' as const, number: row } }];
        },
      };
    });
    engine.currentWorksheet = { ...worksheet('Sparse'), rows } as unknown as Worksheet;
    viewer.setSelection('NTP10000');

    const context = viewer.getSelectionContext();
    expect(context?.kind).toBe('range');
    if (!context || context.kind !== 'range') throw new Error('Expected range context');
    expect(context.cells).toMatchObject([
      { address: { row: 10_000, col: 10_000 }, value: 10_000 },
    ]);
    expect(rowCellReads).toBeLessThan(10);
    viewer.destroy();
  });

  it('extracts context from parser-accepted rows and cells in document order', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const engine = (viewer as unknown as { engine: { currentWorksheet: Worksheet } }).engine;
    engine.currentWorksheet = {
      ...worksheet('Document order'),
      rows: [
        { index: 10, cells: [{ row: 10, col: 1, value: { type: 'text', text: 'later' } }] },
        { index: 1, cells: [
          { row: 1, col: 10, value: { type: 'text', text: 'right' } },
          { row: 1, col: 1, value: { type: 'text', text: 'selected' } },
        ] },
      ],
    } as unknown as Worksheet;
    viewer.setSelection('A1');

    const context = viewer.getSelectionContext();
    expect(context?.kind).toBe('range');
    if (!context || context.kind !== 'range') throw new Error('Expected range context');
    expect(context.cells).toMatchObject([
      { address: { row: 1, col: 1 }, value: 'selected' },
    ]);
    viewer.destroy();
  });

  it('bounds cumulative selection-context text and rejects content access after destroy', async () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const engine = (viewer as unknown as { engine: { currentWorksheet: Worksheet } }).engine;
    engine.currentWorksheet = {
      ...worksheet('Sensitive'),
      rows: [{ index: 1, cells: [{
        row: 1, col: 1, value: { type: 'text', text: '0123456789' }, formula: 'ABCDEFGHIJ',
      }] }],
    } as unknown as Worksheet;
    viewer.setSelection('A1');

    const context = viewer.getSelectionContext({ maxTextCharacters: 5 });
    expect(context).toMatchObject({
      truncated: true,
      truncationReasons: ['text'],
      textCharacters: 5,
      maxTextCharacters: 5,
    });

    viewer.destroy();
    expect(() => viewer.getSelectionContext()).toThrow(/destroyed/);
    await expect(viewer.copySelection()).rejects.toThrow(/destroyed/);
  });

  it('does not split a surrogate pair at the selection-context text limit', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const engine = (viewer as unknown as { engine: { currentWorksheet: Worksheet } }).engine;
    engine.currentWorksheet = {
      ...worksheet('Unicode'),
      rows: [{ index: 1, cells: [{
        row: 1, col: 1, value: { type: 'text', text: '\ud83d\ude00x' },
      }] }],
    } as unknown as Worksheet;
    viewer.setSelection('A1');

    const oneCharacter = viewer.getSelectionContext({ maxTextCharacters: 1 });
    const twoCharacters = viewer.getSelectionContext({ maxTextCharacters: 2 });
    expect(oneCharacter?.kind).toBe('range');
    expect(twoCharacters?.kind).toBe('range');
    if (!oneCharacter || oneCharacter.kind !== 'range' ||
      !twoCharacters || twoCharacters.kind !== 'range') {
      throw new Error('Expected range context');
    }
    expect(oneCharacter.cells[0].value).toBe('');
    expect(twoCharacters.cells[0].value).toBe('\ud83d\ude00');
    viewer.destroy();
  });

  it('quotes tabs, newlines, and quotes in copied TSV', async () => {
    const document = installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const writeText = vi.fn(() => Promise.resolve());
    Object.assign(document.defaultView, { navigator: { clipboard: { writeText } } });
    const engine = (viewer as unknown as { engine: { currentWorksheet: Worksheet } }).engine;
    engine.currentWorksheet = {
      ...worksheet('TSV'),
      rows: [{ index: 1, cells: [
        { row: 1, col: 1, value: { type: 'text', text: 'a\tb' } },
        { row: 1, col: 2, value: { type: 'text', text: 'x\n"y"' } },
      ] }],
    } as unknown as Worksheet;
    viewer.setSelection('A1:B1');

    await expect(viewer.copySelection()).resolves.toMatchObject({ status: 'copied' });
    expect(writeText).toHaveBeenCalledWith('"a\tb"\t"x\n""y"""');
    viewer.destroy();
  });

  it('creates owner-window chrome and scopes copy shortcuts to its viewport', async () => {
    const openerDocument = installDom();
    const popupDocument = makeDocument(2);
    const parent = makeContainer(800, 600, popupDocument);
    const canvas = makeEl('canvas', popupDocument);
    parent.appendChild(canvas);

    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const mounted = descendants(parent);

    expect(mounted.every((element) => element.ownerDocument === popupDocument)).toBe(true);
    const viewerStyle = popupDocument.head.querySelector('style[data-xlsx-viewer-styles]');
    expect(viewerStyle).not.toBeNull();
    expect(viewerStyle?.textContent).toContain(
      '[data-xlsx-viewport-input]:focus{outline:none}',
    );
    expect(viewerStyle?.textContent).toContain(
      '[data-xlsx-viewport-input]:focus-visible{outline:2px solid var(--ooxml-xlsx-focus-ring,transparent);outline-offset:-2px}',
    );
    expect(openerDocument.head.querySelector('style[data-xlsx-viewer-styles]')).toBeNull();
    const viewportInput = mounted.find((element) => element.hasAttribute('data-xlsx-viewport-input')) as FakeEl;
    viewportInput.dispatch('pointerdown', {
      button: 0,
      pointerId: 1,
      pointerType: 'mouse',
      clientX: 0,
      clientY: 0,
      shiftKey: false,
    });
    expect(viewportInput._listeners.get('keydown')).toHaveLength(1);
    expect(popupDocument.listenerCount('keydown')).toBe(0);
    expect(openerDocument.listenerCount('keydown')).toBe(0);

    const popupWrite = vi.fn(() => Promise.resolve());
    const openerWrite = vi.fn(() => Promise.resolve());
    Object.assign(popupDocument.defaultView, { navigator: { clipboard: { writeText: popupWrite } } });
    Object.assign(openerDocument.defaultView, { navigator: { clipboard: { writeText: openerWrite } } });
    const engine = (viewer as unknown as { engine: {
      currentWorksheet: Worksheet;
      selectionController: { select(cell: { row: number; col: number }): void };
    } }).engine;
    engine.currentWorksheet = {
      ...worksheet('Popup'),
      rows: [{
        index: 1,
        cells: [{ row: 1, col: 1, value: { type: 'text', text: 'popup' } }],
      }],
    } as unknown as Worksheet;
    engine.selectionController.select({ row: 1, col: 1 });
    popupDocument.dispatchEvent('keydown', { key: 'c', ctrlKey: true });
    expect(popupWrite).not.toHaveBeenCalled();
    viewportInput.dispatch('keydown', {
      key: 'c', ctrlKey: true, metaKey: false, defaultPrevented: false,
      isComposing: false, target: viewportInput, preventDefault() {},
    });
    await Promise.resolve();
    expect(popupWrite).toHaveBeenCalledWith('popup');
    expect(openerWrite).not.toHaveBeenCalled();

    viewer.destroy();
    expect(popupDocument.listenerCount('keydown')).toBe(0);
    expect(viewportInput._listeners.get('keydown')).toHaveLength(0);
  });

  it('uses the caller canvas with native scrollbars and without workbook footer chrome', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    canvas.clientWidth = 640;
    canvas.clientHeight = 360;
    parent.appendChild(canvas);

    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);

    expect(viewer.canvasElement).toBe(canvas);
    expect(descendants(parent).filter((element) => element.tag === 'canvas')).toContain(canvas);
    expect(descendants(parent).filter((element) => element.tag === 'button')).toHaveLength(0);
    const viewportInput = descendants(parent).find(
      (element) => element.getAttribute('data-xlsx-viewport-input') === 'sheet',
    );
    expect(viewportInput?.style.overflow).toBe('auto');
    expect(viewportInput?.children).toHaveLength(1);
    expect(descendants(parent).some((element) => element.style.overflow === 'auto')).toBe(true);

    viewer.destroy();
  });

  it('can hide native sheet scrollbars without adding workbook footer chrome', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);

    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement, {
      showScrollbars: false,
    });

    const viewportInput = descendants(parent).find(
      (element) => element.getAttribute('data-xlsx-viewport-input') === 'sheet',
    );
    expect(viewportInput?.style.overflow).toBe('clip');
    expect(viewportInput?.children).toHaveLength(0);
    expect(descendants(parent).filter((element) => element.tag === 'button')).toHaveLength(0);

    viewer.destroy();
  });

  it('continues drag selection beyond the visible viewport while auto-scrolling', () => {
    const doc = installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const engine = (viewer as unknown as { engine: {
      currentWorksheet: Worksheet;
      canvasArea: FakeEl;
      scrollHost: FakeEl;
      viewport: {
        setExtent(width: number, height: number): void;
        setViewportSize(width: number, height: number): void;
      };
      getCellAt(clientX: number, clientY: number): { row: number; col: number } | null;
      scheduleRender(): void;
    } }).engine;
    engine.currentWorksheet = {
      ...worksheet('Long sheet'),
      defaultColWidth: 8.43,
    };
    engine.canvasArea.clientWidth = 800;
    engine.canvasArea.clientHeight = 600;
    // Simulate classic scrollbars that reserve a 20px gutter inside the
    // canvasArea (unlike overlay scrollbars on macOS).
    engine.scrollHost.clientWidth = 780;
    engine.scrollHost.clientHeight = 580;
    engine.scrollHost.scrollWidth = 5_000;
    engine.scrollHost.scrollHeight = 5_000;
    engine.viewport.setExtent(5_000, 5_000);
    engine.viewport.setViewportSize(800, 600);
    engine.scheduleRender = () => undefined;

    const frames: FrameRequestCallback[] = [];
    Object.assign(doc.defaultView, {
      requestAnimationFrame: (callback: FrameRequestCallback) => {
        frames.push(callback);
        return frames.length;
      },
      cancelAnimationFrame: () => undefined,
    });
    const pointer = (overrides: Record<string, unknown>) => ({
      button: 0,
      pointerId: 1,
      pointerType: 'mouse',
      clientX: 100,
      clientY: 100,
      shiftKey: false,
      preventDefault: () => undefined,
      ...overrides,
    });

    engine.scrollHost.dispatch('pointerdown', pointer({
      pointerId: 2,
      pointerType: 'touch',
      clientX: 400,
      clientY: 300,
    }));
    engine.scrollHost.dispatch('pointerdown', pointer({}));
    const primarySelection = viewer.selectionState;
    engine.scrollHost.dispatch('pointerup', pointer({
      pointerId: 2,
      pointerType: 'touch',
      clientX: 400,
      clientY: 300,
    }));
    expect(viewer.selectionState).toEqual(primarySelection);

    engine.scrollHost.dispatch('pointerdown', pointer({
      pointerId: 2,
      pointerType: 'touch',
      clientX: 400,
      clientY: 300,
    }));
    engine.scrollHost.dispatch('pointerup', pointer({
      pointerId: 2,
      pointerType: 'touch',
      clientX: 400,
      clientY: 300,
    }));
    engine.scrollHost.dispatch('pointermove', pointer({
      pointerId: 2,
      clientX: 1_600,
      clientY: 1_200,
    }));
    engine.scrollHost.dispatch('pointerup', pointer({ pointerId: 2 }));
    engine.scrollHost.dispatch('pointercancel', pointer({ pointerId: 2 }));
    expect(frames).toHaveLength(0);
    expect(viewer.selectionState).toEqual(primarySelection);

    engine.scrollHost.dispatch('pointermove', pointer({ clientX: 1_600, clientY: 1_200 }));
    const visibleEdgeCell = engine.getCellAt(779, 579);
    expect(viewer.selectionState?.areas[0]).toMatchObject({
      kind: 'cells', bottom: visibleEdgeCell?.row, right: visibleEdgeCell?.col,
    });
    const visibleEdgeSelection = viewer.selectionState;
    engine.scrollHost.dispatch('pointerup', pointer({ pointerId: 2 }));
    engine.scrollHost.dispatch('pointercancel', pointer({ pointerId: 2 }));
    for (let frame = 1; frame <= 20; frame += 1) {
      const callback = frames.shift();
      expect(callback).toBeDefined();
      callback?.(frame * 16);
    }

    expect(viewer.getViewportOffset().x).toBeGreaterThan(0);
    expect(viewer.getViewportOffset().y).toBeGreaterThan(0);
    const finalArea = viewer.selectionState?.areas[0];
    const visibleArea = visibleEdgeSelection?.areas[0];
    expect(finalArea?.kind).toBe('cells');
    expect(visibleArea?.kind).toBe('cells');
    if (finalArea?.kind === 'cells' && visibleArea?.kind === 'cells') {
      expect(finalArea.bottom).toBeGreaterThan(visibleArea.bottom);
      expect(finalArea.right).toBeGreaterThan(visibleArea.right);
    }

    engine.scrollHost.dispatch('pointerup', pointer({ clientX: 1_600, clientY: 1_200 }));
    viewer.destroy();
  });

  it('restores the exact caller-owned canvas position, style, and bitmap dimensions', () => {
    installDom();
    const parent = makeContainer();
    const before = makeEl('span');
    const canvas = makeEl('canvas');
    const after = makeEl('span');
    canvas.setAttribute('style', 'width:320px;height:180px;border:1px solid red');
    canvas.style.cssText = 'width:320px;height:180px;border:1px solid red';
    canvas.width = 960;
    canvas.height = 540;
    parent.appendChild(before);
    parent.appendChild(canvas);
    parent.appendChild(after);

    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    expect(parent.children).toEqual([before, parent.children[1], after]);
    expect(parent.children[1]).not.toBe(canvas);

    viewer.destroy();
    viewer.destroy();

    expect(parent.children).toEqual([before, canvas, after]);
    expect(canvas.getAttribute('style')).toBe('width:320px;height:180px;border:1px solid red');
    expect(canvas.width).toBe(960);
    expect(canvas.height).toBe(540);
  });

  it('retains immutable query snapshots and closes every mutation after destroy', async () => {
    installDom();
    const canvas = makeEl('canvas');
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement, {
      cellScale: 1.25,
      hiddenSheetMode: 'dim',
    });
    viewer.destroy();

    expect(viewer.canvasElement).toBe(canvas);
    expect(viewer.sheetIndex).toBe(0);
    expect(viewer.sheetCount).toBe(0);
    expect(viewer.sheetNames).toEqual([]);
    expect(viewer.getViewportOffset()).toEqual({ x: 0, y: 0 });
    expect(viewer.selectionState).toBeNull();
    expect(viewer.getScale()).toBe(1.25);
    expect(viewer.hiddenSheetMode).toBe('dim');
    expect(viewer.visibleSheetCount).toBe(0);
    expect(viewer.getCellAt(0, 0)).toBeNull();

    const closed = 'XlsxSheetViewer is destroyed';
    await expect(viewer.load(new ArrayBuffer(0))).rejects.toThrow(closed);
    await expect(viewer.goToSheet(0)).rejects.toThrow(closed);
    await expect(viewer.nextSheet()).rejects.toThrow(closed);
    await expect(viewer.prevSheet()).rejects.toThrow(closed);
    await expect(viewer.setViewportOffset({ x: 0, y: 0 })).rejects.toThrow(closed);
    await expect(viewer.scrollToCell('A1')).rejects.toThrow(closed);
    await expect(viewer.goToComment(0, 'A1')).rejects.toThrow(closed);
    await expect(viewer.relayout()).rejects.toThrow(closed);
    await expect(viewer.setHiddenSheetMode('show')).rejects.toThrow(closed);
    await expect(viewer.findText('x')).rejects.toThrow(closed);
    await expect(viewer.findNext()).rejects.toThrow(closed);
    await expect(viewer.findPrev()).rejects.toThrow(closed);
    await expect(viewer.getResourceMetrics()).rejects.toThrow(closed);
    expect(() => viewer.setScale(1)).toThrow(closed);
    expect(() => viewer.zoomIn()).toThrow(closed);
    expect(() => viewer.zoomOut()).toThrow(closed);
    expect(() => viewer.fitWidth()).toThrow(closed);
    expect(() => viewer.fitPage()).toThrow(closed);
    expect(() => viewer.setSelection('A1')).toThrow(closed);
    expect(() => viewer.setSelectionColor('#000')).toThrow(closed);
    expect(() => viewer.clearFind()).toThrow(closed);
  });

  it('retains the successful load metrics snapshot after destroy without an explicit metrics query', async () => {
    installDom();
    const metrics: OoxmlResourceMetrics = {
      schemaVersion: 1,
      scope: 'load',
      format: 'xlsx',
      mode: 'main',
      status: 'ok',
      sourceBytes: 12,
      elapsedMs: 3,
      policy: {
        maxArchiveEntryBytes: null,
        maxTotalInflatedBytes: null,
        maxArchiveEntries: null,
      },
      checkpoints: [],
    };
    const workbook = {
      sheetNames: ['Sheet1'],
      tabColors: {} as Record<number, string>,
      destroy: vi.fn(),
      getResourceMetrics: vi.fn().mockResolvedValue(metrics),
    } as unknown as XlsxWorkbook;
    const loadSpy = vi.spyOn(XlsxWorkbook, 'load').mockImplementation(async (_source, options) => {
      options?.onResourceMetrics?.(metrics);
      return workbook;
    });

    const canvas = makeEl('canvas');
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement, { password: 'secret' });
    const engine = (viewer as unknown as { engine: { showSheet(index: number): Promise<void> } }).engine;
    vi.spyOn(engine, 'showSheet').mockResolvedValue(undefined);

    await viewer.load(new ArrayBuffer(0));
    expect(loadSpy).toHaveBeenCalledWith(
      expect.any(ArrayBuffer),
      expect.objectContaining({ password: 'secret' }),
    );
    viewer.destroy();

    await expect(viewer.getResourceMetrics()).resolves.toEqual(metrics);
    expect(workbook.getResourceMetrics).not.toHaveBeenCalled();
  });

  it('retains terminal error metrics after a rejected load and destroy', async () => {
    installDom();
    const metrics: OoxmlResourceMetrics = {
      schemaVersion: 1,
      scope: 'load',
      format: 'xlsx',
      mode: 'main',
      status: 'error',
      sourceBytes: 12,
      elapsedMs: 3,
      policy: {
        maxArchiveEntryBytes: null,
        maxTotalInflatedBytes: null,
        maxArchiveEntries: null,
      },
      checkpoints: [],
      error: { code: 'ooxml-resource-limit', stage: 'package-open' },
    };
    const failure = new Error('load failed');
    vi.spyOn(XlsxWorkbook, 'load').mockImplementation(async (_source, options) => {
      options?.onResourceMetrics?.(metrics);
      throw failure;
    });

    const viewer = new XlsxSheetViewer(makeEl('canvas') as unknown as HTMLCanvasElement);
    await expect(viewer.load(new ArrayBuffer(0))).rejects.toBe(failure);
    viewer.destroy();

    await expect(viewer.getResourceMetrics()).resolves.toEqual(metrics);
  });

  it('borrows one loaded workbook, opens the requested sheet, and leaves workbook cleanup to the caller', async () => {
    installDom();
    const destroy = vi.fn();
    const getWorksheet = vi.fn((index: number) => Promise.resolve(worksheet(`Sheet${index + 1}`)));
    const workbook = {
      mode: 'main',
      sheetCount: 2,
      sheetNames: ['Sheet1', 'Sheet2'],
      tabColors: {} as Record<number, string>,
      getWorksheet,
      isHidden: () => false,
      destroy,
    } as unknown as XlsxWorkbook;
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);

    const viewer = XlsxSheetViewer.fromWorkbook(
      canvas as unknown as HTMLCanvasElement,
      workbook,
    );
    await viewer.goToSheet(1);

    expect(viewer.sheetIndex).toBe(1);
    expect(getWorksheet).toHaveBeenCalledWith(1);
    expect(getWorksheet).not.toHaveBeenCalledWith(0);
    await expect((viewer as XlsxSheetViewer).load(new ArrayBuffer(0))).rejects.toThrow(
      'XlsxSheetViewer.load() is unsupported on a Viewer created by fromWorkbook()',
    );

    viewer.destroy();
    expect(destroy).not.toHaveBeenCalled();
    expect(parent.children).toEqual([canvas]);
  });

  it('validates a borrowed mode conflict before mounting the caller canvas', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const workbook = {
      mode: 'worker',
      sheetCount: 1,
      sheetNames: ['Sheet1'],
      destroy: vi.fn(),
    } as unknown as XlsxWorkbook;

    expect(() => XlsxSheetViewer.fromWorkbook(
      canvas as unknown as HTMLCanvasElement,
      workbook,
      { mode: 'main' } as never,
    )).toThrow("opts.mode='main' conflicts with the borrowed engine's mode='worker'");
    expect(parent.children).toEqual([canvas]);
  });

  it('keeps mutable view state independent when two viewers borrow the same cached sheet', async () => {
    installDom();
    const source = worksheet('Shared');
    const workbook = {
      mode: 'main',
      sheetCount: 1,
      sheetNames: ['Shared'],
      tabColors: {} as Record<number, string>,
      getWorksheet: vi.fn().mockResolvedValue(source),
      isHidden: () => false,
      destroy: vi.fn(),
    } as unknown as XlsxWorkbook;
    const first = XlsxSheetViewer.fromWorkbook(
      makeEl('canvas') as unknown as HTMLCanvasElement,
      workbook,
    );
    const second = XlsxSheetViewer.fromWorkbook(
      makeEl('canvas') as unknown as HTMLCanvasElement,
      workbook,
    );
    await Promise.all([first.goToSheet(0), second.goToSheet(0)]);
    const firstWorksheet = (first as unknown as {
      engine: { currentWorksheet: Worksheet };
    }).engine.currentWorksheet;
    const secondWorksheet = (second as unknown as {
      engine: { currentWorksheet: Worksheet };
    }).engine.currentWorksheet;

    firstWorksheet.rowHeights[1] = 40;
    firstWorksheet.colWidths[1] = 16;

    expect(firstWorksheet).not.toBe(secondWorksheet);
    expect(secondWorksheet.rowHeights[1]).toBeUndefined();
    expect(secondWorksheet.colWidths[1]).toBeUndefined();
    expect(source.rowHeights[1]).toBeUndefined();
    expect(source.colWidths[1]).toBeUndefined();

    first.destroy();
    second.destroy();
  });
});

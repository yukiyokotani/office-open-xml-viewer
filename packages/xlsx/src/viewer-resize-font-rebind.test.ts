import { afterEach, describe, expect, it, vi } from 'vitest';
import { XlsxViewer } from './viewer.js';
import { installDom, makeContainer, type FakeEl } from './viewer-destroy-test-dom.js';
import { bindXlsxOfficeFontRoutes, getGridGeometryForWorksheet } from './renderer.js';
import { WorksheetViewProjectionCache, type WireSizeOverrides } from './worker-protocol.js';
import { GridGeometry } from './internal/grid-geometry.js';
import type { Worksheet } from './types.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

function sheet(): Worksheet {
  return {
    name: 'Resize',
    rows: [],
    colWidths: { 1: 10, 2: 12 },
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    defaultFontFamily: 'Metric Test',
    defaultFontSize: 11,
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

interface ViewerPriv {
  wb: unknown;
  currentWorksheet: Worksheet;
  currentSheet: number;
  canvasArea: FakeEl;
  viewEdits: {
    wireSizeOverrides(sheetIndex: number):
      { overrides: WireSizeOverrides; revision: number } | undefined;
  };
  selectionInput: {
    resizeDrag: { kind: 'col' | 'row'; index: number; originScaled: number; mdw: number } | null;
    applyResize(clientX: number, clientY: number): void;
  };
  buildOutline(ws: Worksheet): void;
}

/** Controlled Canvas metric: digit advance 7, later 8. This is a test double,
 * not a claim about browser-native font fidelity. */
function metricContext(): { setDigit(width: number): void; ctx: CanvasRenderingContext2D } {
  let digit = 7;
  const ctx = {
    canvas: {},
    font: '',
    save() {},
    restore() {},
    measureText: () => ({ width: digit }),
  };
  return {
    setDigit: (width: number) => { digit = width; },
    ctx: ctx as unknown as CanvasRenderingContext2D,
  };
}

function buildViewer(): ViewerPriv {
  installDom();
  const container = makeContainer();
  const viewer = new XlsxViewer(container as unknown as HTMLElement, { mode: 'worker' });
  const priv = viewer as unknown as ViewerPriv;
  priv.wb = {
    renderViewportToBitmap: vi.fn(() => new Promise<ImageBitmap>(() => {})),
    sheetNames: ['Resize'],
    sheetCount: 1,
    destroy: vi.fn(),
  };
  priv.currentWorksheet = sheet();
  priv.currentSheet = 0;
  priv.canvasArea.clientWidth = 800;
  priv.canvasArea.clientHeight = 600;
  priv.buildOutline(priv.currentWorksheet);
  return priv;
}

describe('SelectionInput resize across a retained-font MDW rebind', () => {
  it('keeps the user 84 CSS px while authored width 12 follows MDW 7 -> 8', () => {
    const priv = buildViewer();
    const ws = priv.currentWorksheet;
    const metric = metricContext();
    bindXlsxOfficeFontRoutes(metric.ctx, ws);
    expect(getGridGeometryForWorksheet(ws).maximumDigitWidth).toBe(7);

    const left = (priv.canvasArea as unknown as { getBoundingClientRect(): DOMRect })
      .getBoundingClientRect().left;
    priv.selectionInput.resizeDrag = { kind: 'col', index: 1, originScaled: 0, mdw: 7 };
    priv.selectionInput.applyResize(left + 84, 0);
    expect(ws.colWidths[1]).toBe(12);
    expect(ws.colWidths[2]).toBe(12);
    const at7 = getGridGeometryForWorksheet(ws);
    expect(at7.col.sizeOf(1)).toBe(84);
    expect(at7.col.sizeOf(2)).toBe(84);

    metric.setDigit(8);
    bindXlsxOfficeFontRoutes(metric.ctx, ws);
    const at8 = getGridGeometryForWorksheet(ws);
    expect(at8.maximumDigitWidth).toBe(8);
    expect(at8.col.sizeOf(1)).toBe(84);
    expect(at8.col.sizeOf(2)).toBe(96);
    expect(at8.axesAtScale(1.25).col.sizeOf(1)).toBe(105);
    expect(at8.axesAtScale(1.25).col.sizeOf(2)).toBe(120);
    // Shared axis consumed by hit testing / overlays / anchors.
    const rect = at8.cellRect(1, 3, { scale: 1, scrollX: 0, scrollY: 0, headerWidth: 0, headerHeight: 0 });
    expect(rect?.x).toBe(84 + 96);
    expect(at8.axesAtScale(1).col.offsetOf(3)).toBe(84 + 96);

    // The worker wire carries the same pixel intent to a render-local projection.
    const wire = priv.viewEdits.wireSizeOverrides(0);
    expect(wire?.overrides.columnCssWidths).toEqual({ 1: 84 });
    const projected = new WorksheetViewProjectionCache()
      .resolve(sheet(), 0, { id: 7, revision: wire!.revision }, wire!.overrides).worksheet;
    expect(GridGeometry.forWorksheet(projected, 8).col.sizeOf(1)).toBe(84);
    expect(GridGeometry.forWorksheet(projected, 8).col.sizeOf(2)).toBe(96);
  });
});

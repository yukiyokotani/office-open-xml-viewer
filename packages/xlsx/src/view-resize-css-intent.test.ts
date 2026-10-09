import { describe, expect, it } from 'vitest';
import type { Styles, Worksheet } from './types.js';
import { SheetViewEdits } from './internal/viewer/sheet-view-edits.js';
import { GridGeometry } from './internal/grid-geometry.js';
import {
  applySizeOverrides,
  createSizeOverriddenWorksheet,
  WorksheetViewProjectionCache,
  type WireSizeOverrides,
} from './worker-protocol.js';
import { colWidthToPx, inheritSheetRenderCache, pxToColWidth } from './renderer.js';
import { worksheetWithAutoRowHeights } from './render-orchestrator.js';
import { createSheetViewModel } from './internal/sheet-viewer-runtime.js';
import {
  bindInitialAnchorSizes,
  captureInitialAnchorSizes,
  resolveWorksheetAnchorRect,
} from './internal/initial-anchor-sizes.js';

/**
 * Settled ACTIVE view-only resize policy: a user drag captures logical CSS px
 * and that pixel intent survives a later MDW change, while authored stored
 * widths keep decoding through the current MDW. The raw `colWidths` numbers of
 * a user-resized column and an authored column can be identical (12 here).
 */
function sheet(): Worksheet {
  return {
    name: 'Resize',
    rows: [],
    colWidths: { 1: 12, 2: 12 },
    colWidthRanges: [{ min: 4, max: 6, width: 12 }],
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    images: [],
    charts: [],
    shapeGroups: [],
  } as unknown as Worksheet;
}

function clone(source: Worksheet): Worksheet {
  return { ...source, colWidths: { ...source.colWidths }, rowHeights: { ...source.rowHeights } };
}

/** Store-level mutations do not invalidate geometry (their callers do), so
 * these helpers rebuild explicitly before reading the axis. */
function colPx(ws: Worksheet, mdw: number, index: number, scale = 1): number {
  GridGeometry.invalidate(ws);
  return GridGeometry.forWorksheet(ws, mdw).axesAtScale(scale).col.sizeOf(index);
}

/** The model writes SelectionInput.applyResize performs at captured MDW 7. */
function userResize(store: SheetViewEdits, ws: Worksheet, index: number, px: number): void {
  ws.colWidths[index] = pxToColWidth(px, 7);
  store.recordSizeOverride(ws, 0, 'col', index, px);
}

function edits(): SheetViewEdits {
  const store = new SheetViewEdits();
  store.bindSheet(0);
  return store;
}

describe('view-only column resize keeps canonical CSS pixel intent', () => {
  it('bumps the projection revision for CSS intent even when raw width 12 is unchanged', () => {
    const ws = sheet();
    const store = edits();
    store.recordSizeOverride(ws, 0, 'col', 1);
    const before = store.sizeRevision(0);
    expect(store.wireSizeOverrides(0)?.overrides).toEqual({ cols: { 1: 12 } });
    const cache = new WorksheetViewProjectionCache();
    const workerSource = sheet();
    const previous = cache.resolve(workerSource, 0, { id: 3, revision: before },
      store.wireSizeOverrides(0)?.overrides).worksheet;

    userResize(store, ws, 1, 84);
    expect(ws.colWidths[1]).toBe(12);
    expect(store.sizeRevision(0)).toBe(before + 1);
    expect(store.wireSizeOverrides(0)?.overrides).toEqual({
      cols: { 1: 12 },
      columnCssWidths: { 1: 84 },
    });
    const replayed = cache.resolve(workerSource, 0, { id: 3, revision: store.sizeRevision(0) },
      store.wireSizeOverrides(0)?.overrides).worksheet;
    expect(replayed).not.toBe(previous);
    expect(colPx(previous, 8, 1)).toBe(96);
    expect(colPx(replayed, 8, 1)).toBe(84);

    // A repeated same-value drag is not a new revision.
    userResize(store, ws, 1, 84);
    expect(store.sizeRevision(0)).toBe(before + 1);

    // MDW 7 -> 8: user pixel width stays 84, authored 12 follows to 96.
    expect(colPx(ws, 8, 1)).toBe(84);
    expect(colPx(ws, 8, 2)).toBe(96);
    // Historical per-band Math.round at a noninteger zoom is unchanged.
    expect(colPx(ws, 8, 1, 1.25)).toBe(105);
    expect(colPx(ws, 8, 2, 1.25)).toBe(120);
  });

  it('keeps CSS intent in the real auto-height clone and refreshes it after a later resize', () => {
    const source = sheet();
    const store = edits();
    userResize(store, source, 1, 84);
    GridGeometry.forWorksheet(source, 8);
    const ctx = {
      canvas: {}, font: '', save() {}, restore() {},
      measureText: () => ({ width: 8 }),
    } as unknown as CanvasRenderingContext2D;
    const styles = { fonts: [], fills: [], borders: [], cellXfs: [], numFmts: [], dxfs: [] } as Styles;

    const first = worksheetWithAutoRowHeights(ctx, source, styles);
    expect(first).not.toBe(source);
    expect(GridGeometry.forWorksheet(first, 8).col.sizeOf(1)).toBe(84);
    // Main/worker apply invalidate the source geometry even when only the
    // private CSS channel changes. Auto-height's cache must observe that too.
    applySizeOverrides(source, { columnCssWidths: { 1: 90 } });
    const next = worksheetWithAutoRowHeights(ctx, source, styles);
    expect(GridGeometry.forWorksheet(next, 8).col.sizeOf(1)).toBe(90);
    expect(GridGeometry.forWorksheet(first, 8).col.sizeOf(1)).toBe(84);
  });

  it('keeps prepared oneCell anchor sizes across fresh viewer and worker projections', () => {
    const source = sheet();
    const anchor = {
      anchorTag: 'twoCellAnchor' as const,
      editAs: 'oneCell',
      fromCol: 1, fromColOff: 0, toCol: 2, toColOff: 0,
      fromRow: 0, fromRowOff: 0, toRow: 1, toRowOff: 0,
    };
    source.images = [anchor, { ...anchor, editAs: 'twoCell' }] as unknown as Worksheet['images'];
    const main = createSheetViewModel(source);
    const reference = captureInitialAnchorSizes(main, GridGeometry.forWorksheet(main, 7));
    expect(reference).toBeDefined();
    bindInitialAnchorSizes(main, reference!);
    const store = edits();
    userResize(store, main, 1, 84);
    const fresh = createSheetViewModel(source);
    bindInitialAnchorSizes(fresh, reference!);
    store.restoreSheetViewState(0, fresh);
    const wire = store.wireSizeOverrides(0)!;
    const worker = new WorksheetViewProjectionCache().resolve(source, 0,
      { id: 5, revision: wire.revision, initialAnchorSizes: structuredClone(reference!) },
      structuredClone(wire.overrides)).worksheet;
    for (const projection of [main, fresh, worker]) {
      const axes = GridGeometry.forWorksheet(projection, 8).axesAtScale(1.25);
      const retained = resolveWorksheetAnchorRect(projection, projection.images[0], axes.col, axes.row, 1.25);
      const moving = resolveWorksheetAnchorRect(projection, projection.images[1], axes.col, axes.row, 1.25);
      expect([retained.x, retained.width, moving.x, moving.width]).toEqual([105, 105, 105, 120]);
    }
    // Legacy untagged marker behaviour continues to use the authored axis.
    const legacy = { ...anchor, anchorTag: undefined };
    const axes = GridGeometry.forWorksheet(source, 8).axesAtScale(1);
    expect(resolveWorksheetAnchorRect(source, legacy, axes.col, axes.row, 1).width).toBe(96);
    expect(source.colWidths).toEqual({ 1: 12, 2: 12 });
  });

  it('replays CSS intent into fresh and worker projections without touching the source', () => {
    const source = sheet();
    const main = clone(source);
    const store = edits();
    userResize(store, main, 1, 84);

    const fresh = clone(source);
    store.restoreSheetViewState(0, fresh);
    expect(colPx(fresh, 8, 1)).toBe(84);
    expect(colPx(fresh, 8, 2)).toBe(96);

    const wire = store.wireSizeOverrides(0);
    expect(wire).toBeDefined();
    const cache = new WorksheetViewProjectionCache();
    const projected = cache.resolve(
      source, 0, { id: 1, revision: wire!.revision }, wire!.overrides,
    ).worksheet;
    expect(projected).not.toBe(source);
    expect(colPx(projected, 8, 1)).toBe(84);
    expect(colPx(projected, 8, 2)).toBe(96);
    // Renderer cache inheritance must not drop the projection sidecar.
    inheritSheetRenderCache(source, projected);
    expect(colPx(projected, 8, 1)).toBe(84);

    // Sibling projection and the immutable source keep the authored decode.
    const sibling = cache.resolve(source, 0, { id: 2, revision: 0 }, undefined).worksheet;
    expect(colPx(sibling, 8, 1)).toBe(96);
    expect(colPx(source, 8, 1)).toBe(96);
    expect(source.colWidths).toEqual({ 1: 12, 2: 12 });
  });

  it('stashes CSS intent (and its absence) across hide/restore of point, range and default columns', () => {
    const ws = sheet();
    const store = edits();
    userResize(store, ws, 1, 84); // authored point column
    userResize(store, ws, 5, 84); // range-backed column, point created by the drag
    const indices = [1, 3, 5, 6]; // 3 = default (no point), 6 = range only
    for (const index of indices) store.setBandHidden(ws, 0, 'col', index, true);
    for (const index of indices) expect(colPx(ws, 8, index)).toBe(0);

    const hiddenWire = store.wireSizeOverrides(0)!.overrides;
    expect(hiddenWire.cols).toEqual({ 1: 0, 3: 0, 5: 0, 6: 0 });
    expect(hiddenWire.columnCssWidths).toEqual({ 1: null, 5: null });
    const hiddenProjection = createSizeOverriddenWorksheet(sheet(), hiddenWire);
    expect(colPx(hiddenProjection, 8, 1)).toBe(0);
    expect(colPx(hiddenProjection, 8, 5)).toBe(0);

    for (const index of indices) store.setBandHidden(ws, 0, 'col', index, false);
    expect(colPx(ws, 8, 1)).toBe(84);
    expect(colPx(ws, 8, 5)).toBe(84);
    expect(colPx(ws, 8, 6)).toBe(96);
    expect(colPx(ws, 8, 3)).toBe(colWidthToPx(8.43, 8));
    expect(Object.hasOwn(ws.colWidths, 3)).toBe(false);
    expect(Object.hasOwn(ws.colWidths, 6)).toBe(false);

    const wire = store.wireSizeOverrides(0)!.overrides;
    expect(wire.cols).toEqual({ 1: 12, 3: null, 5: 12, 6: null });
    expect(wire.columnCssWidths).toEqual({ 1: 84, 5: 84 });
    const restored = createSizeOverriddenWorksheet(sheet(), wire);
    expect(colPx(restored, 8, 1)).toBe(84);
    expect(colPx(restored, 8, 5)).toBe(84);
    expect(colPx(restored, 8, 6)).toBe(96);
    expect(colPx(restored, 8, 3)).toBe(colWidthToPx(8.43, 8));
  });

  it('applies wire CSS overrides with invalidation: null removes the override, zero hides', () => {
    const source = sheet();
    const initial: WireSizeOverrides = { cols: { 1: 12 }, columnCssWidths: { 1: 84 } };
    const ws = createSizeOverriddenWorksheet(source, initial);
    expect(GridGeometry.forWorksheet(ws, 8).col.sizeOf(1)).toBe(84);
    applySizeOverrides(ws, { columnCssWidths: { 1: 0 } });
    expect(GridGeometry.forWorksheet(ws, 8).col.sizeOf(1)).toBe(0);
    applySizeOverrides(ws, { columnCssWidths: { 1: null } });
    expect(GridGeometry.forWorksheet(ws, 8).col.sizeOf(1)).toBe(96);
    expect(GridGeometry.forWorksheet(source, 8).col.sizeOf(1)).toBe(96);
  });

  it('HYPOTHETICAL internal metric substitution (not a current production bug)', () => {
    // computeMdw quantizes desktop MDW to an integer; a fractional MDW such as
    // this only arises if a caller passes it directly to GridGeometry.
    const mdw = 7.430206298828125;
    expect(GridGeometry.forWorksheet(sheet(), mdw).col.sizeOf(2)).toBe(89);
    const projected = createSizeOverriddenWorksheet(sheet(), {
      cols: { 1: 12 },
      columnCssWidths: { 1: 84 },
    });
    expect(GridGeometry.forWorksheet(projected, mdw).col.sizeOf(1)).toBe(84);
  });

  it('documents the existing per-band rounding of fractional scaled bands (out of scope)', () => {
    // A synthetic private wire value, not a current drag (which captures
    // integer CSS px) or an Office/native-unit selector. Fractional logical
    // bands survive construction but the established display axes round them.
    const projected = createSizeOverriddenWorksheet(sheet(), {
      columnCssWidths: { 1: 90.5, 2: 90.5 },
    });
    const geometry = GridGeometry.forWorksheet(projected, 7);
    expect(geometry.col.offsetOf(3)).toBe(181);
    const axes = geometry.axesAtScale(1);
    expect(axes.col.sizeOf(1)).toBe(91);
    const anchor = {
      fromCol: 0, fromColOff: 0, toCol: 2, toColOff: 0,
      fromRow: 0, fromRowOff: 0, toRow: 1, toRowOff: 0,
    };
    expect(resolveWorksheetAnchorRect(projected, anchor, axes.col, axes.row, 1).width).toBe(182);
  });
});

import type { Worksheet } from '../../types.js';
import type { WireSizeOverrides } from '../../worker-protocol.js';
import { derivedAutoRowHeights } from '../../renderer.js';
import type { OutlineAxis } from '../../outline.js';
import { getColumnCssWidth, setColumnCssWidth } from '../column-css-overrides.js';

type OutlineState = {
  rowCollapsed: Map<number, boolean>;
  colCollapsed: Map<number, boolean>;
  stashedRowHeights: Map<number, number | undefined>;
  stashedColWidths: Map<number, number | undefined>;
  /** Paired CSS pixel intent; `undefined` stashes CSS absence. */
  stashedColCssWidths: Map<number, number | undefined>;
};

type SizeOverrides = {
  rows: Map<number, number | null>;
  automaticRows: Map<number, number>;
  cols: Map<number, number | null>;
  /** Canonical logical CSS px of user column resizes; `null` = removed. */
  colCss: Map<number, number | null>;
  revision: number;
  wire?: WireSizeOverrides;
};

/**
 * Viewer-owned, view-only worksheet edits: outline collapse/expand and
 * drag-to-resize. The workbook cache stays immutable; this store records the
 * per-sheet deltas that are replayed onto every fresh viewer projection and
 * serialized into each render request, so both render modes draw the same
 * projection as the pointer/overlay geometry.
 */
export class SheetViewEdits {
  /** Original sizes stashed while the bound sheet has collapsed bands. The
   * maps belong to {@link outlineStateStore} and survive projection eviction. */
  private stashedRowHeights = new Map<number, number | undefined>();
  private stashedColWidths = new Map<number, number | undefined>();
  private stashedColCssWidths = new Map<number, number | undefined>();
  /** Only user-mutated outline flags and pre-collapse sizes survive a sheet
   * switch. The parser's filters and frozen panes are read-only here; selection
   * and scroll are viewer viewport state, reset by navigation as before. */
  private readonly outlineStateStore = new Map<number, OutlineState>();
  /**
   * Per-sheet cumulative record of every view-only size mutation (outline
   * collapse/expand, drag-to-resize #567), keyed by sheet index. Value = the
   * band's current model size, or `null` when the model has no entry (default
   * size). Serialized as {@link WireSizeOverrides} with every render so both
   * modes draw from a render-local projection matching this viewer, while the
   * workbook cache remains immutable for sibling viewers. Entries are updated
   * in place and never removed; the whole store resets with a new workbook.
   */
  readonly sizeOverrideStore = new Map<number, SizeOverrides>();

  /** Forget every sheet's edits (a new workbook was prepared). */
  clear(): void {
    this.sizeOverrideStore.clear();
    this.outlineStateStore.clear();
  }

  /** Bind the pre-collapse size stashes to `sheetIndex`, creating its outline
   * state on first use. Called whenever a sheet's outline is (re)built. */
  bindSheet(sheetIndex: number): void {
    let state = this.outlineStateStore.get(sheetIndex);
    if (!state) {
      state = {
        rowCollapsed: new Map(), colCollapsed: new Map(),
        stashedRowHeights: new Map(), stashedColWidths: new Map(),
        stashedColCssWidths: new Map(),
      };
      this.outlineStateStore.set(sheetIndex, state);
    }
    this.stashedRowHeights = state.stashedRowHeights;
    this.stashedColWidths = state.stashedColWidths;
    this.stashedColCssWidths = state.stashedColCssWidths;
  }

  /** Set a row/column hidden by mapping to the size-0 encoding the axis/renderer
   *  already understand, stashing the original size so expand can restore it. */
  setBandHidden(
    ws: Worksheet,
    sheetIndex: number,
    axis: OutlineAxis,
    index: number,
    hidden: boolean,
  ): void {
    if (axis === 'row') {
      if (hidden) {
        if (!this.stashedRowHeights.has(index)) {
          this.stashedRowHeights.set(index, ws.rowHeights[index]);
        }
        ws.rowHeights[index] = 0;
      } else {
        if (this.stashedRowHeights.has(index)) {
          const orig = this.stashedRowHeights.get(index);
          if (orig === undefined) delete ws.rowHeights[index];
          else ws.rowHeights[index] = orig;
          this.stashedRowHeights.delete(index);
        } else if (ws.rowHeights[index] === 0) {
          // Was hidden in the source file (height 0) with no stash — reveal at
          // the default height.
          delete ws.rowHeights[index];
        }
      }
    } else {
      if (hidden) {
        if (!this.stashedColWidths.has(index)) {
          this.stashedColWidths.set(index, ws.colWidths[index]);
          this.stashedColCssWidths.set(index, getColumnCssWidth(ws, index));
        }
        ws.colWidths[index] = 0;
        // Hidden is size 0; a stale CSS override must not keep the band visible.
        setColumnCssWidth(ws, index, null);
      } else {
        if (this.stashedColWidths.has(index)) {
          const orig = this.stashedColWidths.get(index);
          if (orig === undefined) delete ws.colWidths[index];
          else ws.colWidths[index] = orig;
          // Restore the pixel intent, including its original absence.
          setColumnCssWidth(ws, index, this.stashedColCssWidths.get(index) ?? null);
          this.stashedColWidths.delete(index);
          this.stashedColCssWidths.delete(index);
        } else if (ws.colWidths[index] === 0) {
          delete ws.colWidths[index];
          setColumnCssWidth(ws, index, null);
        }
      }
    }
    // Mirror the post-mutation model value into the render override channel so
    // both modes can draw this viewer's projection without mutating the shared
    // workbook cache.
    this.recordSizeOverride(ws, sheetIndex, axis, index);
  }

  /** Record band `index`'s CURRENT model size (or `null` = no entry) in the
   *  per-sheet override store. Called after every view-only size mutation so
   *  both render modes receive this viewer's independent projection.
   *  `columnCssPx` is the logical CSS px a user column drag captured: it is
   *  stored as the column's canonical view-only width so a later MDW change
   *  keeps the drag's pixel size. Without it, the worksheet's current CSS
   *  override (possibly none) is mirrored. CSS metadata changes bump the
   *  revision even when the raw `colWidths` number is unchanged. */
  recordSizeOverride(
    ws: Worksheet,
    sheetIndex: number,
    axis: OutlineAxis,
    index: number,
    columnCssPx?: number,
  ): void {
    let entry = this.sizeOverrideStore.get(sheetIndex);
    if (!entry) {
      entry = {
        rows: new Map(), automaticRows: new Map(), cols: new Map(), colCss: new Map(), revision: 0,
      };
      this.sizeOverrideStore.set(sheetIndex, entry);
    }
    const target = axis === 'row' ? entry.rows : entry.cols;
    if (axis === 'row') entry.automaticRows.delete(index);
    const value = axis === 'row' ? ws.rowHeights[index] ?? null : ws.colWidths[index] ?? null;
    let changed = false;
    if (target.get(index) !== value) {
      target.set(index, value);
      changed = true;
    }
    if (axis === 'col') {
      if (columnCssPx !== undefined) setColumnCssWidth(ws, index, columnCssPx);
      const css = getColumnCssWidth(ws, index) ?? null;
      if (entry.colCss.has(index) ? entry.colCss.get(index) !== css : css !== null) {
        entry.colCss.set(index, css);
        changed = true;
      }
    }
    if (!changed) return;
    entry.revision++;
    entry.wire = undefined;
  }

  /** A sheet's override store serialized for the wire, or undefined when
   *  nothing has been mutated (keeps the request payload unchanged). */
  wireSizeOverrides(sheetIndex: number): Readonly<{
    overrides: WireSizeOverrides;
    revision: number;
  }> | undefined {
    const entry = this.sizeOverrideStore.get(sheetIndex);
    if (!entry || (entry.rows.size === 0 && entry.automaticRows.size === 0 && entry.cols.size === 0)) {
      return undefined;
    }
    if (!entry.wire) {
      const wire: WireSizeOverrides = {};
      if (entry.rows.size > 0 || entry.automaticRows.size > 0) {
        wire.rows = Object.fromEntries([...entry.automaticRows, ...entry.rows]);
      }
      if (entry.cols.size > 0) wire.cols = Object.fromEntries(entry.cols);
      if (entry.colCss.size > 0) wire.columnCssWidths = Object.fromEntries(entry.colCss);
      entry.wire = wire;
    }
    return { overrides: entry.wire, revision: entry.revision };
  }

  /** Row indices whose height the user set explicitly on `sheetIndex`. */
  manualRows(sheetIndex: number): Iterable<number> {
    return this.sizeOverrideStore.get(sheetIndex)?.rows.keys() ?? [];
  }

  /** Whether `sheetIndex` carries a manual view edit: a row/column resize or
   * an outline collapse/expand (outline writes also call recordSizeOverride).
   * Display-derived `automaticRows` are not edits. #1713 uses this to refuse
   * capturing a prepared-initial anchor reference from an edited projection. */
  hasViewEdits(sheetIndex: number): boolean {
    const sizes = this.sizeOverrideStore.get(sheetIndex);
    if (sizes && (sizes.rows.size > 0 || sizes.cols.size > 0)) return true;
    const outline = this.outlineStateStore.get(sheetIndex);
    return outline !== undefined && (outline.rowCollapsed.size > 0 || outline.colCollapsed.size > 0);
  }

  /** Current projection revision of `sheetIndex` (0 before any size record),
   * also used when only a prepared-initial anchor reference is transported. */
  sizeRevision(sheetIndex: number): number {
    return this.sizeOverrideStore.get(sheetIndex)?.revision ?? 0;
  }

  /** Mirror only display-derived heights into the worker projection channel.
   * Manual/authored sizes remain in `rows`, so a later column refit can replace
   * automatic values without reclassifying a user's row resize. */
  syncAutomaticRowOverrides(sheetIndex: number, worksheet: Worksheet): void {
    const next = new Map(derivedAutoRowHeights(worksheet));
    let entry = this.sizeOverrideStore.get(sheetIndex);
    if (!entry && next.size === 0) return;
    if (!entry) {
      entry = {
        rows: new Map(), automaticRows: new Map(), cols: new Map(), colCss: new Map(), revision: 0,
      };
      this.sizeOverrideStore.set(sheetIndex, entry);
    }
    entry.automaticRows = next;
    entry.revision++;
    entry.wire = undefined;
  }

  /** Update the `collapsed` flag on a band's model entry so the outline rebuild
   *  reflects the new state. */
  setBandCollapsed(
    ws: Worksheet,
    sheetIndex: number,
    axis: OutlineAxis,
    index: number,
    collapsed: boolean,
  ): void {
    const state = this.outlineStateStore.get(sheetIndex);
    (axis === 'row' ? state?.rowCollapsed : state?.colCollapsed)?.set(index, collapsed);
    if (axis === 'row') {
      const row = ws.rows.find((r) => r.index === index);
      if (row) row.collapsed = collapsed;
    } else {
      ws.colCollapsed = ws.colCollapsed ?? {};
      if (collapsed) ws.colCollapsed[index] = true;
      else delete ws.colCollapsed[index];
    }
  }

  /** Rebuild only the mutable projection fields. The large row/cell graph is
   * reacquired from the workbook cache and may have been evicted meanwhile. */
  restoreSheetViewState(sheetIndex: number, worksheet: Worksheet): void {
    const sizes = this.sizeOverrideStore.get(sheetIndex);
    if (sizes) {
      for (const [index, size] of sizes.rows) {
        if (size === null) delete worksheet.rowHeights[index];
        else worksheet.rowHeights[index] = size;
      }
      for (const [index, size] of sizes.cols) {
        if (size === null) delete worksheet.colWidths[index];
        else worksheet.colWidths[index] = size;
      }
      for (const [index, css] of sizes.colCss) setColumnCssWidth(worksheet, index, css);
    }
    const outline = this.outlineStateStore.get(sheetIndex);
    if (!outline) return;
    if (outline.rowCollapsed.size > 0) {
      for (const row of worksheet.rows) {
        const collapsed = outline.rowCollapsed.get(row.index);
        if (collapsed !== undefined) row.collapsed = collapsed;
      }
    }
    if (outline.colCollapsed.size > 0) {
      worksheet.colCollapsed = worksheet.colCollapsed ?? {};
      for (const [index, collapsed] of outline.colCollapsed) {
        if (collapsed) worksheet.colCollapsed[index] = true;
        else delete worksheet.colCollapsed[index];
      }
    }
  }

  /** Teardown: drop every sheet's edits, including the bound stashes. */
  destroy(): void {
    this.outlineStateStore.clear();
    this.stashedRowHeights.clear();
    this.stashedColWidths.clear();
    this.stashedColCssWidths.clear();
    this.sizeOverrideStore.clear();
  }
}

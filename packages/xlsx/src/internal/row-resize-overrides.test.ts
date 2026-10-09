import { describe, expect, it } from 'vitest';
import { GridGeometry } from './grid-geometry.js';
import { replaceRowResizeRanges, MAX_ROW_RESIZE_INTERVALS, type RowResizeRange } from './row-resize-overrides.js';
import { applySizeOverrides, createSizeOverriddenWorksheet } from '../worker-protocol.js';
import type { Worksheet } from '../types.js';

function worksheet(): Worksheet {
  return { name: 'Compact rows', rows: [], rowHeights: { 3: 0, 500000: 30 }, colWidths: {},
    defaultRowHeight: 15, defaultColWidth: 8.43, freezeRows: 1, freezeCols: 0 } as unknown as Worksheet;
}

describe('compact row resize projections', () => {
  it('matches independent dense geometry after overlapping interval edits at different zooms', () => {
    const source = worksheet();
    let ranges: readonly RowResizeRange[] = [];
    const dense: number[] = Array.from({ length: 64 }, (_, i) => i === 2 ? 0 : 20);
    for (let edit = 0; edit < 40; edit++) {
      const first = 1 + (edit * 7 % 55), last = Math.min(64, first + 8);
      const pixels = 5 + (edit * 13 % 100), height = pixels * 0.75;
      const targets = Array.from({ length: last - first + 1 }, (_, i) => first + i)
        .filter(index => index !== 3).map(index => ({ first: index, last: index }));
      ranges = replaceRowResizeRanges(ranges, targets, height);
      for (let row = first; row <= last; row++) if (row !== 3) dense[row - 1] = pixels;
      const projection = createSizeOverriddenWorksheet(source, { rowHeightRanges: ranges });
      const geometry = GridGeometry.forWorksheet(projection, 7);
      for (const scale of [0.65, 1.25]) {
        const axis = geometry.axesAtScale(scale).row;
        let offset = 0;
        for (let row = 1; row <= 64; row++) {
          expect(axis.offsetOf(row)).toBe(offset);
          expect(axis.sizeOf(row)).toBe(Math.round(dense[row - 1] * scale));
          offset += Math.round(dense[row - 1] * scale);
        }
        expect(axis.offsetOf(65)).toBe(offset);
      }
    }
  });
  it('uses per-band scale rounding and prefix offsets across a million-row interval', () => {
    const source = worksheet();
    const candidate = createSizeOverriddenWorksheet(source, {
      rowHeightRanges: [{ first: 2, last: 1048576, height: 63 }],
    });
    const geometry = GridGeometry.forWorksheet(candidate, 7);
    const scaled = geometry.axesAtScale(0.65).row;
    expect(scaled.sizeOf(1)).toBe(13);
    expect(scaled.sizeOf(500000)).toBe(55);
    expect(scaled.sizeOf(3)).toBe(0);
    expect(scaled.offsetOf(1048577)).toBe(13 + 1048574 * 55);
    expect(scaled.indexAt(scaled.offsetOf(1000000) + 12)).toEqual({ index: 1000000, partial: 12 });
    expect(scaled.bandsToCover(999999, 1048576, 100)).toEqual([
      { index: 999999, size: 55 }, { index: 1000000, size: 55 },
    ]);
    expect(GridGeometry.forWorksheet(source, 7).row.sizeOf(500000)).toBe(40);
    // A later outline hide/unhide remains authoritative inside a resize range.
    applySizeOverrides(candidate, { rows: { 500000: 0 } });
    expect(GridGeometry.forWorksheet(candidate, 7).row.sizeOf(500000)).toBe(0);
    applySizeOverrides(candidate, { rows: { 500000: null } });
    expect(GridGeometry.forWorksheet(candidate, 7).row.sizeOf(500000)).toBe(84);
  });

  it('coalesces same-height repeats and refuses fragmented candidates before changing the prior projection', () => {
    const targets = Array.from({ length: MAX_ROW_RESIZE_INTERVALS }, (_, i) => ({ first: 2 * i + 1, last: 2 * i + 1 }));
    const prior = replaceRowResizeRanges([], targets, 63);
    expect(() => replaceRowResizeRanges(prior, [{ first: 100000, last: 100000 }], 60)).toThrow(RangeError);
    expect(prior).toHaveLength(MAX_ROW_RESIZE_INTERVALS);
    expect(replaceRowResizeRanges(prior, [{ first: 1, last: 1048576 }], 63)).toEqual([
      { first: 1, last: 1048576, height: 63 },
    ]);
  });

  it('validates invalid wire intervals before applying any point edits', () => {
    const ws = worksheet(), before = { ...ws.rowHeights };
    expect(() => applySizeOverrides(ws, { rows: { 2: 90 },
      rowHeightRanges: [{ first: 1, last: 1048577, height: 63 }] })).toThrow(RangeError);
    expect(ws.rowHeights).toEqual(before);
    expect(GridGeometry.forWorksheet(ws, 7).row.sizeOf(2)).toBe(20);
  });
});

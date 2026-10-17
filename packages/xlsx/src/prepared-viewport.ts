import {
  collectChartImageFillUsagesForCharts,
  dropDecodedBitmapCache,
  dropSvgImageCache,
  OoxmlResourceLimitError,
} from '@silurus/ooxml-core';
import type { Worksheet, Styles, XlsxTextRunInfo } from './types.js';
import {
  getGridGeometryForWorksheet,
  HEADER_W,
  HEADER_H,
  rtlMirrorX,
  worksheetHasUncachedMath,
} from './renderer.js';
import { renderWorksheetViewport, worksheetWithAutoRowHeights } from './render-orchestrator.js';

export interface WorksheetMediaReference {
  readonly imagePath: string;
  readonly mimeType: string;
}

/** All model media, including oneCellAnchor/flattened group leaves and chart fills. */
export function worksheetMediaReferences(ws: Worksheet): readonly WorksheetMediaReference[] {
  const refs = new Map<string, WorksheetMediaReference>();
  const add = (ref: { imagePath: string; mimeType: string; svgImagePath?: string }): void => {
    refs.set(ref.imagePath, { imagePath: ref.imagePath, mimeType: ref.mimeType });
    if (ref.svgImagePath) refs.set(ref.svgImagePath, { imagePath: ref.svgImagePath, mimeType: 'image/svg+xml' });
  };
  for (const image of ws.images ?? []) add(image);
  for (const group of ws.shapeGroups ?? []) {
    for (const shape of group.shapes) if (shape.geom.type === 'image') add(shape.geom);
  }
  for (const usage of collectChartImageFillUsagesForCharts((ws.charts ?? []).map(anchor => anchor.chart))) add(usage.fill);
  return [...refs.values()];
}

export interface PreparedWorksheetViewportOptions {
  readonly width: number;
  readonly height: number;
  readonly dpr?: number;
  readonly scale?: number;
  /** CSS pixel scroll relative to the unfrozen area's origin. */
  readonly scrollX?: number;
  readonly scrollY?: number;
  readonly fetchImage?: (path: string, mimeType: string) => Promise<Blob>;
  readonly imageResources?: import('@silurus/ooxml-core').ImageResourceOptions;
  /** Maximum cells visited across visible scrollable and frozen bands, before paint. */
  readonly maxVisibleCells?: number;
  readonly onTextRun?: (info: XlsxTextRunInfo) => void;
}

export interface WorksheetViewportFrame {
  readonly width: number;
  readonly height: number;
  readonly dpr: number;
  readonly maximumDigitWidth: number;
  readonly projectedRowHeights: Readonly<Record<number, number>>;
  readonly range: import('./types.js').ViewportRange;
  /** 1-based logical cell identity; headers, hidden cells and outside points return null. */
  cellAt(x: number, y: number): { row: number; col: number } | null;
  /** CSS pixel rectangle, expanded to the entire authored merge. */
  cellRect(row: number, col: number): { x: number; y: number; w: number; h: number } | null;
}

/** Paint an already-admitted immutable worksheet. No archive, worker or network is opened.
 * Font registration is fixed for the lifetime of `ws`; a new layout generation needs a new model.
 * Optional engines are deliberately unsupported by this bounded facade. */
export async function renderPreparedWorksheetViewport(
  target: HTMLCanvasElement | OffscreenCanvas,
  ws: Worksheet,
  styles: Styles,
  options: PreparedWorksheetViewportOptions,
): Promise<WorksheetViewportFrame> {
  if (ws.parseError || ws.isDialogSheet || ws.isChartSheet) throw new Error('XLSX worksheet content is unsupported or incomplete');
  const classic = new Set(['clusteredBar', 'line', 'pie']);
  if ((ws.charts ?? []).some(({ chart }) => chart.threeD || !classic.has(chart.chartType)) || worksheetHasUncachedMath(ws)) {
    throw new Error('XLSX optional content renderer is unavailable');
  }
  const refs = worksheetMediaReferences(ws);
  if (refs.length && !options.fetchImage) throw new Error('XLSX worksheet media source is unavailable');
  if (refs.some(ref => !['image/png', 'image/jpeg', 'image/gif', 'image/webp'].includes(ref.mimeType))) {
    throw new Error('XLSX image codec is unsupported by this viewport');
  }
  const { width, height } = options;
  const scale = options.scale ?? 1;
  const dpr = options.dpr ?? 1;
  const scrollX = options.scrollX ?? 0;
  const scrollY = options.scrollY ?? 0;
  if (![width, height, scale, dpr].every(value => Number.isFinite(value) && value > 0)
    || ![scrollX, scrollY].every(value => Number.isFinite(value) && value >= 0)) throw new RangeError('invalid XLSX viewport');
  const ctx = target.getContext('2d') as CanvasRenderingContext2D | null;
  if (!ctx) throw new Error('XLSX render target does not provide a 2-D canvas context');
  // Resolve once using the same auto-fit projection that the painter receives.
  const projection = worksheetWithAutoRowHeights(ctx, ws, styles);
  const geometry = getGridGeometryForWorksheet(projection);
  const visible = geometry.visibleRange({ scale, scrollX, scrollY, width, height, headerWidth: HEADER_W, headerHeight: HEADER_H, buffer: 0 });
  if (options.maxVisibleCells !== undefined) {
    const limit = options.maxVisibleCells;
    if (!Number.isSafeInteger(limit) || limit < 1) throw new RangeError('invalid XLSX visible cell budget');
    const frozen = geometry.effectiveFrozenBands({ scale, width, height, headerWidth: HEADER_W, headerHeight: HEADER_H, rows: projection.freezeRows, cols: projection.freezeCols });
    const observed = (visible.range.rows + frozen.rows) * (visible.range.cols + frozen.cols);
    if (observed > limit) throw new OoxmlResourceLimitError('XLSX viewport cell limit exceeded', {
      stage: 'rendering',
      violation: { format: 'xlsx', operation: 'renderPreparedWorksheetViewport', resource: 'viewport-cells', metric: 'cell-count', limit, observed, configurable: true,
        usage: { archiveEntryCount: 0, declaredInflatedBytes: 0, distinctInflatedBytes: 0, operationInflatedBytes: 0 } },
    });
  }
  await renderWorksheetViewport({ ws: projection, styles }, target, visible.range, {
    width, height, dpr, cellScale: scale,
    scrollOffsetX: visible.offsetX, scrollOffsetY: visible.offsetY,
    freezeRows: projection.freezeRows, freezeCols: projection.freezeCols,
    fetchImage: options.fetchImage, imageResources: options.imageResources,
    onTextRun: options.onTextRun, strictImages: true,
  });
  const headerX = Math.round(HEADER_W * scale);
  const headerY = Math.round(HEADER_H * scale);
  const mergeFor = (row: number, col: number) => projection.mergeCells.find(merge => row >= merge.top && row <= merge.bottom && col >= merge.left && col <= merge.right);
  return {
    width, height, dpr: target.width / width,
    maximumDigitWidth: geometry.maximumDigitWidth,
    projectedRowHeights: { ...projection.rowHeights },
    range: visible.range,
    cellAt(x, y) {
      if (projection.rightToLeft) x = rtlMirrorX(x, 0, width);
      if (x < headerX || y < headerY || x >= width || y >= height) return null;
      const cell = geometry.cellAt(x - headerX, y - headerY, { scale, scrollX, scrollY });
      if (!cell) return null;
      const merge = mergeFor(cell.row, cell.col);
      return merge ? { row: merge.top, col: merge.left } : cell;
    },
    cellRect(row, col) {
      const merge = mergeFor(row, col);
      const rect = geometry.cellRect(merge?.top ?? row, merge?.left ?? col, { scale, scrollX, scrollY, headerWidth: HEADER_W, headerHeight: HEADER_H });
      if (!rect || rect.w === 0 || rect.h === 0) return null;
      if (merge) {
        const axes = geometry.axesAtScale(scale);
        rect.w = axes.col.offsetOf(merge.right + 1) - axes.col.offsetOf(merge.left);
        rect.h = axes.row.offsetOf(merge.bottom + 1) - axes.row.offsetOf(merge.top);
      }
      if (projection.rightToLeft) rect.x = rtlMirrorX(rect.x, rect.w, width);
      return rect;
    },
  };
}

/** Release the public caller's document-scoped decoded media, including GPU resources. */
export function releaseWorksheetViewportResources(fetchImage: (path: string, mimeType: string) => Promise<Blob>): void {
  dropDecodedBitmapCache(fetchImage);
  dropSvgImageCache(fetchImage);
}

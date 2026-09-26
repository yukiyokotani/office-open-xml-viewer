// Classic chart data table helpers.
import type { ChartModel, ChartSeries } from '../../types/chart';
import { chartDataTableFamilyIsPainted } from '../marker-style.js';
import { chartTextFontSizePx } from '../layout.js';
import { formatCategoryLabel, formatChartValWithCode, formatLocalizedExcelShortDate } from '../chart-number-format.js';
import { legendEntryGlobalIndex, legendEntryRanges } from '../legend-entry-plan.js';
import { elideToWidth } from '../text-elide.js';
import { axisLineWidthPx } from '../axis-style.js';
import { chartFontCss, chartFontFamily } from './fonts.js';
import { chartCategories } from '../category-spacing.js';
import { wrapMeasuredText } from './axis.js';
import { buildLegendEntries, drawLegendSwatch } from './legend.js';
import { dashPatternForPreset } from './geometry.js';


export type ChartDataTableLayout = {
  fontPx: number;
  lineHeight: number;
  headerLines: string[][];
  headerHeight: number;
  rowHeight: number;
  totalHeight: number;
};


/** Office only paints a classic chart data table for category-axis families.
 * CT_DTable is syntactically allowed under plotArea, but an authored table on
 * an XY scatter plot is ignored (confirmed with an Office vector boundary).
 * Keeping this gate beside the shared layout prevents family renderers from
 * inventing different applicability rules. */
export function chartHasDataTable(chart: ChartModel): boolean {
  return chart.dataTable != null && chartDataTableFamilyIsPainted(chart.chartType);
}


export function chartDataTableRows(chart: ChartModel): Array<{ series: ChartSeries; sourceIndex: number }> {
  const horizontal = chart.chartType === 'clusteredBarH'
    || chart.chartType === 'stackedBarH'
    || chart.chartType === 'stackedBarHPct';
  const rows = chart.series.map((series, sourceIndex) => ({ series, sourceIndex }));
  return horizontal ? rows.reverse() : rows;
}


/** Minimum data-table band reserved before the final plot width is known. The
 * header starts as one line; after `computeChartFrame` the measured category
 * cell width may add wrapped lines and the caller shrinks the plot by exactly
 * that measured delta. */
export function chartDataTableBaseHeight(chart: ChartModel, ptToPx: number): number {
  const table = chartHasDataTable(chart) ? chart.dataTable : null;
  if (!table) return 0;
  const fontPx = chartTextFontSizePx(table.fontSizeHpt, ptToPx) ?? 9 * ptToPx;
  const lineHeight = Math.max(1, fontPx * 1.2);
  const rowHeight = lineHeight + 4 * ptToPx;
  return (chart.series.length + 1) * rowHeight;
}


export function chartDataTableHeaderWidth(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  ptToPx: number,
): number {
  const table = chartHasDataTable(chart) ? chart.dataTable : null;
  if (!table) return 0;
  const fontPx = chartTextFontSizePx(table.fontSizeHpt, ptToPx) ?? 9 * ptToPx;
  const face = chartFontFamily(chart, table.fontFace, 'minor');
  ctx.save();
  ctx.font = chartFontCss(fontPx, face, table.fontBold ?? false, table.fontItalic ?? false);
  const nameWidth = chart.series.reduce(
    (width, series) => Math.max(width, ctx.measureText(series.name).width),
    0,
  );
  ctx.restore();
  const keyWidth = table.showKeys ? Math.max(12 * ptToPx, fontPx * 1.7) : 0;
  const keyGap = table.showKeys ? 4 * ptToPx : 0;
  return nameWidth + keyWidth + keyGap + 6 * ptToPx;
}


export function measureChartDataTable(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  categoryWidth: number,
  ptToPx: number,
): ChartDataTableLayout | null {
  const table = chartHasDataTable(chart) ? chart.dataTable : null;
  if (!table) return null;
  const fontPx = chartTextFontSizePx(table.fontSizeHpt, ptToPx) ?? 9 * ptToPx;
  const lineHeight = Math.max(1, fontPx * 1.2);
  const rowHeight = lineHeight + 4 * ptToPx;
  const face = chartFontFamily(chart, table.fontFace, 'minor');
  ctx.save();
  ctx.font = chartFontCss(fontPx, face, table.fontBold ?? false, table.fontItalic ?? false);
  const categoryFormat = chart.series.find(series => series.catFormatCode)?.catFormatCode
    ?? chart.catAxisFormatCode;
  const categoryBuiltinId = chart.series
    .find(series => series.catFormatBuiltinId != null)?.catFormatBuiltinId;
  const headerLines = chartCategories(chart).map(category => {
    const numeric = category.trim() === '' ? Number.NaN : Number(category);
    const label = categoryBuiltinId === 14 && Number.isFinite(numeric)
      ? formatLocalizedExcelShortDate(numeric, chart.date1904)
      : formatCategoryLabel(category, categoryFormat, chart.date1904);
    return wrapMeasuredText(ctx, label, Math.max(1, categoryWidth - 4 * ptToPx));
  });
  ctx.restore();
  const maxHeaderLines = Math.max(1, ...headerLines.map(lines => lines.length));
  const headerHeight = maxHeaderLines * lineHeight + 4 * ptToPx;
  return {
    fontPx,
    lineHeight,
    headerLines,
    headerHeight,
    rowHeight,
    totalHeight: headerHeight + chartDataTableRows(chart).length * rowHeight,
  };
}


/** Draw `CT_DTable` as a measured chart foreground band. Category columns are
 * aligned to the plot's category span; the leading key/name column occupies
 * the already-reserved value-axis gutter. Border switches are honored
 * independently, as authored by the four CT_Boolean children. */
export function drawChartDataTable(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  layout: ChartDataTableLayout | null,
  plotX: number,
  tableY: number,
  plotWidth: number,
  chartLeft: number,
  ptToPx: number,
): void {
  const table = chart.dataTable;
  if (!table || !layout) return;
  const categories = chartCategories(chart);
  if (categories.length === 0) return;
  const categoryWidth = plotWidth / categories.length;
  const face = chartFontFamily(chart, table.fontFace, 'minor');
  const font = chartFontCss(
    layout.fontPx, face, table.fontBold ?? false, table.fontItalic ?? false,
  );
  const keyWidth = table.showKeys ? Math.max(12 * ptToPx, layout.fontPx * 1.7) : 0;
  const keyGap = table.showKeys ? 4 * ptToPx : 0;
  ctx.save();
  ctx.font = font;
  const longestName = chart.series.reduce(
    (width, series) => Math.max(width, ctx.measureText(series.name).width),
    0,
  );
  const desiredHeaderWidth = longestName + keyWidth + keyGap + 6 * ptToPx;
  const headerWidth = Math.min(Math.max(0, plotX - chartLeft), desiredHeaderWidth);
  const tableX = plotX - headerWidth;
  const tableWidth = headerWidth + plotWidth;
  const tableBottom = tableY + layout.totalHeight;
  const tableRows = chartDataTableRows(chart);
  const keyEntries = buildLegendEntries(
    chart.series,
    chart.chartType,
    chart.scatterStyle,
    false,
    chart.categories,
    [],
    true,
    [],
    chart.radarStyle,
    chart,
  );
  const keyEntryRanges = legendEntryRanges(chart);
  // A direct solid dTable fill belongs to each generated body-text box. That
  // semantic is independent of the owning chart family, plot-group count,
  // manual plot layout, line wrapping, and sparse values. Unsupported fill
  // recipes remain present in the model but do not masquerade as a solid.
  const bodyFillColor = table.fillColor ?? null;
  ctx.beginPath();
  ctx.rect(tableX, tableY, tableWidth, layout.totalHeight);
  ctx.clip();
  const fontColor = table.fontHidden === true
    || (table.fontPaintAuthored === true && table.fontColor == null)
    ? 'transparent'
    : table.fontColor ? `#${table.fontColor}` : '#000000';
  const drawBodyText = (text: string, centerX: number, centerY: number): void => {
    // Desktop Excel scopes a direct dTable/spPr fill to the generated body
    // text boxes. It does not fill the table frame or the leading series-name
    // cells. The text layout box is the measured advance by the measured line
    // height, so this remains tied to authored typography rather than a cell-
    // or sample-specific inset.
    if (bodyFillColor && text !== '') {
      const width = ctx.measureText(text).width;
      ctx.fillStyle = `#${bodyFillColor}`;
      ctx.fillRect(
        centerX - width / 2,
        centerY - layout.lineHeight / 2,
        width,
        layout.lineHeight,
      );
    }
    ctx.fillStyle = fontColor;
    ctx.textAlign = 'center';
    ctx.fillText(text, centerX, centerY);
  };
  ctx.fillStyle = fontColor;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';

  for (let categoryIndex = 0; categoryIndex < categories.length; categoryIndex++) {
    const centerX = plotX + (categoryIndex + 0.5) * categoryWidth;
    const lines = layout.headerLines[categoryIndex] ?? [''];
    const textBlockHeight = lines.length * layout.lineHeight;
    const firstY = tableY + (layout.headerHeight - textBlockHeight) / 2 + layout.lineHeight / 2;
    lines.forEach((line, lineIndex) => {
      drawBodyText(line, centerX, firstY + lineIndex * layout.lineHeight);
    });
  }

  for (let seriesIndex = 0; seriesIndex < tableRows.length; seriesIndex++) {
    const { series, sourceIndex } = tableRows[seriesIndex];
    const rowTop = tableY + layout.headerHeight + seriesIndex * layout.rowHeight;
    const rowCenter = rowTop + layout.rowHeight / 2;
    if (headerWidth > 0) {
      const textLeft = tableX + 3 * ptToPx + keyWidth + keyGap;
      ctx.textAlign = 'left';
      ctx.fillText(
        elideToWidth(ctx, series.name, Math.max(0, plotX - textLeft - 2 * ptToPx)),
        textLeft,
        rowCenter,
      );
      if (table.showKeys && keyWidth > 0) {
        const keyX = tableX + 3 * ptToPx;
        const keyEntryIndex = legendEntryGlobalIndex(keyEntryRanges, sourceIndex);
        const entry = keyEntryIndex == null ? undefined : keyEntries[keyEntryIndex];
        if (entry) {
          const keyHeight = Math.min(layout.fontPx, layout.rowHeight - 2 * ptToPx);
          drawLegendSwatch(
            ctx,
            entry.swatchStyle,
            entry.color,
            keyX,
            rowCenter - keyHeight / 2,
            keyWidth,
            keyHeight,
            entry.marker,
            entry.fillPaint,
            entry.outlinePaint,
            entry.outlineColor,
            entry.outlineWidthEmu,
            entry.outlineDash,
            entry.outlineCustomDash,
            entry.outlineCap,
            entry.outlineJoin,
            ptToPx,
            0,
            entry.directEffect,
            entry.fallbackEffect,
            entry.directEffectIndex,
            entry.fallbackEffectIndex,
          );
        }
        ctx.fillStyle = fontColor;
      }
    }
    for (let categoryIndex = 0; categoryIndex < categories.length; categoryIndex++) {
      const value = series.values[categoryIndex];
      const text = value == null ? '' : formatChartValWithCode(value, series.valFormatCode);
      drawBodyText(text, plotX + (categoryIndex + 0.5) * categoryWidth, rowCenter);
    }
  }

  if (table.lineHidden !== true
    && (table.linePaintAuthored !== true || table.lineColor != null)) {
    ctx.strokeStyle = table.lineColor ? `#${table.lineColor}` : '#808080';
    ctx.lineWidth = table.lineWidthEmu != null
      ? axisLineWidthPx(table.lineWidthEmu, ptToPx)
      : Math.max(0.5, ptToPx * 0.75);
    ctx.setLineDash(dashPatternForPreset(table.lineDash ?? undefined, ctx.lineWidth));
    if (table.showHorizontalBorder) {
      let lineY = tableY + layout.headerHeight;
      for (let row = 0; row < tableRows.length; row++) {
        ctx.beginPath(); ctx.moveTo(tableX, lineY); ctx.lineTo(tableX + tableWidth, lineY); ctx.stroke();
        lineY += layout.rowHeight;
      }
    }
    if (table.showVerticalBorder) {
      ctx.beginPath(); ctx.moveTo(plotX, tableY); ctx.lineTo(plotX, tableBottom); ctx.stroke();
      for (let category = 1; category < categories.length; category++) {
        const lineX = plotX + category * categoryWidth;
        ctx.beginPath(); ctx.moveTo(lineX, tableY); ctx.lineTo(lineX, tableBottom); ctx.stroke();
      }
    }
    if (table.showOutline) {
      const half = ctx.lineWidth / 2;
      ctx.strokeRect(
        tableX + half, tableY + half,
        Math.max(0, tableWidth - ctx.lineWidth),
        Math.max(0, layout.totalHeight - ctx.lineWidth),
      );
    }
  }
  ctx.restore();
}

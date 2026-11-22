// Classic chart data table helpers.
import type { ChartModel, ChartSeries } from '../../types/chart';
import { chartDataTableFamilyIsPainted } from '../marker-style.js';
import { catAxisLabelBandH, chartTextFontSizePx } from '../layout.js';
import { formatCategoryLabel, formatChartValWithCode, formatLocalizedExcelShortDate } from '../chart-number-format.js';
import { legendEntryGlobalIndex, legendEntryRanges } from '../legend-entry-plan.js';
import { authoredLineCount, authoredTextLines, fillAuthoredText, measureAuthoredTextWidth } from '../chart-text-lines.js';
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
  /** Height of a single-line body row. */
  rowHeight: number;
  /** Extra height per body row (in paint order) for authored line breaks in
   *  its series name or value text; 0 for single-line rows. */
  rowExtraHeights: number[];
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

/** ECMA-376 §21.2.2.54 defines CT_DTable but not its automatic geometry.
 * Observed desktop Excel behavior (column/line/area/stock, day/month date axes,
 * with/without a secondary value axis): a date table retains one column per
 * source category, independently of calendar ticks, and keeps the axis labels.
 * A text-axis table instead replaces those labels and shares the plot span.
 * Horizontal bars retain their separate value-axis band in the family renderer.
 */
export function chartDataTableSharesCategoryAxis(chart: ChartModel): boolean {
  return chart.catAxisIsDate !== true
    || chart.chartType === 'clusteredBarH'
    || chart.chartType === 'stackedBarH'
    || chart.chartType === 'stackedBarHPct';
}

export function chartDataTableGap(chart: ChartModel, categoryFontPx: number): number {
  if (!chartHasDataTable(chart) || chartDataTableSharesCategoryAxis(chart)) return 0;
  // Excel also separates the table when date tick labels are hidden. Library
  // spacing policy uses one font-height there; visible labels use our normal axis
  // band (including the authored lblOffset), rather than a sample-sized gap.
  return chart.catAxisHidden || chart.catAxisTickLabelPos === 'none'
    || chart.catAxisTickLabelPos === 'high'
    ? categoryFontPx
    : catAxisLabelBandH(categoryFontPx, chart.catAxisLabelOffsetPercent,
      chart.cartesianAutoLayoutProfile);
}

export function chartDataTableCategoryWidth(
  chart: ChartModel,
  plotWidth: number,
  secondaryRightBand: number,
): number {
  // Excel's automatic date table extends beneath the secondary labels/title.
  // Keep the chart/legend outer margin intact. Authored manual plot layouts
  // have not been validated for that extension: retain their explicit span.
  return !chartDataTableSharesCategoryAxis(chart) && !chart.plotAreaManualLayout
    ? plotWidth + secondaryRightBand
    : plotWidth;
}


function chartDataTableRows(chart: ChartModel): Array<{ series: ChartSeries; sourceIndex: number }> {
  const horizontal = chart.chartType === 'clusteredBarH'
    || chart.chartType === 'stackedBarH'
    || chart.chartType === 'stackedBarHPct';
  const rows = chart.series.map((series, sourceIndex) => ({ series, sourceIndex }));
  return horizontal ? rows.reverse() : rows;
}


function dataTableValueText(chart: ChartModel, series: ChartSeries, categoryIndex: number): string {
  const value = series.values[categoryIndex];
  return value == null ? '' : formatChartValWithCode(value, series.valFormatCode, chart.date1904);
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
    (width, series) => Math.max(width, measureAuthoredTextWidth(ctx, series.name)),
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
  const seriesCategoryFormat = chart.series.find(series => series.catFormatCode)?.catFormatCode;
  const categoryFormat = seriesCategoryFormat ?? chart.catAxisFormatCode;
  const categoryDivisors = seriesCategoryFormat == null ? chart.catAxisFormatDivisors : null;
  const categoryBuiltinId = chart.series
    .find(series => series.catFormatBuiltinId != null)?.catFormatBuiltinId;
  const headerLines = chartCategories(chart).map(category => {
    const numeric = category.trim() === '' ? Number.NaN : Number(category);
    const label = categoryBuiltinId === 14 && Number.isFinite(numeric)
      ? formatLocalizedExcelShortDate(numeric, chart.date1904)
      : formatCategoryLabel(category, categoryFormat, chart.date1904, categoryDivisors);
    return wrapMeasuredText(ctx, label, Math.max(1, categoryWidth - 4 * ptToPx));
  });
  ctx.restore();
  const maxHeaderLines = Math.max(1, ...headerLines.map(lines => lines.length));
  const headerHeight = maxHeaderLines * lineHeight + 4 * ptToPx;
  // A series name or value with authored breaks makes its row a block of lines.
  const tableRows = chartDataTableRows(chart);
  const categoryCount = chartCategories(chart).length;
  const rowExtraHeights = tableRows.map(({ series }) => {
    let lines = authoredLineCount(series.name);
    for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++) {
      lines = Math.max(lines, authoredLineCount(dataTableValueText(chart, series, categoryIndex)));
    }
    return (lines - 1) * lineHeight;
  });
  const rowExtraTotal = rowExtraHeights.reduce((sum, extra) => sum + extra, 0);
  return {
    fontPx,
    lineHeight,
    headerLines,
    headerHeight,
    rowHeight,
    rowExtraHeights,
    totalHeight: headerHeight + tableRows.length * rowHeight + rowExtraTotal,
  };
}


/** Draw `CT_DTable` as a measured chart foreground band. Category columns use
 * the resolved table span; the leading key/name column occupies
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
    (width, series) => Math.max(width, measureAuthoredTextWidth(ctx, series.name)),
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
    const lines = authoredTextLines(text);
    if (lines.length > 1) {
      // Authored breaks: one text box per line, centred as a block.
      const firstY = centerY - (lines.length - 1) * layout.lineHeight / 2;
      lines.forEach((line, index) => drawBodyLine(line, centerX, firstY + index * layout.lineHeight));
      return;
    }
    // Outer breaks are trimmed (e.g. a `0"\n"` format yields "10\n").
    drawBodyLine(lines[0], centerX, centerY);
  };
  const drawBodyLine = (text: string, centerX: number, centerY: number): void => {
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

  let rowExtraBefore = 0;
  for (let seriesIndex = 0; seriesIndex < tableRows.length; seriesIndex++) {
    const { series, sourceIndex } = tableRows[seriesIndex];
    const rowExtra = layout.rowExtraHeights[seriesIndex] ?? 0;
    const rowTop = tableY + layout.headerHeight + seriesIndex * layout.rowHeight + rowExtraBefore;
    const rowCenter = rowTop + (layout.rowHeight + rowExtra) / 2;
    rowExtraBefore += rowExtra;
    if (headerWidth > 0) {
      const textLeft = tableX + 3 * ptToPx + keyWidth + keyGap;
      ctx.textAlign = 'left';
      fillAuthoredText(
        ctx,
        series.name,
        textLeft,
        rowCenter,
        layout.lineHeight,
        Math.max(0, plotX - textLeft - 2 * ptToPx),
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
      drawBodyText(
        dataTableValueText(chart, series, categoryIndex),
        plotX + (categoryIndex + 0.5) * categoryWidth,
        rowCenter,
      );
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
        lineY += layout.rowHeight + (layout.rowExtraHeights[row] ?? 0);
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

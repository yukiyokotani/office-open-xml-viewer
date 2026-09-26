// Classic chart axis helpers.
import type { ChartDisplayUnits, ChartLabelBox, ChartManualLayout, ChartModel, ChartRect, SecondaryValueAxis } from '../../types/chart';
import { isCrossBetween, resolveGridline } from '../axis-style.js';
import { formatChartVal, formatChartValWithCode } from '../chart-number-format.js';
import { axisTitleMargin, axisTitleRotationRad, chartTextFontSizePx, resolveManualLayoutRect } from '../layout.js';
import type { ChartAxisTitleSide } from '../layout.js';
import { elideToWidth } from '../text-elide.js';
import type { ChartExStyle } from './chartex-style.js';
import { paintChartLabelBox } from '../label-box.js';
import { rawLinkedChartStyleRole } from '../effective-style.js';
import { automaticPercentMajorUnit, planNumericValueAxis } from '../axis-scale.js';
import { dashPatternForPreset } from './geometry.js';
import { chartFontCss, chartFontFamily } from './fonts.js';
import { effectiveLinkedLabelBox } from './style-roles.js';


export function drawAxisTick(
  ctx: CanvasRenderingContext2D,
  mode: string | null | undefined,
  axis: 'val' | 'cat',
  anchorXOrY: number,
  perpendicular: number,
  color?: string,
  lineWidth?: number,
  // For a vertical value axis "outside" is to the LEFT (the axis sits on the
  // left). A secondary value axis sits on the RIGHT, where "outside" points
  // right — pass `opposite` to flip the out/in direction.
  opposite = false,
  lineHidden = false,
  level: 'major' | 'minor' = 'major',
  ptToPx = 1,
  dash?: string | null,
): void {
  // Axis shape properties style both the rule and its tick marks. An authored
  // `<a:ln><a:noFill/>` therefore suppresses the ticks too, while labels and
  // gridlines remain independently visible.
  if (lineHidden || mode === 'none' || !mode) return;
  // Office's vector output uses 6pt major ticks and 4pt minor ticks. Tick
  // length still scales mildly with an unusually thick authored axis rule.
  const len = axisTickLengthPx(level, lineWidth, ptToPx);
  // Office's 6pt/4pt observation is the complete cross-tick length, not the
  // length on each side of the axis. out/in use the full length on one side;
  // cross splits it evenly around the rule.
  const sideLen = mode === 'cross' ? len / 2 : len;
  const prevS = ctx.strokeStyle;
  const prevW = ctx.lineWidth;
  const prevDash = ctx.getLineDash?.() ?? [];
  ctx.strokeStyle = color ?? '#888';
  ctx.lineWidth = lineWidth ?? 1;
  ctx.setLineDash(dashPatternForPreset(dash ?? undefined, ctx.lineWidth));
  ctx.beginPath();
  if (axis === 'val') {
    // val axis is vertical (x = anchor, y varies). Ticks extend horizontally;
    // `outSign` points away from the plot (left for a left axis, right for a
    // right/secondary axis).
    const x0 = anchorXOrY;
    const y = perpendicular;
    const outSign = opposite ? 1 : -1;
    const outer = mode === 'out' || mode === 'cross' ? outSign * sideLen : 0;
    const inner = mode === 'in' || mode === 'cross' ? -outSign * sideLen : 0;
    ctx.moveTo(x0 + outer, y);
    ctx.lineTo(x0 + inner, y);
  } else {
    // cat axis is horizontal (y = anchor, x varies). Ticks extend vertically.
    const y0 = anchorXOrY;
    const xc = perpendicular;
    const outSign = opposite ? -1 : 1;
    const outer = mode === 'out' || mode === 'cross' ? outSign * sideLen : 0;
    const inner = mode === 'in' || mode === 'cross' ? -outSign * sideLen : 0;
    ctx.moveTo(xc, y0 + outer);
    ctx.lineTo(xc, y0 + inner);
  }
  ctx.stroke();
  ctx.strokeStyle = prevS;
  ctx.lineWidth = prevW;
  ctx.setLineDash(prevDash);
}


export function strokeAxisSegment(
  ctx: CanvasRenderingContext2D,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  color: string,
  lineWidth: number,
  dash?: string | null,
): void {
  const previousDash = ctx.getLineDash?.() ?? [];
  const resolvedDash = dashPatternForPreset(dash ?? undefined, lineWidth);
  const dashChanged = resolvedDash.length !== previousDash.length
    || resolvedDash.some((value, index) => value !== previousDash[index]);
  ctx.strokeStyle = color;
  ctx.lineWidth = lineWidth;
  if (dashChanged) ctx.setLineDash(resolvedDash);
  ctx.beginPath();
  ctx.moveTo(x1, y1);
  ctx.lineTo(x2, y2);
  ctx.stroke();
  if (dashChanged) ctx.setLineDash(previousDash);
}


export function axisTickLengthPx(
  level: 'major' | 'minor',
  lineWidth: number | undefined,
  ptToPx: number,
): number {
  const baseLen = (level === 'minor' ? 4 : 6) * ptToPx;
  return lineWidth ? Math.max(baseLen, lineWidth + 2 * ptToPx) : baseLen;
}


/** Distance an axis tick occupies outside the plot-side axis rule. */
export function axisTickOutwardExtentPx(
  mode: string | null | undefined,
  level: 'major' | 'minor',
  lineWidth: number | undefined,
  ptToPx: number,
): number {
  if (mode !== 'out' && mode !== 'cross') return 0;
  const length = axisTickLengthPx(level, lineWidth, ptToPx);
  return mode === 'cross' ? length / 2 : length;
}


/** Stroke one horizontal value-axis gridline spanning the plot width at `gy`.
 *  Extracted from the identical stroke the column-bar, line and area renderers
 *  each emitted inline. `isZero` is the caller's "this is the value-0 line"
 *  predicate (`si === 0` / `v === 0`). Callers set their own font/label
 *  BEFORE/AFTER this call, which is why those (drifted) parts stay at the call
 *  sites. Scatter is deliberately NOT a caller — it has no baseline special-case.
 *
 *  `grid` is the resolved `{ color, width }` from `resolveGridline` (the file's
 *  `<c:majorGridlines><c:spPr><a:ln>` or the faint `#e0e0e0`/0.5 px default).
 *  When the file supplies NO explicit gridline color (`grid.explicit === false`)
 *  the historical baseline emphasis applies: the value-0 line is a darker
 *  `#aaa` 1 px rule. When the file DOES pin a gridline color, PowerPoint strokes
 *  every major gridline in that one color/width uniformly, so the zero-line
 *  override is suppressed. Omitting `grid` reproduces the pre-CH-gridline
 *  default exactly (byte-stable for callers that haven't resolved a style). */
export function strokeValueGridlineH(
  ctx: CanvasRenderingContext2D,
  px0: number,
  pw: number,
  gy: number,
  isZero: boolean,
  grid?: { color: string; width: number; explicit: boolean; dash: number[] },
): void {
  if (grid && grid.explicit) {
    ctx.strokeStyle = grid.color;
    ctx.lineWidth = grid.width;
  } else {
    ctx.strokeStyle = isZero ? '#aaa' : grid?.color ?? '#e0e0e0';
    ctx.lineWidth = isZero ? 1 : grid?.width ?? 0.5;
  }
  const authoredDash = grid?.dash ?? [];
  const previousDash = authoredDash.length > 0 && ctx.getLineDash ? ctx.getLineDash() : [];
  if (authoredDash.length > 0) ctx.setLineDash(authoredDash);
  ctx.beginPath();
  ctx.moveTo(px0, gy);
  ctx.lineTo(px0 + pw, gy);
  ctx.stroke();
  if (authoredDash.length > 0) ctx.setLineDash(previousDash);
}


/** Resolve the value-axis MAJOR gridline stroke for `chart` at the current
 *  display scale. `explicit` is true when the file pinned any line property
 *  (color, width, or dash) under `<c:valAx><c:majorGridlines><c:spPr><a:ln>`;
 *  that flag tells
 *  `strokeValueGridlineH` to stroke every gridline in the resolved color
 *  uniformly (no `#aaa` zero-line emphasis), matching PowerPoint. With no
 *  explicit color the resolved `{ color: '#e0e0e0', width: 0.5 }` reproduces the
 *  historical faint hairline (byte-stable). */
export function valGridStroke(
  chart: ChartModel,
  ptToPx: number,
): { color: string; width: number; explicit: boolean; dash: number[] } {
  const { color, width } = resolveGridline(chart.valAxisGridlineColor, chart.valAxisGridlineWidthEmu, ptToPx);
  return {
    color,
    width,
    explicit: chart.valAxisGridlineColor != null
      || chart.valAxisGridlineWidthEmu != null
      || chart.valAxisGridlineDash != null,
    dash: dashPatternForPreset(chart.valAxisGridlineDash ?? undefined, width),
  };
}


export function valMinorGridStroke(
  chart: ChartModel,
  ptToPx: number,
): { color: string; width: number; explicit: boolean; dash: number[] } {
  const { color, width } = resolveGridline(
    chart.valAxisMinorGridlineColor,
    chart.valAxisMinorGridlineWidthEmu,
    ptToPx,
  );
  return {
    color,
    width,
    explicit: chart.valAxisMinorGridlineColor != null,
    dash: dashPatternForPreset(chart.valAxisMinorGridlineDash ?? undefined, width),
  };
}


export function secondaryMinorGridStroke(
  axis: SecondaryValueAxis,
  ptToPx: number,
): { color: string; width: number; explicit: boolean; dash: number[] } {
  const { color, width } = resolveGridline(
    axis.minorGridlineColor,
    axis.minorGridlineWidthEmu,
    ptToPx,
  );
  return {
    color,
    width,
    explicit: axis.minorGridlineColor != null
      || axis.minorGridlineWidthEmu != null
      || axis.minorGridlineDash != null,
    dash: dashPatternForPreset(axis.minorGridlineDash ?? undefined, width),
  };
}


export function secondaryMajorGridStroke(
  axis: SecondaryValueAxis,
  ptToPx: number,
): { color: string; width: number; explicit: boolean; dash: number[] } {
  const { color, width } = resolveGridline(
    axis.majorGridlineColor,
    axis.majorGridlineWidthEmu,
    ptToPx,
  );
  return {
    color,
    width,
    explicit: axis.majorGridlineColor != null
      || axis.majorGridlineWidthEmu != null
      || axis.majorGridlineDash != null,
    dash: dashPatternForPreset(axis.majorGridlineDash ?? undefined, width),
  };
}


/** Whether to draw CATEGORY-axis MAJOR gridlines (`<c:catAx><c:majorGridlines>`,
 *  ECMA-376 §21.2.2.100). Office omits them by default, so only `true` turns
 *  them on (null/undefined/false ⇒ off, byte-stable). */
export function drawCatMajorGridlines(chart: ChartModel): boolean {
  return chart.catAxisMajorGridlines === true;
}


/** Resolve the CATEGORY-axis major gridline stroke, mirroring
 *  {@link valGridStroke}. `<c:catAx><c:majorGridlines><c:spPr><a:ln>` gives the
 *  color/width (`chart.catAxisGridlineColor`/`catAxisGridlineWidthEmu`); absent
 *  ⇒ the same faint `#e0e0e0`/0.5 px default as the value axis. Category
 *  gridlines have no zero-line emphasis (there is no "zero category"), so a
 *  single resolved stroke suffices. */
export function catGridStroke(chart: ChartModel, ptToPx: number): { color: string; width: number; dash: number[] } {
  const stroke = resolveGridline(chart.catAxisGridlineColor, chart.catAxisGridlineWidthEmu, ptToPx);
  return {
    ...stroke,
    dash: dashPatternForPreset(chart.catAxisGridlineDash ?? undefined, stroke.width),
  };
}


export function catMinorGridStroke(chart: ChartModel, ptToPx: number): { color: string; width: number; dash: number[] } {
  const stroke = resolveGridline(
    chart.catAxisMinorGridlineColor,
    chart.catAxisMinorGridlineWidthEmu,
    ptToPx,
  );
  return {
    ...stroke,
    dash: dashPatternForPreset(chart.catAxisMinorGridlineDash ?? undefined, stroke.width),
  };
}


/** The plot-fraction positions (0..1 across the category extent) of the CATEGORY
 *  major gridlines / ticks for `n` categories. With crossBetween="between" (the
 *  bar/column default) they sit on the `n+1` band BOUNDARIES; under "midCat"
 *  they sit at the `n` category CENTERS. Shared by the category tick loop and
 *  the category-gridline pass so both stay aligned (§21.2.2.100/§21.2.2.32). */
export function catGridlineFractions(chart: ChartModel, n: number): number[] {
  if (n <= 0) return [];
  const onBoundary = isCrossBetween(chart);
  const fracs: number[] = [];
  const last = onBoundary ? n : n - 1;
  for (let ci = 0; ci <= last; ci++) {
    fracs.push(onBoundary ? ci / n : (n === 1 ? 0.5 : ci / (n - 1)));
  }
  return fracs;
}


/** True when the value axis is reversed (`<c:valAx><c:scaling><c:orientation
 *  val="maxMin">`, ECMA-376 §21.2.2.130). Absent/"minMax" ⇒ false (byte-stable). */
export function valAxisReversed(chart: ChartModel): boolean {
  return chart.valAxisOrientation === 'maxMin';
}


/** True when the category axis is reversed (`<c:catAx>…orientation="maxMin">`). */
export function catAxisReversed(chart: ChartModel): boolean {
  return chart.catAxisOrientation === 'maxMin';
}


/** Whether to draw value-axis MAJOR gridlines. Office writes `<c:majorGridlines>`
 *  on the value axis by default, so the historical always-on behavior maps to
 *  "draw unless the model explicitly says the element is absent". `undefined`
 *  (parser didn't model it) ⇒ true (byte-stable); `false` (axis present without
 *  the element) ⇒ off. */
export function drawValMajorGridlines(chart: ChartModel): boolean {
  return chart.valAxisMajorGridlines !== false;
}


/** A resolved value-axis plan: rounded bounds, the major gridline VALUES to
 *  stroke, an optional minor gridline VALUES list, and the value→fraction map
 *  (0 at the axis min end, 1 at the max end — before any pixel flip). Centralizes
 *  the CH6 major unit / logBase / orientation handling so every value-axis
 *  family shares one spec-faithful code path. With no CH6 fields set the plan is
 *  byte-identical to the old inline math: `step`/bounds from `valueAxisScale`,
 *  `majorLines = [min, min+step, … max]`, `frac(v) = (v-min)/(max-min)`. */
export interface ValueAxisPlan {
  min: number;
  max: number;
  step: number;
  majorLines: number[];
  minorLines: number[];
  minorTicks: number[];
  /** 0..1 position of `v` from the axis minimum toward the maximum (log-aware,
   *  orientation-aware). Renderers turn this into a pixel with
   *  `plotBottom - frac(v) * plotHeight` (vertical) — the reversal is already
   *  baked in, so callers keep their existing `- frac*len` form. */
  frac: (v: number) => number;
}


/** Convert an OOXML percent-axis value (stored as a 0..1 ratio) into the
 * renderer's percentStacked geometry space (0..100 percentage points). */
export function valueAxisUnitInRendererSpace(
  value: number | null | undefined,
  percentStacked: boolean,
): number | null | undefined {
  return value == null || !percentStacked ? value : value * 100;
}


/** Format a primary value-axis tick from the renderer's data space. For a
 * percentStacked chart the plotted values are percentage points, while the
 * axis numFmt still expects the OOXML ratio (0.5 → 50%). */
export function formatPrimaryValueAxisTick(
  chart: ChartModel,
  value: number,
  percentStacked: boolean,
): string {
  return formatChartValWithCode(
    (percentStacked ? value / 100 : value) / displayUnitDivisor(chart.valAxisDisplayUnits),
    percentStacked ? (chart.valAxisFormatCode ?? '0%') : chart.valAxisFormatCode,
    chart.date1904,
  );
}


export function displayUnitDivisor(units: ChartDisplayUnits | null | undefined): number {
  const divisor = units?.divisor;
  return divisor != null && Number.isFinite(divisor) && divisor > 0 ? divisor : 1;
}


export function formatAxisTickWithUnits(
  value: number,
  formatCode: string | null | undefined,
  date1904: boolean | undefined,
  units: ChartDisplayUnits | null | undefined,
): string {
  return formatChartValWithCode(value / displayUnitDivisor(units), formatCode, date1904);
}


export function automaticDisplayUnitLabel(units: ChartDisplayUnits): string {
  const names: Record<string, string> = {
    hundreds: 'Hundreds',
    thousands: 'Thousands',
    tenThousands: 'Ten Thousands',
    hundredThousands: 'Hundred Thousands',
    millions: 'Millions',
    tenMillions: 'Ten Millions',
    hundredMillions: 'Hundred Millions',
    billions: 'Billions',
    trillions: 'Trillions',
  };
  return units.builtInUnit ? (names[units.builtInUnit] ?? units.builtInUnit) : formatChartVal(units.divisor);
}


/** Paint the optional §21.2.2.46 display-unit labels after the family painter.
 * Their manual layout is chart-space (not plot-space), so this shared overlay
 * avoids a separate approximation in every chart family. */
export function drawChartDisplayUnitLabels(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  rect: ChartRect,
  ptToPx: number,
): void {
  const entries = [
    { units: chart.valAxisDisplayUnits, vertical: true, fallbackX: rect.x + rect.w * 0.08, fallbackY: rect.y + rect.h * 0.12, axis: { size: chart.valAxisFontSizeHpt, bold: chart.valAxisFontBold, italic: chart.valAxisFontItalic, color: chart.valAxisFontColor, paintAuthored: chart.valAxisFontPaintAuthored, face: chart.valAxisFontFace } },
    { units: chart.catAxisDisplayUnits, vertical: false, fallbackX: rect.x + rect.w * 0.82, fallbackY: rect.y + rect.h * 0.82, axis: { size: chart.catAxisFontSizeHpt, bold: chart.catAxisFontBold, italic: chart.catAxisFontItalic, color: chart.catAxisFontColor, paintAuthored: chart.catAxisFontPaintAuthored, face: chart.catAxisFontFace } },
    { units: chart.secondaryValAxis?.displayUnits, vertical: true, fallbackX: rect.x + rect.w * 0.92, fallbackY: rect.y + rect.h * 0.12, axis: { size: chart.secondaryValAxis?.fontSizeHpt, bold: chart.secondaryValAxis?.fontBold, italic: chart.secondaryValAxis?.fontItalic, color: chart.secondaryValAxis?.fontColor, paintAuthored: chart.secondaryValAxis?.fontPaintAuthored, face: chart.secondaryValAxis?.fontFace } },
    { units: chart.secondaryCatAxis?.displayUnits, vertical: false, fallbackX: rect.x + rect.w * 0.82, fallbackY: rect.y + rect.h * 0.08, axis: { size: chart.secondaryCatAxis?.fontSizeHpt, bold: chart.secondaryCatAxis?.fontBold, italic: chart.secondaryCatAxis?.fontItalic, color: chart.secondaryCatAxis?.fontColor, paintAuthored: chart.secondaryCatAxis?.fontPaintAuthored, face: chart.secondaryCatAxis?.fontFace } },
  ];
  for (const { units, vertical, fallbackX, fallbackY, axis } of entries) {
    const label = units?.label;
    if (!units || !label) continue;
    const text = label.text ?? automaticDisplayUnitLabel(units);
    const chartText = chart.chartTextStyle;
    const fontPx = chartTextFontSizePx(
      label.fontSizeHpt ?? axis.size ?? chartText?.fontSizeHpt,
      ptToPx,
    ) ?? 10 * ptToPx;
    const fontBold = label.fontBold ?? axis.bold ?? chartText?.fontBold ?? false;
    const fontItalic = label.fontItalic ?? axis.italic ?? chartText?.fontItalic ?? false;
    const fontPaint = label.fontPaintAuthored === true
      ? { color: label.fontColor, hidden: label.fontHidden === true || label.fontColor == null }
      : axis.paintAuthored === true
        ? { color: axis.color, hidden: axis.color == null }
        : chartText?.fontPaintAuthored === true
          ? { color: chartText.fontColor, hidden: chartText.fontColor == null }
          : { color: label.fontColor ?? axis.color ?? chartText?.fontColor, hidden: false };
    if (fontPaint.hidden) continue;
    const fontColor = fontPaint.color;
    const fontFace = label.fontFace ?? axis.face ?? chartText?.fontFace;
    ctx.save();
    ctx.font = chartFontCss(
      fontPx,
      chartFontFamily(chart, fontFace, 'minor'),
      fontBold,
      fontItalic,
    );
    const rotation = label.rotation != null
      ? (label.rotation / 60_000) * Math.PI / 180
      : vertical ? -Math.PI / 2 : 0;
    const textWidth = ctx.measureText(text).width;
    const rotatedW = Math.abs(Math.cos(rotation)) * textWidth + Math.abs(Math.sin(rotation)) * fontPx;
    const rotatedH = Math.abs(Math.sin(rotation)) * textWidth + Math.abs(Math.cos(rotation)) * fontPx;
    const automatic = {
      x: fallbackX - rotatedW / 2,
      y: fallbackY - rotatedH / 2,
      w: rotatedW,
      h: rotatedH,
    };
    const positioned = label.manualLayout
      ? resolveManualLayoutRect(
          { ...label.manualLayout, w: undefined, h: undefined },
          rect,
          automatic,
        )
      : automatic;
    if (!positioned) { ctx.restore(); continue; }
    const cx = positioned.x + positioned.w / 2;
    const cy = positioned.y + positioned.h / 2;
    paintChartLabelBox(
      ctx,
      effectiveLinkedLabelBox(
        chart,
        label.boxStyle,
        chart.chartStyleRoles?.axisTitle,
        rawLinkedChartStyleRole(chart, 'axisTitle'),
        true,
      ),
      positioned,
      ptToPx,
    );
    ctx.translate(cx, cy);
    if (rotation !== 0) ctx.rotate(rotation);
    ctx.fillStyle = fontColor ? `#${fontColor}` : '#595959';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(text, 0, 0);
    ctx.restore();
  }
}


/** Build a {@link ValueAxisPlan} for the primary value axis. `dataMin`/`dataMax`
 *  are the raw data extents already massaged by the caller (0-anchoring, pct
 *  normalization, explicit valMin/valMax). `axisLenPt` drives the auto major
 *  unit. Reversal is read from the chart's value-axis orientation. */
export function planValueAxis(
  chart: ChartModel,
  dataMin: number,
  dataMax: number,
  axisLenPt?: number,
  percentStacked = false,
  axisOrientation: 'vertical' | 'horizontal' = 'vertical',
): ValueAxisPlan {
  const reversed = valAxisReversed(chart);
  const logBase = chart.valAxisLogBase;
  // c:valAx values remain ratios for percentStacked charts, but all plotted
  // geometry in this renderer is expressed as percentage points. Explicit
  // bounds/units therefore cross the same ×100 boundary as the series values.
  // With no explicit bounds, percentStacked uses its exact normalized extent
  // (0..100 or -100..100) instead of adding ordinary numeric-axis headroom.
  const explicitMin = valueAxisUnitInRendererSpace(chart.valMin, percentStacked)
    ?? (percentStacked ? dataMin : chart.valMin);
  const explicitMax = valueAxisUnitInRendererSpace(chart.valMax, percentStacked)
    ?? (percentStacked ? dataMax : chart.valMax);
  const authoredMajorUnit = valueAxisUnitInRendererSpace(
    chart.valAxisMajorUnit,
    percentStacked,
  );
  const majorUnit = percentStacked
    && !(logBase != null && isFinite(logBase) && logBase >= 2)
    && !(authoredMajorUnit != null && isFinite(authoredMajorUnit) && authoredMajorUnit > 0)
      ? automaticPercentMajorUnit(dataMin, dataMax, axisOrientation, axisLenPt)
      : authoredMajorUnit;
  const needsMinorTicks = chart.valAxisMinorTickMark != null
    && chart.valAxisMinorTickMark !== 'none';
  const mu = valueAxisUnitInRendererSpace(chart.valAxisMinorUnit, percentStacked);
  const numeric = planNumericValueAxis({
    dataMin,
    dataMax,
    explicitMin,
    explicitMax,
    axisLenPt,
    axisOrientation,
    majorUnit,
    minorUnit: mu,
    needMinor: chart.valAxisMinorGridlines === true || needsMinorTicks,
    logBase,
    reversed,
  });
  const { min, max, majorUnit: step, majorTicks: majorLines } = numeric;
  const minorLines = chart.valAxisMinorGridlines ? numeric.minorTicks : [];
  return {
    min, max, step, majorLines, minorLines, minorTicks: numeric.minorTicks,
    frac: numeric.fraction,
  };
}


/** Resolve an axis label font size (px) from <c:txPr> hpt or a proportional
 *  fallback. ptToPx comes from the host renderer (EMU/px scale at display). */
export function axisLabelPx(sizeHpt: number | null | undefined, h: number, ptToPx: number): number {
  return chartTextFontSizePx(sizeHpt, ptToPx) ?? Math.max(8, h * 0.045);
}


/** Wrap text against the active canvas font without discarding characters.
 * Words are kept intact when possible; a single over-wide token is split at
 * measured character boundaries. Used by chart families whose category-label
 * band is an input to plot layout. */
export function wrapMeasuredText(
  ctx: CanvasRenderingContext2D,
  text: string,
  maxWidth: number,
  singleTokenOverhangPx = 0,
): string[] {
  const words = text.trim().split(/\s+/).filter(Boolean);
  if (words.length === 0) return [''];
  const lines: string[] = [];
  let line = '';
  const pushToken = (token: string): void => {
    const trial = line ? `${line} ${token}` : token;
    if (ctx.measureText(trial).width <= maxWidth) {
      line = trial;
      return;
    }
    if (line) {
      lines.push(line);
      line = '';
    }
    if (ctx.measureText(token).width <= maxWidth + singleTokenOverhangPx) {
      line = token;
      return;
    }
    // Find each largest fitting code-point prefix by binary search. Measuring
    // every growing prefix makes a single long unbroken label quadratic.
    const chars = Array.from(token);
    let start = 0;
    while (start < chars.length) {
      let low = start + 1;
      let high = chars.length;
      let end = start + 1; // Always make progress, even if one glyph is wider.
      while (low <= high) {
        const mid = Math.floor((low + high) / 2);
        if (ctx.measureText(chars.slice(start, mid).join('')).width <= maxWidth) {
          end = mid;
          low = mid + 1;
        } else {
          high = mid - 1;
        }
      }
      const chunk = chars.slice(start, end).join('');
      start = end;
      if (start < chars.length) lines.push(chunk);
      else line = chunk;
    }
  };
  for (const word of words) pushToken(word);
  if (line) lines.push(line);
  return lines.length ? lines : [''];
}


/** Whether the CATEGORY tick labels should be drawn. `<c:catAx><c:tickLblPos
 *  val="none">` (ECMA-376 §21.2.2.207) hides them; anything else (incl. absent)
 *  shows them, so the default is byte-stable. */
export function catLabelsVisible(chart: ChartModel): boolean {
  return chart.catAxisTickLabelPos !== 'none';
}


/** 90° in 60000ths of a degree. `ST_FixedAngle` (ECMA-376 §20.1.10.23) bounds
 *  a fixed-range angle to the OPEN interval "greater than -5400000 / less than
 *  5400000", so ±5400000 itself lies outside the schema type — but Office's
 *  Format-Axis "Custom angle" control accepts -90°…+90° INCLUSIVE, so the code
 *  below deliberately uses a closed boundary (`> LIMIT` rejects, `== LIMIT`
 *  honors) to keep genuine ±90° (vertical) axis labels working. */
export const FIXED_ANGLE_LIMIT_60K = 5_400_000;


/** Category-axis label rotation in RADIANS (canvas convention), from
 *  `<c:catAx|dateAx><c:txPr><a:bodyPr rot>` (DrawingML `ST_Angle`
 *  §20.1.10.3, 60000ths of a degree). Returns 0 when unset — the un-rotated
 *  fast path callers keep.
 *
 *  `bodyPr@rot` is typed `ST_Angle` (a restriction of XML Schema `int`, so any
 *  integer is schema-valid), but a *text* rotation is only meaningful within
 *  the `ST_FixedAngle` (§20.1.10.23) fixed-angle domain — an open interval
 *  (-90°, 90°) at the schema level, which Office's Format-Axis "Custom angle"
 *  control widens to -90°…+90° inclusive (we follow the UI's closed range; see
 *  {@link FIXED_ANGLE_LIMIT_60K}). Office writes `rot="-60000000"` (-1000°,
 *  ≈2.8 full turns) as a sentinel for "auto / horizontal" axis text and renders
 *  those labels horizontal; the identical value even appears on the numeric
 *  value axes whose Office-rendered labels are horizontal. So a rot whose magnitude exceeds ±90°
 *  is outside the valid text-rotation domain and is treated as no rotation
 *  (0°) rather than reduced mod 360 (which would map -1000° → +80°,
 *  near-vertical). Genuine rotations within the
 *  closed range (-45° = -2700000, -90° = -5400000) are honored unchanged. */
export function catLabelRotationRad(chart: ChartModel): number {
  const rot = chart.catAxisLabelRotation;
  if (rot == null || rot === 0) return 0;
  if (Math.abs(rot) > FIXED_ANGLE_LIMIT_60K) return 0;
  return (rot / 60000) * (Math.PI / 180);
}


/** Draw a category label at `(x, y)` with optional rotation. `rotRad === 0`
 *  keeps the exact non-rotated draw the callers used before (byte-stable):
 *  `ctx.fillText(text, x, y)` with the caller's current align/baseline. When
 *  rotated, the label pivots around `(x, y)` and is right-aligned+middle so the
 *  text trails up-left from the tick, matching PowerPoint's angled axis labels. */
export function drawRotatedCatLabel(
  ctx: CanvasRenderingContext2D, text: string, x: number, y: number, rotRad: number,
): void {
  if (rotRad === 0) {
    ctx.fillText(text, x, y);
    return;
  }
  ctx.save();
  ctx.translate(x, y);
  ctx.rotate(rotRad);
  ctx.textAlign = 'right';
  ctx.textBaseline = 'middle';
  ctx.fillText(text, 0, 0);
  ctx.restore();
}

/** Draw an axis title at an explicit anchor in the outer gutter band. The
 *  side-based compatibility rotation is resolved in one place, with authored
 *  DrawingML body orientation remaining authoritative. */
export function drawAxisTitle(
  ctx: CanvasRenderingContext2D,
  text: string,
  anchorX: number, anchorY: number,
  side: ChartAxisTitleSide,
  fontSizePx: number,
  bold: boolean,
  italic: boolean,
  color: string,
  // Available run length along the axis (plot width for the bottom cat title,
  // plot height for the rotated val title). Titles longer than the axis are
  // elided with an ellipsis rather than hard-cut at a fixed char count.
  maxPx: number,
  // Resolved CSS font-family (element face ?? theme heading ?? sans-serif).
  fontFamily = 'sans-serif',
  authoredRotation?: number | null,
  authoredVerticalMode?: ChartModel['catAxisTitleVerticalMode'],
  manualLayout?: ChartManualLayout | null,
  chartRect?: ChartRect,
  box?: ChartLabelBox | null,
  ptToPx = 1,
): void {
  ctx.save();
  ctx.font = chartFontCss(fontSizePx, fontFamily, bold, italic);
  ctx.fillStyle = color;
  // Automatic titles stay bounded to the axis run. An authored title layout
  // is authoritative and keeps its complete text rather than being elided by
  // the automatic plot-width estimate.
  const label = manualLayout ? text : elideToWidth(ctx, text, maxPx);
  const rotation = axisTitleRotationRad(side, authoredRotation, authoredVerticalMode);
  let resolvedAnchorX = anchorX;
  let resolvedAnchorY = anchorY;
  if (manualLayout && chartRect) {
    const textWidth = ctx.measureText(label).width;
    // CT_Title manual-layout x/y position the title's axis-aligned box after
    // DrawingML rotation. A vertical title therefore has a box approximately
    // one font line wide and one text run tall; using the unrotated dimensions
    // shifts it into the tick-label/plot bands by half the text length.
    const cos = Math.abs(Math.cos(rotation));
    const sin = Math.abs(Math.sin(rotation));
    const fittedWidth = textWidth * cos + fontSizePx * sin;
    const fittedHeight = textWidth * sin + fontSizePx * cos;
    const automatic = {
      x: anchorX - fittedWidth / 2,
      y: anchorY - fittedHeight / 2,
      w: fittedWidth,
      h: fittedHeight,
    };
    // CT_Title manual layout positions the title box, while Office keeps the
    // box fitted to its text. Match the existing chart-title rule: x/y win,
    // authored w/h do not stretch or shrink the text box.
    const resolved = resolveManualLayoutRect(
      { ...manualLayout, w: undefined, h: undefined },
      chartRect,
      automatic,
    );
    if (resolved) {
      resolvedAnchorX = resolved.x + resolved.w / 2;
      resolvedAnchorY = resolved.y + resolved.h / 2;
    }
  }
  ctx.translate(resolvedAnchorX, resolvedAnchorY);
  if (rotation !== 0) ctx.rotate(rotation);
  const textWidth = ctx.measureText(label).width;
  paintChartLabelBox(ctx, box, {
    x: -textWidth / 2,
    y: -fontSizePx / 2,
    w: textWidth,
    h: fontSizePx,
  }, ptToPx);
  ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
  ctx.fillText(label, 0, 0);
  ctx.restore();
}

/** Resolve the per-axis title color string for `drawAxisTitle`. Returns
 *  '#rrggbb' when the XML supplied a srgb color, else the legacy '#555'. */
export function axisTitleColor(hex: string | null | undefined): string {
  return hex ? `#${hex}` : '#555';
}

/** Draw both axis titles for a cartesian chart (bar/line/area/scatter),
 *  anchored in the reserved outer gutter bands so they sit OUTSIDE the tick
 *  labels. `catTitlePx`/`valTitlePx` are the title font sizes the caller used
 *  to size `catTitleH`/`valTitleW`; the anchor centers each title within its
 *  band. Column/line/area/scatter use cat-bottom + val-left. Horizontal bars
 *  use cat-left + val-bottom because their value axis runs horizontally.
 *  Bold and italic are independent DrawingML character properties. The parser
 *  resolves authored/inherited OOXML values, including the regular-weight
 *  DrawingML base fallback. A hand-built public model that leaves bold unset
 *  retains the renderer's established bold compatibility fallback. The
 *  separate 10pt size fallback is the product policy in #1228. */
export function drawAxisTitles(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  x: number, y: number, w: number, h: number,
  px0: number, py0: number, pw: number, ph: number,
  legLeftW: number, legBottomH: number,
  catTitlePx: number, valTitlePx: number,
  horizontalValueAxis = false,
): void {
  const drawPrimaryTitle = (
    text: string,
    side: ChartAxisTitleSide,
    fontSizePx: number,
    bold: boolean,
    italic: boolean,
    color: string,
    fontFamily: string,
    authoredRotation: number | null | undefined,
    authoredVerticalMode: ChartModel['catAxisTitleVerticalMode'],
    manualLayout: ChartManualLayout | null | undefined,
    directStyle: ChartExStyle | null | undefined,
  ): void => {
    const box = effectiveLinkedLabelBox(
      chart,
      directStyle ? { style: directStyle } : undefined,
      chart.chartStyleRoles?.axisTitle,
      rawLinkedChartStyleRole(chart, 'axisTitle'),
      true,
    );
    if (side === 'left') {
      drawAxisTitle(
        ctx, text,
        x + legLeftW + axisTitleMargin(w) + fontSizePx / 2,
        py0 + ph / 2,
        side, fontSizePx, bold, italic, color, ph, fontFamily, authoredRotation,
        authoredVerticalMode, manualLayout, { x, y, w, h }, box,
        fontSizePx / 10,
      );
      return;
    }
    drawAxisTitle(
      ctx, text,
      px0 + pw / 2,
      y + h - legBottomH - axisTitleMargin(h) - fontSizePx / 2,
      side, fontSizePx, bold, italic, color, pw, fontFamily, authoredRotation,
      authoredVerticalMode, manualLayout, { x, y, w, h }, box,
      fontSizePx / 10,
    );
  };
  if (chart.valAxisTitle) {
    drawPrimaryTitle(
      chart.valAxisTitle, horizontalValueAxis ? 'horizontal' : 'left',
      valTitlePx, chart.valAxisTitleFontBold ?? true, chart.valAxisTitleFontItalic ?? false,
      axisTitleColor(chart.valAxisTitleFontColor),
      chartFontFamily(chart, chart.valAxisTitleFontFace, 'major'), chart.valAxisTitleRotation,
      chart.valAxisTitleVerticalMode,
      chart.valAxisTitleManualLayout,
      chart.valAxisTitleStyle,
    );
  }
  if (chart.catAxisTitle) {
    drawPrimaryTitle(
      chart.catAxisTitle, horizontalValueAxis ? 'left' : 'horizontal',
      catTitlePx, chart.catAxisTitleFontBold ?? true, chart.catAxisTitleFontItalic ?? false,
      axisTitleColor(chart.catAxisTitleFontColor),
      chartFontFamily(chart, chart.catAxisTitleFontFace, 'major'), chart.catAxisTitleRotation,
      chart.catAxisTitleVerticalMode,
      chart.catAxisTitleManualLayout,
      chart.catAxisTitleStyle,
    );
  }
}

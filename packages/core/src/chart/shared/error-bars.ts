// Classic chart error bars helpers.
import type { ChartSeries } from '../../types/chart';
import { EMU_PER_PT } from '../../units.js';
import { dashPatternForPreset } from './geometry.js';
import { scatterXValue } from './scatter-geometry.js';


/** Draw error bars for one series + one direction. Each segment is a line
 *  from the data point to the offset point, plus an optional perpendicular
 *  end-cap (skipped when `eb.noEndCap`). */
export function drawSeriesErrorBars(
  ctx: CanvasRenderingContext2D,
  s: ChartSeries,
  eb: NonNullable<ChartSeries['errBars']>[number],
  cats: string[],
  useIndexX: boolean,
  toX: (v: number) => number,
  toY: (v: number) => number,
  fallbackColor: string,
): void {
  if (eb.hidden === true || (eb.linePaintAuthored === true && eb.color == null)) return;
  ctx.save();
  ctx.strokeStyle = eb.color ? `#${eb.color}` : fallbackColor;
  ctx.lineWidth = eb.lineWidthEmu ? Math.max(0.5, eb.lineWidthEmu / EMU_PER_PT) : 1;
  ctx.setLineDash(dashPatternForPreset(eb.dash, ctx.lineWidth));
  const drawPlus = eb.barType === 'plus' || eb.barType === 'both';
  const drawMinus = eb.barType === 'minus' || eb.barType === 'both';
  const isX = eb.dir === 'x';
  // Office's error-bar cap spans one stroke width. Keeping the cap square with
  // the authored error-bar stroke also lets a same-size endpoint marker cover
  // it, as Excel does; the former 3× stroke-width cap protruded above/below
  // overlaid markers.
  const capHalf = ctx.lineWidth / 2;
  for (let i = 0; i < s.values.length; i++) {
    const yv = s.values[i]; if (yv == null) continue;
    const xv = scatterXValue(cats, i, useIndexX);
    if (xv == null) continue;
    const px = toX(xv); const py = toY(yv);
    const drawSeg = (dataDelta: number) => {
      let x2 = px, y2 = py;
      if (isX) {
        // X delta is in data X units, so map (xv + delta) → px. For the
        // minus side delta is already a positive magnitude, flip the sign.
        x2 = toX(xv + dataDelta);
      } else {
        // Y delta similar; positive moves the bar toward higher data values
        // (visually upward for our orientation).
        y2 = toY(yv + dataDelta);
      }
      ctx.beginPath();
      ctx.moveTo(px, py); ctx.lineTo(x2, y2); ctx.stroke();
      if (!eb.noEndCap) {
        ctx.save(); ctx.setLineDash([]);
        ctx.beginPath();
        if (isX) {
          ctx.moveTo(x2, y2 - capHalf); ctx.lineTo(x2, y2 + capHalf);
        } else {
          ctx.moveTo(x2 - capHalf, y2); ctx.lineTo(x2 + capHalf, y2);
        }
        ctx.stroke();
        ctx.restore();
      }
    };
    // ECMA-376 §21.2.2.20: plus side is `point + plus[i]`, minus side is
    // `point - minus[i]`. For `cust` errValType the values may be signed
    // (e.g. negative minus values that effectively flip direction); for
    // `fixedVal`/`stdErr`/`stdDev`/`percentage` the parser stores positive
    // magnitudes, so the same formula gives the expected direction.
    if (drawPlus) {
      const v = eb.plus[i]; if (v != null) drawSeg(v);
    }
    if (drawMinus) {
      const v = eb.minus[i]; if (v != null) drawSeg(-v);
    }
  }
  ctx.restore();
}


/** Draw error bars for a category-axis series (line / area). Mirrors the scatter
 *  {@link drawSeriesErrorBars} cap/dash geometry, but maps points by CATEGORY
 *  INDEX (`xAt(ci)`) with a per-series value→px mapping (`yAt`) instead of the
 *  numeric X mapping scatter uses. Only the Y direction is drawn: a category
 *  axis has no data-unit X scale, so `<c:errBars dir="x">` cannot be positioned
 *  (Excel likewise only shows Y error bars on category charts). `plotted`
 *  returns the point's plotted (possibly stacked) value so bars ride the drawn
 *  line. Null cells are skipped. */
export function drawCategoryErrorBars(
  ctx: CanvasRenderingContext2D,
  s: ChartSeries,
  eb: NonNullable<ChartSeries['errBars']>[number],
  n: number,
  xAt: (ci: number) => number,
  yAt: (v: number) => number,
  plotted: (ci: number) => number,
  fallbackColor: string,
): void {
  if (eb.hidden === true || (eb.linePaintAuthored === true && eb.color == null)
    || eb.dir === 'x') return; // no data-unit X scale on a category axis
  const drawPlus = eb.barType === 'plus' || eb.barType === 'both';
  const drawMinus = eb.barType === 'minus' || eb.barType === 'both';
  ctx.save();
  ctx.strokeStyle = eb.color ? `#${eb.color}` : fallbackColor;
  ctx.lineWidth = eb.lineWidthEmu ? Math.max(0.5, eb.lineWidthEmu / EMU_PER_PT) : 1;
  ctx.setLineDash(dashPatternForPreset(eb.dash, ctx.lineWidth));
  const capHalf = ctx.lineWidth / 2;
  for (let ci = 0; ci < n; ci++) {
    if (s.values[ci] == null) continue;
    const pv = plotted(ci);
    const px = xAt(ci); const py = yAt(pv);
    const drawSeg = (dataDelta: number): void => {
      const y2 = yAt(pv + dataDelta);
      ctx.beginPath(); ctx.moveTo(px, py); ctx.lineTo(px, y2); ctx.stroke();
      if (!eb.noEndCap) {
        ctx.save(); ctx.setLineDash([]);
        ctx.beginPath();
        ctx.moveTo(px - capHalf, y2); ctx.lineTo(px + capHalf, y2);
        ctx.stroke();
        ctx.restore();
      }
    };
    if (drawPlus) { const v = eb.plus[ci]; if (v != null) drawSeg(v); }
    if (drawMinus) { const v = eb.minus[ci]; if (v != null) drawSeg(-v); }
  }
  ctx.restore();
}

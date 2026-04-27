// Sparkline renderer (Office 2010 `x14:sparklineGroup`, ECMA-376 §18.2).
// Drawn directly inside a single cell's pixel rect — no axis labels, legend,
// or title. Theme colors are expected to be flattened to `#RRGGBB` strings
// at parse time so this module has zero theme awareness.

export type SparklineKind = 'line' | 'column' | 'stem';

/** Render-ready sparkline. The xlsx renderer flattens its parser output into
 *  this shape (resolving group-level colors / flags onto every cell) and
 *  passes it to {@link renderSparkline}. */
export interface SparklineModel {
  kind: SparklineKind;
  /** Numeric values; null = empty / non-numeric (rendered as a gap). */
  values: (number | null)[];
  /** Optional axis bounds. Used directly when present; otherwise derived
   *  from the values themselves. For the `group` axis sharing mode the
   *  caller computes the shared bounds and passes them in. */
  min?: number;
  max?: number;
  /** ECMA-376 §18.2.7. `gap` (default) breaks the line at empty cells;
   *  `zero` substitutes 0; `span` connects across empty cells. */
  displayEmptyCellsAs?: 'gap' | 'zero' | 'span';
  /** Show horizontal axis line at y=0 when the data crosses zero
   *  (`displayXAxis`). */
  displayXAxis?: boolean;
  /** Stroke weight in pt for `line`. ECMA-376 default 0.75. */
  lineWeight?: number;
  /** Show a marker dot at every data point (line type only). */
  markers?: boolean;
  /** Enable individual point highlights (resolved colors below override
   *  `colorSeries` for that point). */
  high?: boolean;
  low?: boolean;
  first?: boolean;
  last?: boolean;
  negative?: boolean;
  /** Resolved RGB hex strings (`#RRGGBB`). */
  colorSeries?: string;
  colorNegative?: string;
  colorAxis?: string;
  colorMarkers?: string;
  colorFirst?: string;
  colorLast?: string;
  colorHigh?: string;
  colorLow?: string;
}

export interface SparklineRect {
  x: number; y: number; w: number; h: number;
}

/** EMU-free, theme-free sparkline renderer. The cell rect is in device px
 *  matching the canvas's current transform; padding is applied internally. */
export function renderSparkline(
  ctx: CanvasRenderingContext2D,
  rect: SparklineRect,
  model: SparklineModel,
): void {
  const { values } = model;
  if (values.length === 0 || rect.w <= 0 || rect.h <= 0) return;

  // Inset the drawing area so strokes / markers don't kiss the cell edge.
  // Half a pixel on each side is plenty at typical sizes; a marker radius
  // worth on the top / bottom keeps high / low dots inside.
  const series = model.colorSeries ?? '#5B9BD5';
  const PAD_X = Math.min(2, rect.w * 0.08);
  const PAD_Y = Math.min(2, rect.h * 0.12);
  const dx = rect.x + PAD_X;
  const dy = rect.y + PAD_Y;
  const dw = Math.max(1, rect.w - PAD_X * 2);
  const dh = Math.max(1, rect.h - PAD_Y * 2);

  // Bounds. Empty data → bail quietly.
  const numeric = values.filter((v): v is number => typeof v === 'number');
  if (numeric.length === 0) return;
  const dataMin = Math.min(...numeric);
  const dataMax = Math.max(...numeric);
  let lo = model.min ?? dataMin;
  let hi = model.max ?? dataMax;
  if (lo === hi) {
    // Flat data — center it in the rect.
    hi = lo + 1;
    lo = lo - 1;
  }
  const range = hi - lo;
  const yOf = (v: number) => dy + dh - ((v - lo) / range) * dh;

  if (model.kind === 'stem') {
    drawStem(ctx, model, dx, dy, dw, dh);
    return;
  }

  if (model.kind === 'column') {
    drawColumn(ctx, model, values, dx, dy, dw, dh, lo, hi);
    return;
  }

  // ---- line ----
  // Axis line at y=0 when the data crosses zero and displayXAxis is on.
  if (model.displayXAxis && lo < 0 && hi > 0) {
    ctx.save();
    ctx.strokeStyle = model.colorAxis ?? '#000000';
    ctx.lineWidth = 1;
    ctx.beginPath();
    const ay = yOf(0);
    ctx.moveTo(dx, ay);
    ctx.lineTo(dx + dw, ay);
    ctx.stroke();
    ctx.restore();
  }

  // Each point's center X. With one point we still want a stable position.
  const n = values.length;
  const xOf = (i: number) => n === 1 ? dx + dw / 2 : dx + (i / (n - 1)) * dw;

  // Stroke the polyline, breaking at gaps unless `span` was requested.
  ctx.save();
  ctx.strokeStyle = series;
  ctx.lineCap = 'round';
  ctx.lineJoin = 'round';
  // ECMA-376 lineWeight is in points; the canvas is already scaled, so
  // 1 pt ≈ 1.333 px at 96 DPI. We multiply by ptToPx to keep visual weight
  // consistent with PowerPoint / Excel.
  const PT_TO_PX = 1.333;
  ctx.lineWidth = (model.lineWeight ?? 0.75) * PT_TO_PX;
  ctx.beginPath();
  let penDown = false;
  const empty = model.displayEmptyCellsAs ?? 'gap';
  for (let i = 0; i < n; i++) {
    const v = values[i];
    if (v == null) {
      if (empty === 'zero') {
        const x = xOf(i), y = yOf(0);
        if (!penDown) { ctx.moveTo(x, y); penDown = true; }
        else ctx.lineTo(x, y);
      } else if (empty === 'gap') {
        penDown = false;
      }
      // span: skip — the next non-null point will continue the line.
      continue;
    }
    const x = xOf(i), y = yOf(v);
    if (!penDown) { ctx.moveTo(x, y); penDown = true; }
    else ctx.lineTo(x, y);
  }
  ctx.stroke();
  ctx.restore();

  // Markers / highlights. Only line type honors these visually.
  const dotR = Math.max(1, Math.min(2.5, dh * 0.12));
  const flagged = computeFlagged(values, model);
  for (let i = 0; i < n; i++) {
    const v = values[i];
    if (v == null) continue;
    const flagColor = flagged[i];
    const drawDot = model.markers || flagColor != null;
    if (!drawDot) continue;
    ctx.save();
    ctx.fillStyle = flagColor ?? model.colorMarkers ?? series;
    ctx.beginPath();
    ctx.arc(xOf(i), yOf(v), dotR, 0, Math.PI * 2);
    ctx.fill();
    ctx.restore();
  }
}

/** Resolve the per-point highlight color for line / column. Returns
 *  `null` for points with no special treatment. Order of precedence
 *  (later wins): negative < first < last < high < low. */
function computeFlagged(
  values: (number | null)[],
  m: SparklineModel,
): Array<string | null> {
  const flagged: Array<string | null> = values.map(() => null);
  const numeric = values.map(v => (typeof v === 'number' ? v : null));
  const firstIdx = numeric.findIndex(v => v != null);
  let lastIdx = -1;
  for (let i = numeric.length - 1; i >= 0; i--) {
    if (numeric[i] != null) { lastIdx = i; break; }
  }
  const present = numeric.filter((v): v is number => v != null);
  let hi = NaN, lo = NaN;
  if (present.length > 0) {
    hi = Math.max(...present);
    lo = Math.min(...present);
  }
  if (m.negative && m.colorNegative) {
    for (let i = 0; i < numeric.length; i++) {
      const v = numeric[i];
      if (v != null && v < 0) flagged[i] = m.colorNegative;
    }
  }
  if (m.first && m.colorFirst && firstIdx >= 0) flagged[firstIdx] = m.colorFirst;
  if (m.last && m.colorLast && lastIdx >= 0) flagged[lastIdx] = m.colorLast;
  // ECMA-376 §18.18.74 doesn't say so explicitly, but Excel highlights
  // *every* point tied for the high or low value (not just the first
  // occurrence). With a 12-point series of 9 zeros + 3 non-zeros and
  // `low="1"`, all 9 zero points are dotted, not just the first.
  if (m.high && m.colorHigh && !Number.isNaN(hi)) {
    for (let i = 0; i < numeric.length; i++) {
      if (numeric[i] === hi) flagged[i] = m.colorHigh;
    }
  }
  if (m.low && m.colorLow && !Number.isNaN(lo)) {
    for (let i = 0; i < numeric.length; i++) {
      if (numeric[i] === lo) flagged[i] = m.colorLow;
    }
  }
  return flagged;
}

function drawColumn(
  ctx: CanvasRenderingContext2D,
  m: SparklineModel,
  values: (number | null)[],
  x: number, y: number, w: number, h: number,
  lo: number, hi: number,
) {
  const n = values.length;
  if (n === 0) return;
  // Bars share a baseline of 0 when the data crosses it; otherwise the
  // baseline is the min edge (matches Excel's positive-only / negative-only
  // sparkline columns growing from the bottom / top of the cell).
  const baseValue = (lo < 0 && hi > 0) ? 0 : lo;
  const range = hi - lo;
  const yOf = (v: number) => y + h - ((v - lo) / range) * h;
  const baseY = yOf(baseValue);
  const slot = w / n;
  // Tiny cell-side gap so adjacent bars are distinguishable. Excel uses
  // ~1 px on a small bar and we match by clamping at 1.
  const gap = Math.min(1.5, slot * 0.15);
  const flagged = computeFlagged(values, m);
  for (let i = 0; i < n; i++) {
    const v = values[i];
    if (v == null) continue;
    const fill = flagged[i] ?? (v < 0 && m.colorNegative ? m.colorNegative : (m.colorSeries ?? '#5B9BD5'));
    const top = yOf(v);
    const barX = x + slot * i + gap / 2;
    const barW = Math.max(1, slot - gap);
    ctx.save();
    ctx.fillStyle = fill;
    ctx.fillRect(barX, Math.min(baseY, top), barW, Math.abs(baseY - top));
    ctx.restore();
  }
}

function drawStem(
  ctx: CanvasRenderingContext2D,
  m: SparklineModel,
  x: number, y: number, w: number, h: number,
) {
  // ECMA-376: win/loss bars are fixed half-height — positives reach up
  // from a midline, negatives reach down. Zeros are skipped.
  const n = m.values.length;
  if (n === 0) return;
  const midY = y + h / 2;
  const halfH = h / 2;
  const slot = w / n;
  const gap = Math.min(1.5, slot * 0.15);
  const flagged = computeFlagged(m.values, m);
  for (let i = 0; i < n; i++) {
    const v = m.values[i];
    if (v == null || v === 0) continue;
    const isNeg = v < 0;
    const fill = flagged[i] ?? (isNeg && m.colorNegative ? m.colorNegative : (m.colorSeries ?? '#5B9BD5'));
    const barX = x + slot * i + gap / 2;
    const barW = Math.max(1, slot - gap);
    ctx.save();
    ctx.fillStyle = fill;
    if (isNeg) ctx.fillRect(barX, midY, barW, halfH);
    else ctx.fillRect(barX, midY - halfH, barW, halfH);
    ctx.restore();
  }
}

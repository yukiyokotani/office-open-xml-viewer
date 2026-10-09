/** Clip a local x interval across the device viewport, preserving vertical
 * overflow. Projected corners and their y interval must be representable as
 * finite numbers; singular or unrepresentable geometry leaves clipping unchanged.
 * The caller owns save/restore. PPTX cells consume this primitive. DOCX uses
 * layout-owned 2-D clipBounds; XLSX clips row height and adjacent-cell ranges.
 */
export function clipCanvasHorizontally(
  ctx: CanvasRenderingContext2D, left: number, width: number,
): void {
  const m = typeof ctx.getTransform === 'function' ? ctx.getTransform()
    : { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
  if (![left, width, left + width, ctx.canvas.width, ctx.canvas.height,
    m.a, m.b, m.c, m.d, m.e, m.f].every(Number.isFinite)) return;
  // Normalize each matrix row before inversion. Direct a*d - b*c can
  // overflow/underflow for an otherwise invertible, representable viewport.
  const xScale = Math.max(Math.abs(m.a), Math.abs(m.c));
  const yScale = Math.max(Math.abs(m.b), Math.abs(m.d));
  if (xScale === 0 || yScale === 0) return;
  const a = m.a / xScale, c = m.c / xScale;
  const b = m.b / yScale, d = m.d / yScale;
  const det = a * d - b * c;
  if (det === 0) return;
  let top = Infinity;
  let bottom = -Infinity;
  for (const x of [0, ctx.canvas.width]) for (const y of [0, ctx.canvas.height]) {
    const localY = ((b === 0 ? 0 : -b * ((x - m.e) / xScale)) +
      (a === 0 ? 0 : a * ((y - m.f) / yScale))) / det;
    if (!Number.isFinite(localY)) return;
    top = Math.min(top, localY);
    bottom = Math.max(bottom, localY);
  }
  const height = bottom - top;
  if (!Number.isFinite(height)) return;
  ctx.beginPath();
  ctx.rect(left, top, width, height);
  ctx.clip();
}

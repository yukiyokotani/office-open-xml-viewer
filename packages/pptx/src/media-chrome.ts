/**
 * Draw a centered play/pause badge over a media poster rectangle.
 *
 * Shared by the static renderer ({@link ./renderer.ts}) and the interactive
 * presentation handle ({@link ./presentation-handle.ts}) so the two paths
 * produce a visually identical badge.
 */
export function drawPlayBadge(
  ctx: CanvasRenderingContext2D,
  cx: number,
  cy: number,
  posterW: number,
  posterH: number,
  state: 'paused' | 'playing',
): void {
  const badgeR = Math.max(18, Math.min(32, Math.min(posterW, posterH) * 0.25));
  ctx.save();
  ctx.shadowColor = 'rgba(0, 0, 0, 0.3)';
  ctx.shadowBlur = badgeR * 0.35;
  ctx.fillStyle = 'rgba(20, 20, 20, 0.7)';
  ctx.beginPath();
  ctx.arc(cx, cy, badgeR, 0, Math.PI * 2);
  ctx.fill();
  ctx.shadowColor = 'transparent';
  ctx.shadowBlur = 0;
  ctx.fillStyle = '#fff';
  if (state === 'paused') {
    ctx.beginPath();
    const tri = badgeR * 0.48;
    ctx.moveTo(cx - tri * 0.4, cy - tri);
    ctx.lineTo(cx - tri * 0.4, cy + tri);
    ctx.lineTo(cx + tri * 0.75, cy);
    ctx.closePath();
    ctx.fill();
  } else {
    const bw = badgeR * 0.2;
    const bh = badgeR * 0.8;
    const gap = badgeR * 0.15;
    ctx.fillRect(cx - gap - bw, cy - bh / 2, bw, bh);
    ctx.fillRect(cx + gap, cy - bh / 2, bw, bh);
  }
  ctx.restore();
}

import {
  ERROR_PLACEHOLDER_DETAIL_LINES,
  errorPlaceholderText,
  resolvedFromWire,
  type ErrorPlaceholderTextHost,
} from '@silurus/ooxml-core/internal/error-placeholder-text';
import type { WireErrorPlaceholder } from '@silurus/ooxml-core/internal/part-failures';

/**
 * RB7: paint a placeholder for a slide whose part failed to parse. A neutral
 * card with a warning glyph, a heading, and the part-tagged error — so a viewer
 * shows "this one slide is broken" in place, and the rest of the deck renders
 * normally. Widths are in CSS px (the ctx is already dpr-scaled). Only ever
 * called for a slide carrying `parseError`, so healthy slides never touch it:
 * the renderer loads this module with the first such slide (#2248), and hands
 * it the functions of the static graph the text needs (see
 * `error-placeholder-text.ts` for why they are not imported).
 */
export function parseErrorPlaceholder(host: ErrorPlaceholderTextHost) {
  const { errorPlaceholderCanvasLines, errorPlaceholderDetail, errorPlaceholderTitle } = errorPlaceholderText(host);
  function drawParseErrorPlaceholder(
    ctx: CanvasRenderingContext2D,
    widthPx: number,
    heightPx: number,
    slideNumber: number,
    message: string,
    wire: WireErrorPlaceholder | undefined,
  ): void {
    const placeholder = resolvedFromWire(wire);
    ctx.save();
    // Neutral card fill + dashed border, inset from the slide edge.
    ctx.fillStyle = '#f7f7f8';
    ctx.fillRect(0, 0, widthPx, heightPx);
    const pad = Math.max(12, Math.min(widthPx, heightPx) * 0.04);
    ctx.strokeStyle = '#c8ccd2';
    ctx.lineWidth = Math.max(1, Math.min(widthPx, heightPx) * 0.004);
    ctx.setLineDash([ctx.lineWidth * 6, ctx.lineWidth * 5]);
    ctx.strokeRect(pad, pad, widthPx - pad * 2, heightPx - pad * 2);
    ctx.setLineDash([]);

    const cx = widthPx / 2;
    // Warning glyph, scaled to the card.
    const glyph = Math.max(18, Math.min(widthPx, heightPx) * 0.14);
    ctx.fillStyle = '#b23b3b';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.font = `${glyph}px sans-serif`;
    ctx.fillText('⚠', cx, heightPx * 0.34);

    // Heading.
    const headSize = Math.max(11, Math.min(widthPx, heightPx) * 0.045);
    ctx.fillStyle = '#333333';
    ctx.font = `600 ${headSize}px sans-serif`;
    const maxLineWidth = widthPx - pad * 4;
    const title = errorPlaceholderTitle(placeholder, `Slide ${slideNumber} could not be displayed`);
    for (const line of errorPlaceholderCanvasLines(ctx, title, maxLineWidth, 1)) {
      ctx.fillText(line, cx, heightPx * 0.52);
    }
    const detail = errorPlaceholderDetail(placeholder, message);
    if (detail === null) {
      ctx.restore();
      return;
    }

    // Error detail (part path + reason), wrapped to the card width and clipped to
    // a few lines so a long message never overflows the frame.
    const detailSize = Math.max(9, Math.min(widthPx, heightPx) * 0.028);
    ctx.fillStyle = '#666666';
    ctx.font = `${detailSize}px sans-serif`;
    const lines = errorPlaceholderCanvasLines(ctx, detail, maxLineWidth, ERROR_PLACEHOLDER_DETAIL_LINES);
    const lineHeight = detailSize * 1.35;
    let y = heightPx * 0.6 + lineHeight;
    for (const l of lines) {
      ctx.fillText(l, cx, y);
      y += lineHeight;
    }
    ctx.restore();
  }

  return drawParseErrorPlaceholder;
}

import {
  ERROR_PLACEHOLDER_DETAIL_LINES,
  errorPlaceholderText,
  resolvedFromWire,
  type ErrorPlaceholderTextHost,
} from '@silurus/ooxml-core/internal/error-placeholder-text';
import type { WireErrorPlaceholder } from '@silurus/ooxml-core/internal/part-failures';

/**
 * RB7: paint a placeholder overlay for a worksheet whose part failed to parse.
 * A neutral fill, a warning glyph, a heading naming the sheet, and the
 * part-tagged error wrapped to a few lines. Coordinates are in CSS px (the ctx
 * is already dpr-scaled by the caller). Only ever called for a sheet carrying
 * `parseError`: the renderer loads this module with the first such sheet
 * (#2248), and hands it the functions of the static graph the text needs (see
 * `error-placeholder-text.ts` for why they are not imported).
 */
export function sheetErrorOverlay(host: ErrorPlaceholderTextHost) {
  const { errorPlaceholderCanvasLines, errorPlaceholderDetail, errorPlaceholderTitle } = errorPlaceholderText(host);
  function drawSheetParseErrorOverlay(
    ctx: CanvasRenderingContext2D,
    widthPx: number,
    heightPx: number,
    sheetName: string,
    message: string,
    wire: WireErrorPlaceholder | undefined,
  ): void {
    const placeholder = resolvedFromWire(wire);
    ctx.save();
    ctx.fillStyle = '#f7f7f8';
    ctx.fillRect(0, 0, widthPx, heightPx);
    const cx = widthPx / 2;
    const base = Math.min(widthPx, heightPx);

    const glyph = Math.max(20, base * 0.1);
    ctx.fillStyle = '#b23b3b';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.font = `${glyph}px sans-serif`;
    ctx.fillText('⚠', cx, heightPx * 0.32);

    const headSize = Math.max(13, base * 0.035);
    ctx.fillStyle = '#333333';
    ctx.font = `600 ${headSize}px sans-serif`;
    // The card is the whole viewport: the title only has to stay inside it,
    // while the detail keeps a narrower reading measure.
    const title = errorPlaceholderTitle(placeholder, `Sheet "${sheetName}" could not be displayed`);
    for (const line of errorPlaceholderCanvasLines(ctx, title, widthPx * 0.9, 1)) {
      ctx.fillText(line, cx, heightPx * 0.46);
    }
    const detail = errorPlaceholderDetail(placeholder, message);
    if (detail === null) {
      ctx.restore();
      return;
    }

    const detailSize = Math.max(10, base * 0.022);
    ctx.fillStyle = '#666666';
    ctx.font = `${detailSize}px sans-serif`;
    const lines = errorPlaceholderCanvasLines(ctx, detail, Math.min(widthPx * 0.8, 640), ERROR_PLACEHOLDER_DETAIL_LINES);
    const lineHeight = detailSize * 1.4;
    let y = heightPx * 0.52 + lineHeight;
    for (const l of lines) {
      ctx.fillText(l, cx, y);
      y += lineHeight;
    }
    ctx.restore();
  }

  return drawSheetParseErrorOverlay;
}

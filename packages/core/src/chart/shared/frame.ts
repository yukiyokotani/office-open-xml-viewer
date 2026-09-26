// Classic chart frame helpers.
import type { ChartModel, ChartRect, ChartTextBox } from '../../types/chart';
import { DEFAULT_TEXT_INSET_LR_EMU, DEFAULT_TEXT_INSET_TB_EMU, EMU_PER_PT } from '../../units.js';
import { chartFontFamily } from './fonts.js';


export function drawChartTextBoxes(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  rect: ChartRect,
  ptToPx: number,
): void {
  const boxes = chart.chartTextBoxes;
  if (!boxes?.length) return;

  for (const box of boxes) {
    const bx = rect.x + box.x * rect.w;
    const by = rect.y + box.y * rect.h;
    const bw = box.w * rect.w;
    const bh = box.h * rect.h;
    if (!(bw > 0 && bh > 0)) continue;
    const contentX = bx + ((box.lIns ?? DEFAULT_TEXT_INSET_LR_EMU) / EMU_PER_PT) * ptToPx;
    const contentY0 = by + ((box.tIns ?? DEFAULT_TEXT_INSET_TB_EMU) / EMU_PER_PT) * ptToPx;
    const contentRight = bx + bw - ((box.rIns ?? DEFAULT_TEXT_INSET_LR_EMU) / EMU_PER_PT) * ptToPx;
    const contentBottom = by + bh - ((box.bIns ?? DEFAULT_TEXT_INSET_TB_EMU) / EMU_PER_PT) * ptToPx;
    const contentW = contentRight - contentX;
    const contentH = contentBottom - contentY0;
    if (!(contentW > 0 && contentH > 0)) continue;

    type MeasuredTextRun = {
      run: ChartTextBox['paragraphs'][number]['runs'][number];
      text: string;
      fontPx: number;
      font: string;
      width: number;
    };
    type MeasuredLine = {
      paragraph: ChartTextBox['paragraphs'][number];
      runs: MeasuredTextRun[];
      width: number;
      height: number;
      baseline: number;
    };

    const makeLine = (
      paragraph: ChartTextBox['paragraphs'][number],
      runs: MeasuredTextRun[],
    ): MeasuredLine => {
      const maxFontPx = Math.max(1, ...runs.map(run => run.fontPx));
      return {
        paragraph,
        runs,
        width: runs.reduce((sum, run) => sum + run.width, 0),
        height: maxFontPx * 1.2,
        baseline: maxFontPx * 0.9,
      };
    };

    const lines = box.paragraphs.flatMap(paragraph => {
      const measuredRuns = paragraph.runs.map(run => {
        const fontPx = Math.max(1, ((run.fontSizeHpt ?? 1000) / 100) * ptToPx);
        const font = `${run.bold ? 'bold ' : ''}${fontPx}px ${chartFontFamily(chart, run.fontFace, 'minor')}`;
        ctx.font = font;
        return { run, text: run.text, fontPx, font, width: ctx.measureText(run.text).width };
      });
      const paragraphWidth = measuredRuns.reduce((sum, run) => sum + run.width, 0);
      if (box.wrap === 'none' || paragraphWidth <= contentW) {
        return [makeLine(paragraph, measuredRuns)];
      }

      const wrapped: MeasuredLine[] = [];
      let current: MeasuredTextRun[] = [];
      let currentWidth = 0;
      const flush = () => {
        if (!current.length) return;
        wrapped.push(makeLine(paragraph, current));
        current = [];
        currentWidth = 0;
      };

      for (const measured of measuredRuns) {
        const tokens = measured.text.match(/\s+|\S+/g) ?? [];
        for (const token of tokens) {
          const whitespace = /^\s+$/.test(token);
          ctx.font = measured.font;
          const tokenWidth = ctx.measureText(token).width;
          if (current.length && currentWidth + tokenWidth > contentW) {
            flush();
          }
          // A wrapped line does not begin with the inter-word whitespace that
          // caused the previous line to overflow.
          if (whitespace && !current.length) continue;
          current.push({ ...measured, text: token, width: tokenWidth });
          currentWidth += tokenWidth;
        }
      }
      flush();
      return wrapped.length ? wrapped : [makeLine(paragraph, measuredRuns)];
    });
    const textHeight = lines.reduce((sum, line) => sum + line.height, 0);
    const contentY = box.verticalAnchor === 'b'
      ? contentBottom - textHeight
      : box.verticalAnchor === 'ctr'
        ? contentY0 + (contentH - textHeight) / 2
        : contentY0;

    ctx.save();
    ctx.beginPath();
    ctx.rect(bx, by, bw, bh);
    ctx.clip();
    ctx.textAlign = 'left';
    ctx.textBaseline = 'alphabetic';
    let lineY = contentY;
    for (const metric of lines) {
      const align = metric.paragraph.align;
      let runX = align === 'ctr'
        ? contentX + (contentW - metric.width) / 2
        : align === 'r'
          ? contentRight - metric.width
          : contentX;
      for (const measured of metric.runs) {
        ctx.font = measured.font;
        ctx.fillStyle = measured.run.color ? `#${measured.run.color}` : '#000000';
        ctx.fillText(measured.text, runX, lineY + metric.baseline);
        runX += measured.width;
      }
      lineY += metric.height;
    }
    ctx.restore();
  }
}


// ─── Background frame + dispatcher ──────────────────────────────────────────

/** ECMA-376 §21.2.2.159 defines only whether chart-space corners are rounded,
 * not the application geometry. Desktop Excel vector output uses a fixed 10pt
 * radius across square, wide, and tall chart frames; keep that observed Office
 * policy isolated from fill, border, and clipping semantics. */
export const CHART_SPACE_CORNER_RADIUS_PT = 10;


export function chartSpaceRoundedPath(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  radius: number,
): void {
  const r = Math.max(0, Math.min(radius, w / 2, h / 2));
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.lineTo(x + w - r, y);
  ctx.quadraticCurveTo(x + w, y, x + w, y + r);
  ctx.lineTo(x + w, y + h - r);
  ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
  ctx.lineTo(x + r, y + h);
  ctx.quadraticCurveTo(x, y + h, x, y + h - r);
  ctx.lineTo(x, y + r);
  ctx.quadraticCurveTo(x, y, x + r, y);
  ctx.closePath();
}

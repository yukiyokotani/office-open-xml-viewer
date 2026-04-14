import type {
  Slide,
  SlideElement,
  ShapeElement,
  PictureElement,
  TableElement,
  Fill,
  Stroke,
  TextBody,
  Paragraph,
  TextRun,
} from './types';

/** EMU per point (OOXML: 1 pt = 12700 EMU). Used to scale font sizes with the canvas. */
const PT_TO_EMU = 12700;

/**
 * Convert EMU to canvas pixels.
 * scale = canvasWidthPx / slideWidthEMU  (so that slideWidth EMU == canvasWidth px)
 */
function emuToPx(emu: number, scale: number): number {
  return emu * scale;
}

function hexToRgba(hex: string, alpha = 1): string {
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  return `rgba(${r},${g},${b},${alpha})`;
}

function resolveFill(fill: Fill | null): string | null {
  if (!fill || fill.fillType === 'none') return null;
  return hexToRgba(fill.color);
}

// ===== Text layout helpers =====

interface LayoutLine {
  segments: Array<{ text: string; font: string; sizePx: number; color: string; underline: boolean }>;
}

function buildFont(bold: boolean, italic: boolean, sizePx: number, family: string): string {
  const style  = italic ? 'italic ' : '';
  const weight = bold   ? 'bold '   : '';
  return `${style}${weight}${sizePx}px ${family}`;
}

/**
 * Lay out a paragraph into display lines.
 * Handles:
 *  - Explicit line breaks (TextRun type='break')
 *  - Space-based word wrap (Latin text)
 *  - Character-level wrap fallback for CJK / words wider than container
 */
function layoutParagraph(
  ctx: CanvasRenderingContext2D,
  para: Paragraph,
  maxWidthPx: number,
  defaultFontSizePx: number,
  defaultColor: string,
  scale: number
): LayoutLine[] {
  const lines: LayoutLine[] = [];
  let currentLine: LayoutLine = { segments: [] };
  let lineW = 0; // current line's accumulated width

  const newLine = () => {
    lines.push(currentLine);
    currentLine = { segments: [] };
    lineW = 0;
  };

  // Append text to the current line, merging with the last segment when possible
  const push = (text: string, font: string, sizePx: number, color: string, underline: boolean) => {
    if (!text) return;
    ctx.font = font;
    lineW += ctx.measureText(text).width;
    const last = currentLine.segments.at(-1);
    if (last && last.font === font && last.color === color && last.underline === underline) {
      last.text += text;
    } else {
      currentLine.segments.push({ text, font, sizePx, color, underline });
    }
  };

  for (const run of para.runs) {
    if (run.type === 'break') {
      newLine();
      continue;
    }

    const sizePx = run.fontSize != null ? run.fontSize * PT_TO_EMU * scale : defaultFontSizePx;
    const family = run.fontFamily ?? 'sans-serif';
    const color  = run.color ? hexToRgba(run.color) : defaultColor;
    const font   = buildFont(run.bold, run.italic, sizePx, family);
    ctx.font = font;

    // Split on whitespace boundaries, keeping the whitespace tokens
    const tokens = run.text.split(/(\s+)/);

    for (const token of tokens) {
      if (!token) continue;
      ctx.font = font;
      const tokW = ctx.measureText(token).width;
      const isWhitespace = /^\s+$/.test(token);

      if (lineW + tokW <= maxWidthPx) {
        // Token fits on the current line
        push(token, font, sizePx, color, run.underline);
      } else if (isWhitespace) {
        // Whitespace that would overflow → break here, discard leading space
        if (lineW > 0) newLine();
      } else if (tokW > maxWidthPx) {
        // Token is wider than the whole container → character-level wrap
        if (lineW > 0) newLine();
        for (const ch of token) {
          ctx.font = font;
          const chW = ctx.measureText(ch).width;
          if (lineW + chW > maxWidthPx && lineW > 0) newLine();
          push(ch, font, sizePx, color, run.underline);
        }
      } else {
        // Token fits on a fresh line but not on the current one → word wrap
        newLine();
        push(token, font, sizePx, color, run.underline);
      }
    }
  }

  // Always emit the last (possibly empty) line
  lines.push(currentLine);

  return lines;
}

// ===== Element renderers =====

function renderBackground(
  ctx: CanvasRenderingContext2D,
  fill: Fill | null,
  canvasW: number,
  canvasH: number
) {
  ctx.fillStyle = resolveFill(fill) ?? '#FFFFFF';
  ctx.fillRect(0, 0, canvasW, canvasH);
}

function renderShape(ctx: CanvasRenderingContext2D, el: ShapeElement, scale: number) {
  const x = emuToPx(el.x, scale);
  const y = emuToPx(el.y, scale);
  const w = emuToPx(el.width, scale);
  const h = emuToPx(el.height, scale);

  ctx.save();
  if (el.rotation !== 0) {
    ctx.translate(x + w / 2, y + h / 2);
    ctx.rotate((el.rotation * Math.PI) / 180);
    ctx.translate(-(x + w / 2), -(y + h / 2));
  }

  const geom = el.geometry.toLowerCase();
  const fillColor = resolveFill(el.fill);

  ctx.beginPath();
  buildShapePath(ctx, geom, x, y, w, h);

  if (fillColor) {
    ctx.fillStyle = fillColor;
    ctx.fill();
  }
  if (el.stroke) {
    ctx.strokeStyle = hexToRgba(el.stroke.color);
    ctx.lineWidth = Math.max(1, emuToPx(el.stroke.width, scale));
    ctx.stroke();
  }

  ctx.restore();

  if (el.textBody) {
    const defaultTextColor = el.defaultTextColor ? hexToRgba(el.defaultTextColor) : null;
    renderTextBody(ctx, el.textBody, x, y, w, h, scale, defaultTextColor);
  }
}

/** Build the canvas path for a given OOXML preset geometry. */
function buildShapePath(
  ctx: CanvasRenderingContext2D,
  geom: string,
  x: number,
  y: number,
  w: number,
  h: number
) {
  const cx = x + w / 2;
  const cy = y + h / 2;

  switch (geom) {
    case 'ellipse':
    case 'oval':
      ctx.ellipse(cx, cy, w / 2, h / 2, 0, 0, Math.PI * 2);
      break;

    case 'rtriangle':
      ctx.moveTo(x, y + h);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x, y);
      ctx.closePath();
      break;

    case 'triangle':
      ctx.moveTo(cx, y);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;

    case 'diamond':
      ctx.moveTo(cx, y);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(cx, y + h);
      ctx.lineTo(x, cy);
      ctx.closePath();
      break;

    case 'parallelogram': {
      const offset = w * 0.25;
      ctx.moveTo(x + offset, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w - offset, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;
    }

    case 'trapezoid': {
      const inset = w * 0.25;
      ctx.moveTo(x + inset, y);
      ctx.lineTo(x + w - inset, y);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;
    }

    case 'roundrect':
    case 'roundrectangle': {
      const r = Math.min(w, h) / 6;
      ctx.roundRect(x, y, w, h, r);
      break;
    }

    case 'pentagon': {
      for (let i = 0; i < 5; i++) {
        const angle = (i * 2 * Math.PI / 5) - Math.PI / 2;
        if (i === 0) ctx.moveTo(cx + (w / 2) * Math.cos(angle), cy + (h / 2) * Math.sin(angle));
        else         ctx.lineTo(cx + (w / 2) * Math.cos(angle), cy + (h / 2) * Math.sin(angle));
      }
      ctx.closePath();
      break;
    }

    case 'hexagon': {
      for (let i = 0; i < 6; i++) {
        const angle = (i * Math.PI / 3);
        if (i === 0) ctx.moveTo(cx + (w / 2) * Math.cos(angle), cy + (h / 2) * Math.sin(angle));
        else         ctx.lineTo(cx + (w / 2) * Math.cos(angle), cy + (h / 2) * Math.sin(angle));
      }
      ctx.closePath();
      break;
    }

    case 'rightarrow': {
      const arrowHeadW = w * 0.4;
      const shaftH = h * 0.5;
      const shaftY = y + (h - shaftH) / 2;
      ctx.moveTo(x, shaftY);
      ctx.lineTo(x + w - arrowHeadW, shaftY);
      ctx.lineTo(x + w - arrowHeadW, y);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(x + w - arrowHeadW, y + h);
      ctx.lineTo(x + w - arrowHeadW, shaftY + shaftH);
      ctx.lineTo(x, shaftY + shaftH);
      ctx.closePath();
      break;
    }

    case 'leftarrow': {
      const arrowHeadW = w * 0.4;
      const shaftH = h * 0.5;
      const shaftY = y + (h - shaftH) / 2;
      ctx.moveTo(x + w, shaftY);
      ctx.lineTo(x + arrowHeadW, shaftY);
      ctx.lineTo(x + arrowHeadW, y);
      ctx.lineTo(x, cy);
      ctx.lineTo(x + arrowHeadW, y + h);
      ctx.lineTo(x + arrowHeadW, shaftY + shaftH);
      ctx.lineTo(x + w, shaftY + shaftH);
      ctx.closePath();
      break;
    }

    case 'line':
    case 'straightconnector1':
      ctx.moveTo(x, y + h / 2);
      ctx.lineTo(x + w, y + h / 2);
      break;

    default:
      // rect and everything else
      ctx.rect(x, y, w, h);
      break;
  }
}

/** Format an autoNum bullet label from a counter value and OOXML numType. */
function formatAutoNum(counter: number, numType: string): string {
  switch (numType) {
    case 'arabicPeriod':    return `${counter}.`;
    case 'arabicParenR':    return `${counter})`;
    case 'arabicParenBoth': return `(${counter})`;
    case 'alphaLcPeriod':   return `${String.fromCharCode(96 + counter)}.`;
    case 'alphaUcPeriod':   return `${String.fromCharCode(64 + counter)}.`;
    case 'romanLcPeriod':   return `${toRoman(counter).toLowerCase()}.`;
    case 'romanUcPeriod':   return `${toRoman(counter)}.`;
    default:                return `${counter}.`;
  }
}

function toRoman(n: number): string {
  const vals = [1000,900,500,400,100,90,50,40,10,9,5,4,1];
  const syms = ['M','CM','D','CD','C','XC','L','XL','X','IX','V','IV','I'];
  let result = '';
  for (let i = 0; i < vals.length; i++) {
    while (n >= vals[i]) { result += syms[i]; n -= vals[i]; }
  }
  return result;
}

function renderTextBody(
  ctx: CanvasRenderingContext2D,
  body: TextBody,
  bx: number,
  by: number,
  bw: number,
  bh: number,
  scale: number,
  shapeDefaultTextColor: string | null = null
) {
  const lPad = emuToPx(body.lIns, scale);
  const rPad = emuToPx(body.rIns, scale);
  const tPad = emuToPx(body.tIns, scale);
  const bPad = emuToPx(body.bIns, scale);
  const doWrap = body.wrap !== 'none';

  const bodyDefaultFontSizePx = (body.defaultFontSize ?? 18) * PT_TO_EMU * scale;
  const bodyDefaultColor = shapeDefaultTextColor ?? '#000000';

  // ── Pass 1: lay out all paragraphs ──────────────────────────────────────

  interface LineEntry {
    line: LayoutLine;
    linePx: number;       // line height (baseline-to-baseline)
    topGapPx: number;     // spaceBefore for first line of paragraph
    textXOffset: number;  // additional X offset for first-line indent (non-bullet)
    bulletLabel: string;  // text to render as bullet ('' = none)
    bulletFont: string;
    bulletColor: string;
    bulletX: number;      // canvas X for bullet
    textX: number;        // canvas X for text
    textMaxW: number;     // max wrap width
    alignment: string;
    para: Paragraph;
  }

  const allLines: LineEntry[] = [];
  let totalHeight = 0;

  // AutoNum counters per list level
  const autoNumCounters = new Map<number, number>();

  for (const para of body.paragraphs) {
    const marLPx   = emuToPx(para.marL,   scale);
    const marRPx   = emuToPx(para.marR,   scale);
    const indentPx = emuToPx(para.indent, scale);

    // Para-level defaults (cascade: para defRPr → body default)
    const paraDefaultFontSizePx = para.defFontSize != null
      ? para.defFontSize * PT_TO_EMU * scale : bodyDefaultFontSizePx;
    const paraDefaultColor = para.defColor
      ? hexToRgba(para.defColor) : bodyDefaultColor;

    // Bullet resolution
    const hasBullet = para.bullet.type === 'char' || para.bullet.type === 'autoNum';

    let bulletLabel  = '';
    let bulletFont   = buildFont(false, false, paraDefaultFontSizePx, 'sans-serif');
    let bulletColor  = paraDefaultColor;

    if (para.bullet.type === 'char') {
      const b = para.bullet;
      const bSizePx = b.sizePct != null
        ? paraDefaultFontSizePx * (b.sizePct / 100)
        : paraDefaultFontSizePx;
      bulletLabel = b.char;
      bulletFont  = buildFont(false, false, bSizePx, b.fontFamily ?? 'sans-serif');
      bulletColor = b.color ? hexToRgba(b.color) : paraDefaultColor;
      // Reset counters when switching to char bullets
      autoNumCounters.clear();
    } else if (para.bullet.type === 'autoNum') {
      const b = para.bullet;
      const lvl = para.lvl;
      if (!autoNumCounters.has(lvl)) {
        autoNumCounters.set(lvl, b.startAt ?? 1);
      } else {
        autoNumCounters.set(lvl, autoNumCounters.get(lvl)! + 1);
      }
      bulletLabel = formatAutoNum(autoNumCounters.get(lvl)!, b.numType);
      bulletFont  = buildFont(false, false, paraDefaultFontSizePx, 'sans-serif');
      bulletColor = paraDefaultColor;
    } else {
      // Not a list paragraph — reset autoNum counters
      autoNumCounters.clear();
    }

    // Text start X and wrap width
    // For bullet paragraphs: text always at marL, bullet at marL+indent (hanging)
    // For non-bullet with positive indent: first line at marL+indent, others at marL
    const textX    = bx + lPad + marLPx;
    const bulletX  = bx + lPad + marLPx + indentPx;
    const textMaxW = bw - lPad - rPad - marLPx - marRPx;

    const maxW = doWrap ? textMaxW : Infinity;
    const lines = layoutParagraph(ctx, para, maxW, paraDefaultFontSizePx, paraDefaultColor, scale);

    // spaceBefore/After are in hundredths of a point → convert to canvas px
    const spaceBeforePx = para.spaceBefore != null ? (para.spaceBefore / 100) * PT_TO_EMU * scale : 0;
    const spaceAfterPx  = para.spaceAfter  != null ? (para.spaceAfter  / 100) * PT_TO_EMU * scale : 0;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const isFirst = i === 0;
      const isLast  = i === lines.length - 1;

      // Line height: use max font size in the line
      let maxSizePx = paraDefaultFontSizePx;
      for (const seg of line.segments) {
        if (seg.sizePx > maxSizePx) maxSizePx = seg.sizePx;
      }
      // Bullet font size also counts
      if (isFirst && bulletLabel) {
        ctx.font = bulletFont;
        const bm = ctx.measureText('M');
        const bSizeApprox = bm.actualBoundingBoxAscent + bm.actualBoundingBoxDescent;
        if (bSizeApprox > maxSizePx) maxSizePx = bSizeApprox;
      }

      let lineHeight: number;
      if (para.spaceLine) {
        if (para.spaceLine.type === 'pct') {
          lineHeight = maxSizePx * (para.spaceLine.val / 100000);
        } else {
          lineHeight = para.spaceLine.val * PT_TO_EMU * scale;
        }
      } else {
        lineHeight = maxSizePx * 1.2;
      }
      const linePx  = lineHeight + (isLast ? spaceAfterPx : 0);
      const topGap  = isFirst ? spaceBeforePx : 0;
      // Non-bullet first-line indent
      const textXOffset = (!hasBullet && isFirst) ? indentPx : 0;

      allLines.push({
        line, linePx, topGapPx: topGap,
        textXOffset,
        bulletLabel: isFirst ? bulletLabel : '',
        bulletFont, bulletColor, bulletX,
        textX, textMaxW,
        alignment: para.alignment,
        para,
      });
      totalHeight += linePx + topGap;
    }
  }

  // ── Vertical anchor ─────────────────────────────────────────────────────
  let cursorY: number;
  const anchor = body.verticalAnchor ?? 't';
  if (anchor === 'ctr') {
    cursorY = by + (bh - totalHeight) / 2;
  } else if (anchor === 'b') {
    cursorY = by + bh - totalHeight - bPad;
  } else {
    cursorY = by + tPad;
  }

  // ── Pass 2: render ───────────────────────────────────────────────────────
  ctx.save();
  ctx.beginPath();
  ctx.rect(bx, by, bw, bh);
  ctx.clip();

  for (const entry of allLines) {
    const { line, linePx, topGapPx, textXOffset, bulletLabel, bulletFont, bulletColor, bulletX, textX, textMaxW, alignment } = entry;
    cursorY += topGapPx;

    const baseline = cursorY + linePx * 0.8;

    // Draw bullet
    if (bulletLabel) {
      ctx.font = bulletFont;
      ctx.fillStyle = bulletColor;
      ctx.fillText(bulletLabel, bulletX, baseline);
    }

    // Measure line for alignment
    let lineWidth = 0;
    for (const seg of line.segments) {
      ctx.font = seg.font;
      lineWidth += ctx.measureText(seg.text).width;
    }

    const effectiveTextX = textX + textXOffset;
    let penX: number;
    if (alignment === 'ctr') {
      penX = effectiveTextX + (textMaxW - textXOffset - lineWidth) / 2;
    } else if (alignment === 'r') {
      penX = textX + textMaxW - lineWidth;
    } else {
      penX = effectiveTextX;
    }

    for (const seg of line.segments) {
      ctx.font = seg.font;
      ctx.fillStyle = seg.color;
      ctx.fillText(seg.text, penX, baseline);

      if (seg.underline) {
        ctx.font = seg.font;
        const segW = ctx.measureText(seg.text).width;
        ctx.beginPath();
        ctx.moveTo(penX, baseline + 2);
        ctx.lineTo(penX + segW, baseline + 2);
        ctx.strokeStyle = seg.color;
        ctx.lineWidth = 1;
        ctx.stroke();
      }

      ctx.font = seg.font;
      penX += ctx.measureText(seg.text).width;
    }

    cursorY += linePx;
  }

  ctx.restore();
}

async function renderPicture(
  ctx: CanvasRenderingContext2D,
  el: PictureElement,
  scale: number
) {
  return new Promise<void>((resolve) => {
    const img = new Image();
    img.onload = () => {
      ctx.save();
      const x = emuToPx(el.x, scale);
      const y = emuToPx(el.y, scale);
      const w = emuToPx(el.width, scale);
      const h = emuToPx(el.height, scale);
      if (el.rotation !== 0) {
        ctx.translate(x + w / 2, y + h / 2);
        ctx.rotate((el.rotation * Math.PI) / 180);
        ctx.translate(-(x + w / 2), -(y + h / 2));
      }
      ctx.drawImage(img, x, y, w, h);
      ctx.restore();
      resolve();
    };
    img.onerror = () => resolve(); // silently skip broken images
    img.src = el.dataUrl;
  });
}

// ===== Table renderer =====

function applyStroke(ctx: CanvasRenderingContext2D, stroke: Stroke | null, scale: number) {
  if (!stroke) {
    ctx.strokeStyle = 'transparent';
    ctx.lineWidth = 0;
    return;
  }
  ctx.strokeStyle = hexToRgba(stroke.color);
  ctx.lineWidth = Math.max(0.5, emuToPx(stroke.width, scale));
}

function renderTable(ctx: CanvasRenderingContext2D, el: TableElement, scale: number) {
  const x0 = emuToPx(el.x, scale);
  const y0 = emuToPx(el.y, scale);

  // Convert col widths and row heights to pixels
  const colWidths = el.cols.map(c => emuToPx(c, scale));
  const rowHeights = el.rows.map(r => emuToPx(r.height, scale));

  let rowY = y0;
  for (let ri = 0; ri < el.rows.length; ri++) {
    const row = el.rows[ri];
    const rowH = rowHeights[ri];
    let colX = x0;

    for (let ci = 0; ci < row.cells.length; ci++) {
      const cell = row.cells[ci];

      // Merged cells that are continuations: skip drawing
      if (cell.hMerge || cell.vMerge) {
        colX += colWidths[ci] ?? 0;
        continue;
      }

      // Cell size: span multiple columns/rows
      let cellW = 0;
      for (let span = 0; span < (cell.gridSpan || 1); span++) {
        cellW += colWidths[ci + span] ?? 0;
      }
      let cellH = 0;
      for (let span = 0; span < (cell.rowSpan || 1); span++) {
        cellH += rowHeights[ri + span] ?? 0;
      }

      // Fill
      const fillColor = resolveFill(cell.fill);
      if (fillColor) {
        ctx.fillStyle = fillColor;
        ctx.fillRect(colX, rowY, cellW, cellH);
      }

      // Text body
      if (cell.textBody) {
        renderTextBody(ctx, cell.textBody, colX, rowY, cellW, cellH, scale);
      }

      // Borders
      ctx.save();
      if (cell.borderT) {
        applyStroke(ctx, cell.borderT, scale);
        ctx.beginPath();
        ctx.moveTo(colX, rowY);
        ctx.lineTo(colX + cellW, rowY);
        ctx.stroke();
      }
      if (cell.borderB) {
        applyStroke(ctx, cell.borderB, scale);
        ctx.beginPath();
        ctx.moveTo(colX, rowY + cellH);
        ctx.lineTo(colX + cellW, rowY + cellH);
        ctx.stroke();
      }
      if (cell.borderL) {
        applyStroke(ctx, cell.borderL, scale);
        ctx.beginPath();
        ctx.moveTo(colX, rowY);
        ctx.lineTo(colX, rowY + cellH);
        ctx.stroke();
      }
      if (cell.borderR) {
        applyStroke(ctx, cell.borderR, scale);
        ctx.beginPath();
        ctx.moveTo(colX + cellW, rowY);
        ctx.lineTo(colX + cellW, rowY + cellH);
        ctx.stroke();
      }
      ctx.restore();

      colX += colWidths[ci] ?? 0;
    }
    rowY += rowH;
  }
}

// ===== Public API =====

export interface RenderOptions {
  /** Target canvas width in CSS pixels (height is computed from slide aspect ratio) */
  width?: number;
}

/**
 * Render a single slide onto a <canvas> element.
 * Returns the canvas for convenience.
 */
export async function renderSlide(
  canvas: HTMLCanvasElement,
  slide: Slide,
  slideWidth: number,
  slideHeight: number,
  opts: RenderOptions = {}
): Promise<HTMLCanvasElement> {
  const targetWidth = opts.width ?? (canvas.offsetWidth || 960);
  const scale = targetWidth / slideWidth;
  const canvasW = Math.round(targetWidth);
  const canvasH = Math.round(slideHeight * scale);

  // Use devicePixelRatio for crisp rendering on HiDPI screens
  const dpr = window.devicePixelRatio || 1;
  canvas.width  = canvasW * dpr;
  canvas.height = canvasH * dpr;
  canvas.style.width  = `${canvasW}px`;
  canvas.style.height = `${canvasH}px`;

  const ctx = canvas.getContext('2d');
  if (!ctx) throw new Error('Could not get 2D context');
  ctx.scale(dpr, dpr);

  renderBackground(ctx, slide.background, canvasW, canvasH);

  for (const el of slide.elements) {
    if (el.type === 'shape') {
      renderShape(ctx, el, scale);
    } else if (el.type === 'picture') {
      await renderPicture(ctx, el, scale);
    } else if (el.type === 'table') {
      renderTable(ctx, el, scale);
    }
  }

  return canvas;
}

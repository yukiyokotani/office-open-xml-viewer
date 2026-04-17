import type {
  Slide,
  SlideElement,
  ShapeElement,
  PictureElement,
  TableElement,
  ChartElement,
  ChartSeriesData,
  Fill,
  Stroke,
  TextBody,
  Paragraph,
  TextRun,
  PathCmd,
  Shadow,
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
  // 8-char hex (RRGGBBAA) encodes alpha in the last two chars
  const a = hex.length >= 8 ? parseInt(hex.slice(6, 8), 16) / 255 : alpha;
  return `rgba(${r},${g},${b},${a})`;
}

/** Simple fill resolver that returns a CSS color string.
 *  For gradient fills, returns the first stop's color (used by table cells etc.) */
function resolveFill(fill: Fill | null): string | null {
  if (!fill || fill.fillType === 'none') return null;
  if (fill.fillType === 'solid') return hexToRgba(fill.color);
  if (fill.fillType === 'gradient') {
    return fill.stops.length > 0 ? hexToRgba(fill.stops[0].color) : null;
  }
  return null;
}

/** Context-aware fill resolver that creates a CanvasGradient for gradient fills. */
function resolveShapeFill(
  fill: Fill | null,
  ctx: CanvasRenderingContext2D,
  x: number, y: number, w: number, h: number
): string | CanvasGradient | null {
  if (!fill || fill.fillType === 'none') return null;
  if (fill.fillType === 'solid') return hexToRgba(fill.color);
  if (fill.fillType === 'gradient') {
    const stops = fill.stops;
    if (stops.length === 0) return null;
    if (stops.length === 1) return hexToRgba(stops[0].color);

    let gradient: CanvasGradient;
    if (fill.gradType === 'radial') {
      const cx = x + w / 2;
      const cy = y + h / 2;
      const r = Math.sqrt(w * w + h * h) / 2;
      gradient = ctx.createRadialGradient(cx, cy, 0, cx, cy, r);
    } else {
      // Linear gradient — OOXML angle: 0 = left→right, 90 = top→bottom
      const rad = (fill.angle * Math.PI) / 180;
      const cx = x + w / 2;
      const cy = y + h / 2;
      // Compute extent so the gradient line covers the entire bounding box
      const gradLen = (Math.abs(Math.cos(rad)) * w + Math.abs(Math.sin(rad)) * h) / 2;
      gradient = ctx.createLinearGradient(
        cx - Math.cos(rad) * gradLen, cy - Math.sin(rad) * gradLen,
        cx + Math.cos(rad) * gradLen, cy + Math.sin(rad) * gradLen
      );
    }

    for (const stop of stops) {
      gradient.addColorStop(Math.min(1, Math.max(0, stop.position)), hexToRgba(stop.color));
    }
    return gradient;
  }
  return null;
}

// ===== Text layout helpers =====

type LayoutSegment = { text: string; font: string; sizePx: number; color: string; underline: boolean; strikethrough: boolean };

interface LayoutLine {
  segments: LayoutSegment[];
  /** Segments right-aligned at a tab stop (set when paragraph contains \t and a right-aligned tabStop) */
  tabStop?: {
    /** Tab stop position in px from the left edge of the text area (bx + lPad + tabStop.px = canvas X) */
    px: number;
    algn: string;
    segments: LayoutSegment[];
  };
}

/**
 * Resolve OOXML theme font references (e.g. "+mn-ea", "+mj-lt") to CSS-safe font names.
 * Canvas will silently ignore an invalid CSS font string, keeping whatever font was set before —
 * leading to wrong text size. Map theme references to generic families as a safe fallback.
 */
const WINGDINGS_MAP: Record<number, string> = {
  0x21: '✏', 0x22: '✂', 0x23: '✁', 0x24: '👁',
  0x4A: '☺', 0x4B: '☻', 0x4C: '☹',
  0x76: '✔', 0xFC: '✓', 0xFB: '✗', 0xFE: '■',
  0xA7: '▪', 0xB7: '•', 0xB8: '◦', 0xB9: '–',
  0xF0A7: '▪', 0xF0B7: '•',
};

function applySymbolFont(char: string, fontFamily: string): string {
  const lower = fontFamily.toLowerCase();
  if (lower.includes('wingdings') || lower === 'symbol') {
    const code = char.charCodeAt(0);
    return WINGDINGS_MAP[code] ?? char;
  }
  return char;
}

function normalizeFontFamily(family: string): string {
  if (!family || family.startsWith('+')) {
    // +mn-lt = minor Latin, +mj-lt = major Latin, +mn-ea = minor East Asian, +mj-ea = major East Asian
    // All fall back to system sans-serif (close enough for layout purposes)
    return 'sans-serif';
  }
  return family;
}

/** CSS generic font families — must NOT be quoted in a canvas font string. */
const CSS_GENERIC_FAMILIES = new Set([
  'serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui',
]);

function buildFont(bold: boolean, italic: boolean, sizePx: number, family: string): string {
  const style  = italic ? 'italic ' : '';
  const weight = bold   ? 'bold '   : '';
  const normalized = normalizeFontFamily(family);
  // Generic families must be unquoted; named families must be quoted.
  const quotedFamily = CSS_GENERIC_FAMILIES.has(normalized) ? normalized : `"${normalized}"`;
  return `${style}${weight}${sizePx}px ${quotedFamily}`;
}

/**
 * Lay out a paragraph into display lines.
 * Handles:
 *  - Explicit line breaks (TextRun type='break')
 *  - Space-based word wrap (Latin text)
 *  - Character-level wrap fallback for CJK / words wider than container
 *  - Tab stops (right-aligned and left-aligned)
 *
 * @param marLPx  Paragraph left margin in canvas px (used for tab stop position calculation)
 */
function layoutParagraph(
  ctx: CanvasRenderingContext2D,
  para: Paragraph,
  maxWidthPx: number,
  defaultFontSizePx: number,
  defaultColor: string,
  scale: number,
  marLPx: number,
  defaultBold: boolean = false,
  defaultItalic: boolean = false,
  fontScale: number = 1.0,
): LayoutLine[] {
  const lines: LayoutLine[] = [];
  let currentLine: LayoutLine = { segments: [] };
  let lineW = 0; // current line's accumulated width

  // Tab stop state: once we hit a \t we switch to collecting tabStop.segments
  let tabActive = false;
  let tabStopPx = 0;   // position of tab stop from text area left (px)

  const newLine = () => {
    lines.push(currentLine);
    currentLine = { segments: [] };
    lineW = 0;
    tabActive = false; // reset tab state per line
  };

  // Push to the active segment list (main or tab-stop group)
  const push = (text: string, font: string, sizePx: number, color: string, underline: boolean, strikethrough: boolean) => {
    if (!text) return;
    ctx.font = font;
    const w = ctx.measureText(text).width;
    if (tabActive && currentLine.tabStop) {
      const segs = currentLine.tabStop.segments;
      const last = segs.at(-1);
      if (last && last.font === font && last.color === color && last.underline === underline && last.strikethrough === strikethrough) {
        last.text += text;
      } else {
        segs.push({ text, font, sizePx, color, underline, strikethrough });
      }
    } else {
      lineW += w;
      const last = currentLine.segments.at(-1);
      if (last && last.font === font && last.color === color && last.underline === underline && last.strikethrough === strikethrough) {
        last.text += text;
      } else {
        currentLine.segments.push({ text, font, sizePx, color, underline, strikethrough });
      }
    }
  };

  for (const run of para.runs) {
    if (run.type === 'break') {
      newLine();
      continue;
    }

    const sizePx = run.fontSize != null ? run.fontSize * PT_TO_EMU * scale * fontScale : defaultFontSizePx;
    const family = normalizeFontFamily(run.fontFamily ?? 'sans-serif');
    const color  = run.color ? hexToRgba(run.color) : defaultColor;
    // Cascade: run → paragraph defRPr → body/layout default → false
    const isBold   = run.bold   ?? para.defBold   ?? defaultBold;
    const isItalic = run.italic ?? para.defItalic ?? defaultItalic;
    const font   = buildFont(isBold, isItalic, sizePx, family);
    ctx.font = font;

    // Split on whitespace boundaries, keeping the whitespace tokens
    const tokens = run.text.split(/(\s+)/);

    for (const token of tokens) {
      if (!token) continue;

      // ── Tab character ────────────────────────────────────────────────────
      if (/^\t+$/.test(token)) {
        // Find first tab stop whose position (from text area left) is beyond the current pen
        const currentAbsW = marLPx + lineW; // current position from text area left
        const ts = (para.tabStops ?? []).find(
          t => emuToPx(t.pos, scale) > currentAbsW
        );
        if (ts) {
          tabStopPx = emuToPx(ts.pos, scale);
          if (ts.algn === 'r' || ts.algn === 'ctr') {
            // Switch to tab-stop accumulation mode
            tabActive = true;
            currentLine.tabStop = { px: tabStopPx, algn: ts.algn, segments: [] };
          } else {
            // Left-aligned tab: advance lineW to the tab stop
            lineW = tabStopPx - marLPx;
          }
        } else {
          // No matching tab stop — treat as a single space
          push(' ', font, sizePx, color, run.underline, run.strikethrough);
        }
        continue;
      }

      ctx.font = font;
      const tokW = ctx.measureText(token).width;
      const isWhitespace = /^\s+$/.test(token);

      // If already in tab mode, collect all text into tabStop.segments (no wrap)
      if (tabActive) {
        push(token, font, sizePx, color, run.underline, run.strikethrough);
        continue;
      }

      // CJK characters allow line-breaking at any character boundary (no whitespace
      // needed). When a token contains CJK, wrap character-by-character so that CJK
      // text flows onto the same line as preceding Latin text (e.g. "EC市場で…").
      const hasCJK = /[\u3000-\u9FFF\uAC00-\uD7FF\uF900-\uFAFF\uFF00-\uFFEF]/.test(token);
      if (hasCJK) {
        for (const ch of token) {
          ctx.font = font;
          const chW = ctx.measureText(ch).width;
          if (lineW + chW > maxWidthPx && lineW > 0) newLine();
          push(ch, font, sizePx, color, run.underline, run.strikethrough);
        }
        continue;
      }

      if (lineW + tokW <= maxWidthPx) {
        push(token, font, sizePx, color, run.underline, run.strikethrough);
      } else if (isWhitespace) {
        if (lineW > 0) newLine();
      } else if (tokW > maxWidthPx) {
        if (lineW > 0) newLine();
        for (const ch of token) {
          ctx.font = font;
          const chW = ctx.measureText(ch).width;
          if (lineW + chW > maxWidthPx && lineW > 0) newLine();
          push(ch, font, sizePx, color, run.underline, run.strikethrough);
        }
      } else {
        newLine();
        push(token, font, sizePx, color, run.underline, run.strikethrough);
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
  const bg = resolveShapeFill(fill, ctx, 0, 0, canvasW, canvasH);
  ctx.fillStyle = bg ?? '#FFFFFF';
  ctx.fillRect(0, 0, canvasW, canvasH);
}

function applyShadow(ctx: CanvasRenderingContext2D, shadow: Shadow | null, scale: number) {
  if (!shadow) return;
  const dirRad = (shadow.dir * Math.PI) / 180;
  const dist = emuToPx(shadow.dist, scale);
  ctx.shadowColor = hexToRgba(shadow.color, shadow.alpha);
  ctx.shadowBlur = emuToPx(shadow.blur, scale);
  ctx.shadowOffsetX = Math.cos(dirRad) * dist;
  ctx.shadowOffsetY = Math.sin(dirRad) * dist;
}

function clearShadow(ctx: CanvasRenderingContext2D) {
  ctx.shadowColor = 'transparent';
  ctx.shadowBlur = 0;
  ctx.shadowOffsetX = 0;
  ctx.shadowOffsetY = 0;
}

function renderShape(ctx: CanvasRenderingContext2D, el: ShapeElement, scale: number, themeDefaultColor = '#000000') {
  const x = emuToPx(el.x, scale);
  const y = emuToPx(el.y, scale);
  const w = emuToPx(el.width, scale);
  const h = emuToPx(el.height, scale);

  // anchor="b" + h=0: shape grows upward from y; render stroke as bottom border,
  // then let renderTextBody handle positioning.
  if (h === 0 && el.textBody?.verticalAnchor === 'b') {
    if (el.stroke) {
      ctx.save();
      ctx.strokeStyle = hexToRgba(el.stroke.color);
      ctx.lineWidth = Math.max(1, emuToPx(el.stroke.width, scale));
      ctx.beginPath();
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.stroke();
      ctx.restore();
    }
    if (el.textBody) {
      const defaultTextColor = el.defaultTextColor ? hexToRgba(el.defaultTextColor) : null;
      renderTextBody(ctx, el.textBody, x, y, w, h, scale, defaultTextColor, el.rotation, el.flipH, el.flipV, themeDefaultColor);
    }
    return;
  }

  ctx.save();
  if (el.rotation !== 0 || el.flipH || el.flipV) {
    ctx.translate(x + w / 2, y + h / 2);
    ctx.rotate((el.rotation * Math.PI) / 180);
    if (el.flipH) ctx.scale(-1, 1);
    if (el.flipV) ctx.scale(1, -1);
    ctx.translate(-(x + w / 2), -(y + h / 2));
  }

  const geom = el.geometry.toLowerCase();
  const fillStyle = resolveShapeFill(el.fill, ctx, x, y, w, h);

  // Apply shadow before fill/stroke drawing; ctx.restore() will clear it
  applyShadow(ctx, el.shadow ?? null, scale);

  ctx.beginPath();
  if (el.custGeom && el.custGeom.length > 0) {
    buildCustomPath(ctx, el.custGeom, x, y, w, h);
  } else {
    buildShapePath(ctx, geom, x, y, w, h, el.adj, el.adj2);
  }

  if (fillStyle) {
    ctx.fillStyle = fillStyle;
    // donut/noSmoking use evenodd winding to create a hole
    if (geom === 'donut' || geom === 'nosmokingsign') {
      ctx.fill('evenodd');
    } else {
      ctx.fill();
    }
    // Clear shadow after fill so stroke/text don't double-shadow
    clearShadow(ctx);
  }
  if (el.stroke) {
    ctx.strokeStyle = hexToRgba(el.stroke.color);
    ctx.lineWidth = Math.max(1, emuToPx(el.stroke.width, scale));
    ctx.stroke();
  }

  ctx.restore();

  if (el.textBody) {
    const defaultTextColor = el.defaultTextColor ? hexToRgba(el.defaultTextColor) : null;
    renderTextBody(ctx, el.textBody, x, y, w, h, scale, defaultTextColor, el.rotation, el.flipH, el.flipV, themeDefaultColor);
  }
}

/**
 * Build a canvas path from custGeom path commands.
 * Coordinates are in [0,1] relative to the shape bounding box;
 * the renderer maps them to canvas pixels.
 * Tracks pen position so arcTo can compute the ellipse centre correctly.
 */
function buildCustomPath(
  ctx: CanvasRenderingContext2D,
  subpaths: PathCmd[][],
  x: number,
  y: number,
  w: number,
  h: number
) {
  for (const cmds of subpaths) {
    // Pen position in normalised [0,1] space
    let penX = 0, penY = 0;
    for (const cmd of cmds) {
      switch (cmd.cmd) {
        case 'moveTo':
          ctx.moveTo(x + cmd.x * w, y + cmd.y * h);
          penX = cmd.x; penY = cmd.y;
          break;
        case 'lineTo':
          ctx.lineTo(x + cmd.x * w, y + cmd.y * h);
          penX = cmd.x; penY = cmd.y;
          break;
        case 'cubicBezTo':
          ctx.bezierCurveTo(
            x + cmd.x1 * w, y + cmd.y1 * h,
            x + cmd.x2 * w, y + cmd.y2 * h,
            x + cmd.x  * w, y + cmd.y  * h
          );
          penX = cmd.x; penY = cmd.y;
          break;
        case 'arcTo': {
          // OOXML arcTo: the current pen is on the ellipse at angle stAng.
          // Back-calculate the ellipse centre, then draw to stAng+swAng.
          const rw = cmd.wr * w;
          const rh = cmd.hr * h;
          if (rw <= 0 || rh <= 0) break;
          const stRad = (cmd.stAng * Math.PI) / 180;
          const swRad = (cmd.swAng * Math.PI) / 180;
          const penAbsX = x + penX * w;
          const penAbsY = y + penY * h;
          // Centre of the ellipse that passes through the current pen at stAng
          const cx = penAbsX - rw * Math.cos(stRad);
          const cy = penAbsY - rh * Math.sin(stRad);
          const endRad = stRad + swRad;
          ctx.ellipse(cx, cy, rw, rh, 0, stRad, endRad, swRad < 0);
          // Update pen to the arc end point
          penX = (cx + rw * Math.cos(endRad) - x) / w;
          penY = (cy + rh * Math.sin(endRad) - y) / h;
          break;
        }
        case 'close':
          ctx.closePath();
          break;
      }
    }
  }
}

// ── Star polygon helper ─────────────────────────────────────────────────────
function drawStar(
  ctx: CanvasRenderingContext2D,
  cx: number, cy: number,
  rx: number, ry: number,
  points: number,
  innerRatio: number,
  startAngle = -Math.PI / 2
) {
  const total = points * 2;
  for (let i = 0; i < total; i++) {
    const angle = startAngle + (i * Math.PI) / points;
    const r = i % 2 === 0 ? 1.0 : innerRatio;
    const px = cx + rx * r * Math.cos(angle);
    const py = cy + ry * r * Math.sin(angle);
    if (i === 0) ctx.moveTo(px, py);
    else ctx.lineTo(px, py);
  }
  ctx.closePath();
}

// ── Regular polygon helper ───────────────────────────────────────────────────
function drawPolygon(
  ctx: CanvasRenderingContext2D,
  cx: number, cy: number,
  rx: number, ry: number,
  sides: number,
  startAngle = -Math.PI / 2
) {
  for (let i = 0; i < sides; i++) {
    const angle = startAngle + (i * 2 * Math.PI) / sides;
    const px = cx + rx * Math.cos(angle);
    const py = cy + ry * Math.sin(angle);
    if (i === 0) ctx.moveTo(px, py);
    else ctx.lineTo(px, py);
  }
  ctx.closePath();
}

/** Build the canvas path for a given OOXML preset geometry.
 * @param adj  First adjustment value from avLst (0–100000 range), used by shapes like trapezoid.
 */
function buildShapePath(
  ctx: CanvasRenderingContext2D,
  geom: string,
  x: number,
  y: number,
  w: number,
  h: number,
  adj: number | null = null,
  adj2: number | null = null,
) {
  const cx = x + w / 2;
  const cy = y + h / 2;

  switch (geom) {
    // ── Ellipses ──────────────────────────────────────────────────────────────
    case 'ellipse':
    case 'oval':
      ctx.ellipse(cx, cy, w / 2, h / 2, 0, 0, Math.PI * 2);
      break;

    // ── Triangles ─────────────────────────────────────────────────────────────
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

    // ── Quadrilaterals ────────────────────────────────────────────────────────
    case 'diamond':
      ctx.moveTo(cx, y);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(cx, y + h);
      ctx.lineTo(x, cy);
      ctx.closePath();
      break;

    case 'parallelogram': {
      // adj controls horizontal slant; default 25000 = 25% of width
      const offset = w * Math.min(0.5, (adj ?? 25000) / 100000);
      ctx.moveTo(x + offset, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w - offset, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;
    }

    case 'trapezoid': {
      const ss = Math.min(w, h);
      const inset = Math.min(w / 2, (adj ?? 25000) / 100000 * ss);
      ctx.moveTo(x + inset, y);
      ctx.lineTo(x + w - inset, y);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;
    }

    case 'roundrect':
    case 'roundrectangle': {
      // OOXML: circular corners — r = min(w,h) * adj / 100000
      // adj range 0–50000 (default 16667); at adj=50000 r=min(w,h)/2 (stadium)
      const a = Math.min(50000, Math.max(0, adj ?? 16667));
      const r = Math.min(w, h) * a / 100000;
      ctx.roundRect(x, y, w, h, r);
      break;
    }

    // ── Regular polygons ──────────────────────────────────────────────────────
    case 'pentagon':
      drawPolygon(ctx, cx, cy, w / 2, h / 2, 5);
      break;
    case 'hexagon':
      drawPolygon(ctx, cx, cy, w / 2, h / 2, 6, 0);
      break;
    case 'heptagon':
      drawPolygon(ctx, cx, cy, w / 2, h / 2, 7);
      break;
    case 'octagon':
      drawPolygon(ctx, cx, cy, w / 2, h / 2, 8, -Math.PI / 8);
      break;
    case 'decagon':
      drawPolygon(ctx, cx, cy, w / 2, h / 2, 10);
      break;
    case 'dodecagon':
      drawPolygon(ctx, cx, cy, w / 2, h / 2, 12);
      break;

    // ── Stars ─────────────────────────────────────────────────────────────────
    case 'star4':
      drawStar(ctx, cx, cy, w / 2, h / 2, 4, 0.38);
      break;
    case 'star5':
    case 'star':
      drawStar(ctx, cx, cy, w / 2, h / 2, 5, 0.382);
      break;
    case 'star6':
      drawStar(ctx, cx, cy, w / 2, h / 2, 6, 0.5, 0);
      break;
    case 'star7':
      drawStar(ctx, cx, cy, w / 2, h / 2, 7, 0.37);
      break;
    case 'star8':
      drawStar(ctx, cx, cy, w / 2, h / 2, 8, 0.38, -Math.PI / 8);
      break;
    case 'star10':
      drawStar(ctx, cx, cy, w / 2, h / 2, 10, 0.45);
      break;
    case 'star12':
      drawStar(ctx, cx, cy, w / 2, h / 2, 12, 0.45, 0);
      break;
    case 'star16':
      drawStar(ctx, cx, cy, w / 2, h / 2, 16, 0.55, -Math.PI / 16);
      break;
    case 'star24':
      drawStar(ctx, cx, cy, w / 2, h / 2, 24, 0.6, 0);
      break;
    case 'star32':
      drawStar(ctx, cx, cy, w / 2, h / 2, 32, 0.65, -Math.PI / 32);
      break;

    // ── Arrows ────────────────────────────────────────────────────────────────
    case 'rightarrow': {
      // adj1=shaft height (% of h, default 50000), adj2=arrowhead from right (% of w, default 50000)
      const sh = h * Math.min(1, (adj  ?? 50000) / 100000);
      const ahw = w * Math.min(1, (adj2 ?? 50000) / 100000);
      const sy = y + (h - sh) / 2;
      ctx.moveTo(x, sy);
      ctx.lineTo(x + w - ahw, sy);
      ctx.lineTo(x + w - ahw, y);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(x + w - ahw, y + h);
      ctx.lineTo(x + w - ahw, sy + sh);
      ctx.lineTo(x, sy + sh);
      ctx.closePath();
      break;
    }
    case 'leftarrow': {
      const sh = h * Math.min(1, (adj  ?? 50000) / 100000);
      const ahw = w * Math.min(1, (adj2 ?? 50000) / 100000);
      const sy = y + (h - sh) / 2;
      ctx.moveTo(x + w, sy);
      ctx.lineTo(x + ahw, sy);
      ctx.lineTo(x + ahw, y);
      ctx.lineTo(x, cy);
      ctx.lineTo(x + ahw, y + h);
      ctx.lineTo(x + ahw, sy + sh);
      ctx.lineTo(x + w, sy + sh);
      ctx.closePath();
      break;
    }
    case 'uparrow': {
      const sw = w * Math.min(1, (adj  ?? 50000) / 100000);
      const ahh = h * Math.min(1, (adj2 ?? 50000) / 100000);
      const sx = x + (w - sw) / 2;
      ctx.moveTo(cx, y);
      ctx.lineTo(x + w, y + ahh);
      ctx.lineTo(sx + sw, y + ahh);
      ctx.lineTo(sx + sw, y + h);
      ctx.lineTo(sx, y + h);
      ctx.lineTo(sx, y + ahh);
      ctx.lineTo(x, y + ahh);
      ctx.closePath();
      break;
    }
    case 'downarrow': {
      const sw = w * Math.min(1, (adj  ?? 50000) / 100000);
      const ahh = h * Math.min(1, (adj2 ?? 50000) / 100000);
      const sx = x + (w - sw) / 2;
      ctx.moveTo(cx, y + h);
      ctx.lineTo(x + w, y + h - ahh);
      ctx.lineTo(sx + sw, y + h - ahh);
      ctx.lineTo(sx + sw, y);
      ctx.lineTo(sx, y);
      ctx.lineTo(sx, y + h - ahh);
      ctx.lineTo(x, y + h - ahh);
      ctx.closePath();
      break;
    }
    case 'leftrightarrow': {
      const sh = h * Math.min(1, (adj  ?? 50000) / 100000);
      const ahw = w * Math.min(0.5, (adj2 ?? 25000) / 100000);
      const sy = y + (h - sh) / 2;
      ctx.moveTo(x, cy);
      ctx.lineTo(x + ahw, y);
      ctx.lineTo(x + ahw, sy);
      ctx.lineTo(x + w - ahw, sy);
      ctx.lineTo(x + w - ahw, y);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(x + w - ahw, y + h);
      ctx.lineTo(x + w - ahw, sy + sh);
      ctx.lineTo(x + ahw, sy + sh);
      ctx.lineTo(x + ahw, y + h);
      ctx.closePath();
      break;
    }
    case 'updownarrow': {
      const sw = w * Math.min(1, (adj  ?? 50000) / 100000);
      const ahh = h * Math.min(0.5, (adj2 ?? 25000) / 100000);
      const sx = x + (w - sw) / 2;
      ctx.moveTo(cx, y);
      ctx.lineTo(x + w, y + ahh);
      ctx.lineTo(sx + sw, y + ahh);
      ctx.lineTo(sx + sw, y + h - ahh);
      ctx.lineTo(x + w, y + h - ahh);
      ctx.lineTo(cx, y + h);
      ctx.lineTo(x, y + h - ahh);
      ctx.lineTo(sx, y + h - ahh);
      ctx.lineTo(sx, y + ahh);
      ctx.lineTo(x, y + ahh);
      ctx.closePath();
      break;
    }
    case 'notchedrightarrow': {
      const sh = h * Math.min(1, (adj  ?? 50000) / 100000);
      const ahw = w * Math.min(1, (adj2 ?? 35000) / 100000);
      const sy = y + (h - sh) / 2;
      const notch = ahw * 0.43; // notch depth relative to arrowhead width
      ctx.moveTo(x, sy);
      ctx.lineTo(x + w - ahw, sy);
      ctx.lineTo(x + w - ahw, y);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(x + w - ahw, y + h);
      ctx.lineTo(x + w - ahw, sy + sh);
      ctx.lineTo(x, sy + sh);
      ctx.lineTo(x + notch, cy);
      ctx.closePath();
      break;
    }

    // ── Process flow shapes ───────────────────────────────────────────────────
    case 'chevron': {
      // adj = kink position from left as fraction of width; default 50000 (50%)
      // Kink at x=kink: right arrow-tip spans from kink to w; left V-notch at kink
      const kink = w * Math.min(1, Math.max(0, (adj ?? 50000) / 100000));
      ctx.moveTo(x, y);
      ctx.lineTo(x + kink, y);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(x + kink, y + h);
      ctx.lineTo(x, y + h);
      if (kink > 0) ctx.lineTo(x + kink, cy);
      ctx.closePath();
      break;
    }
    case 'homeplate': {
      const tip = h * 0.4;
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + h - tip);
      ctx.lineTo(cx, y + h);
      ctx.lineTo(x, y + h - tip);
      ctx.closePath();
      break;
    }

    // ── Brackets / braces ─────────────────────────────────────────────────────
    case 'leftbracket': {
      // Square bracket [ shape. adj (default 8333) controls corner arc height
      // as fraction of h; clamp to [0, 50000] per OOXML spec.
      const a = Math.min(50000, Math.max(0, adj ?? 8333));
      const arcH2 = Math.min(h * a / 100000, h / 2); // never let arcs overlap
      // Top arc: (w, 0) → quadratic via (0, 0) → (0, arcH)
      ctx.moveTo(x + w, y);
      ctx.quadraticCurveTo(x, y, x, y + arcH2);
      // Straight left side — omit when arcs just meet (path continues from arc end)
      if (h - 2 * arcH2 > 0.5) ctx.lineTo(x, y + h - arcH2);
      // Bottom arc: (0, h-arcH) → quadratic via (0, h) → (w, h)
      ctx.quadraticCurveTo(x, y + h, x + w, y + h);
      break;
    }
    case 'rightbracket': {
      // Square bracket ] shape — mirror of leftBracket.
      const a = Math.min(50000, Math.max(0, adj ?? 8333));
      const arcH2 = Math.min(h * a / 100000, h / 2);
      ctx.moveTo(x, y);
      ctx.quadraticCurveTo(x + w, y, x + w, y + arcH2);
      if (h - 2 * arcH2 > 0.5) ctx.lineTo(x + w, y + h - arcH2);
      ctx.quadraticCurveTo(x + w, y + h, x, y + h);
      break;
    }
    case 'leftbrace': {
      // { shape
      const mid = cy;
      const nb = w * 0.45;
      ctx.moveTo(x + w, y);
      ctx.bezierCurveTo(x + w - nb, y, x + w - nb, mid - h * 0.08, x, mid);
      ctx.bezierCurveTo(x + w - nb, mid + h * 0.08, x + w - nb, y + h, x + w, y + h);
      break;
    }
    case 'rightbrace': {
      const mid = cy;
      const nb = w * 0.45;
      ctx.moveTo(x, y);
      ctx.bezierCurveTo(x + nb, y, x + nb, mid - h * 0.08, x + w, mid);
      ctx.bezierCurveTo(x + nb, mid + h * 0.08, x + nb, y + h, x, y + h);
      break;
    }

    // ── Callouts ──────────────────────────────────────────────────────────────
    case 'wedgerectcallout':
    case 'callout1':
    case 'callout2':
    case 'callout3':
    case 'bordercallout1':
    case 'bordercallout2':
    case 'bordercallout3': {
      // Simplified: rect with a small triangular pointer at the bottom-left
      ctx.rect(x, y, w, h * 0.8);
      const tipX = x + w * 0.2;
      const tipY = y + h;
      ctx.moveTo(x + w * 0.1, y + h * 0.8);
      ctx.lineTo(tipX, tipY);
      ctx.lineTo(x + w * 0.3, y + h * 0.8);
      ctx.closePath();
      break;
    }

    // ── Connectors ────────────────────────────────────────────────────────────
    case 'line':
    case 'straightconnector1':
    case 'bentconnector2':
    case 'bentconnector3':
    case 'bentconnector4':
    case 'bentconnector5':
    case 'curvedconnector2':
    case 'curvedconnector3':
    case 'curvedconnector4':
    case 'curvedconnector5':
      // Connectors run diagonally from top-left to bottom-right of their bounding box.
      // Flip transforms (already applied to ctx) handle other orientations.
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y + h);
      break;

    // ── Heart ─────────────────────────────────────────────────────────────────
    case 'heart': {
      ctx.moveTo(cx, y + h * 0.32);
      ctx.bezierCurveTo(cx, y, x + w * 0.05, y, x, y + h * 0.3);
      ctx.bezierCurveTo(x, y + h * 0.68, cx - w * 0.05, y + h * 0.78, cx, y + h);
      ctx.bezierCurveTo(cx + w * 0.05, y + h * 0.78, x + w, y + h * 0.68, x + w, y + h * 0.3);
      ctx.bezierCurveTo(x + w - w * 0.05, y, cx, y, cx, y + h * 0.32);
      break;
    }

    // ── Donut / ring ──────────────────────────────────────────────────────────
    case 'donut': {
      ctx.arc(cx, cy, w / 2, 0, Math.PI * 2);
      const ir = Math.min(w, h) * (adj != null ? (adj / 100000) * 0.5 : 0.25);
      ctx.moveTo(cx + ir, cy);
      ctx.arc(cx, cy, ir, 0, Math.PI * 2, true);
      break;
    }

    // ── No smoking / prohibition sign ─────────────────────────────────────────
    // OOXML spec: outer ellipse ring (evenodd hole) + two diagonal arc segments
    case 'nosmoking':
    case 'nosmokingsign': {
      const iwd2 = w / 2 * (1 - (adj ?? 18750) / 100000 * 2);
      const ihd2 = h / 2 * (1 - (adj ?? 18750) / 100000 * 2);
      // outer ring: outer circle CCW + inner circle CW (evenodd creates ring)
      ctx.arc(cx, cy, w / 2, 0, Math.PI * 2);
      ctx.moveTo(cx + iwd2, cy);
      ctx.arc(cx, cy, iwd2, 0, Math.PI * 2, true);
      // diagonal bar: two thick arc segments that form a slash
      // stAng1 ≈ 225°, stAng2 ≈ 45°, swAng ≈ calculated from ring width
      const ang = Math.atan2(ihd2, iwd2);
      const stAng1 = Math.PI + ang;
      const stAng2 = ang;
      const swAng  = Math.PI - 2 * ang;
      // first arc segment (upper-left to lower-right chord, outer radius)
      ctx.moveTo(cx + (w / 2) * Math.cos(stAng1), cy + (h / 2) * Math.sin(stAng1));
      ctx.ellipse(cx, cy, w / 2, h / 2, 0, stAng1, stAng1 + swAng);
      ctx.lineTo(cx + iwd2 * Math.cos(stAng1 + swAng), cy + ihd2 * Math.sin(stAng1 + swAng));
      ctx.ellipse(cx, cy, iwd2, ihd2, 0, stAng1 + swAng, stAng1, true);
      ctx.closePath();
      // second arc segment (lower-left to upper-right)
      ctx.moveTo(cx + (w / 2) * Math.cos(stAng2), cy + (h / 2) * Math.sin(stAng2));
      ctx.ellipse(cx, cy, w / 2, h / 2, 0, stAng2, stAng2 + swAng);
      ctx.lineTo(cx + iwd2 * Math.cos(stAng2 + swAng), cy + ihd2 * Math.sin(stAng2 + swAng));
      ctx.ellipse(cx, cy, iwd2, ihd2, 0, stAng2 + swAng, stAng2, true);
      ctx.closePath();
      break;
    }

    // ── Wedge / pie slice ─────────────────────────────────────────────────────
    case 'pie':
    case 'pieWedge': {
      const startA = -Math.PI / 2;
      const sweepA = Math.PI * 1.5; // 270° default
      ctx.moveTo(cx, cy);
      ctx.arc(cx, cy, Math.min(w, h) / 2, startA, startA + sweepA);
      ctx.closePath();
      break;
    }

    // ── Cloud ─────────────────────────────────────────────────────────────────
    case 'cloud': {
      // Simplified cloud using arcs
      const r = h * 0.28;
      ctx.arc(x + w * 0.25, y + h * 0.55, r, Math.PI, Math.PI * 1.8);
      ctx.arc(x + w * 0.45, y + h * 0.35, r * 1.1, Math.PI * 1.3, Math.PI * 1.9);
      ctx.arc(x + w * 0.65, y + h * 0.4, r, Math.PI * 1.5, Math.PI * 2);
      ctx.arc(x + w * 0.8, y + h * 0.6, r * 0.9, Math.PI * 1.6, Math.PI * 0.1);
      ctx.arc(x + w * 0.55, y + h * 0.75, r, 0, Math.PI * 0.7);
      ctx.arc(x + w * 0.25, y + h * 0.7, r * 0.9, 0, Math.PI);
      ctx.closePath();
      break;
    }

    // ── Parallelogram / funnel ────────────────────────────────────────────────
    case 'funnel': {
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(cx + w * 0.15, y + h);
      ctx.lineTo(cx - w * 0.15, y + h);
      ctx.closePath();
      break;
    }

    // ── Smiley face ───────────────────────────────────────────────────────────
    // Spec: filled circle body + two filled eye circles + smile quadratic arc
    case 'smileyface': {
      // Body circle
      ctx.ellipse(cx, cy, w / 2, h / 2, 0, 0, Math.PI * 2);
      ctx.closePath();
      // Left eye (filled sub-path, evenodd makes it a hole in fill)
      const eyeRx = w * 0.05;
      const eyeRy = h * 0.05;
      const eyeY  = cy - h * 0.12;
      ctx.moveTo(cx - w * 0.2 + eyeRx, eyeY);
      ctx.ellipse(cx - w * 0.2, eyeY, eyeRx, eyeRy, 0, 0, Math.PI * 2);
      // Right eye
      ctx.moveTo(cx + w * 0.2 + eyeRx, eyeY);
      ctx.ellipse(cx + w * 0.2, eyeY, eyeRx, eyeRy, 0, 0, Math.PI * 2);
      // Smile: open arc rendered as stroke sub-path
      ctx.moveTo(cx - w * 0.25, cy + h * 0.05);
      ctx.quadraticCurveTo(cx, cy + h * 0.3, cx + w * 0.25, cy + h * 0.05);
      break;
    }

    // ── Fold / document ───────────────────────────────────────────────────────
    case 'document':
    case 'foldedcorner': {
      const fold = Math.min(w, h) * 0.15;
      ctx.moveTo(x, y);
      ctx.lineTo(x + w - fold, y);
      ctx.lineTo(x + w, y + fold);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      ctx.moveTo(x + w - fold, y);
      ctx.lineTo(x + w - fold, y + fold);
      ctx.lineTo(x + w, y + fold);
      break;
    }

    // ── Snipped-corner rectangles ─────────────────────────────────────────────
    case 'snip1rect': {
      // One snipped top-right corner; adj = snip size (default 16667)
      const a = Math.min(50000, Math.max(0, adj ?? 16667));
      const s = Math.min(w, h) * a / 100000;
      ctx.moveTo(x, y);
      ctx.lineTo(x + w - s, y);
      ctx.lineTo(x + w, y + s);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;
    }
    case 'snip2samerect': {
      // Two snipped corners (top-right + bottom-left); adj = snip size
      const a = Math.min(50000, Math.max(0, adj ?? 16667));
      const s = Math.min(w, h) * a / 100000;
      ctx.moveTo(x, y);
      ctx.lineTo(x + w - s, y);
      ctx.lineTo(x + w, y + s);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x + s, y + h);
      ctx.lineTo(x, y + h - s);
      ctx.closePath();
      break;
    }
    case 'snip2diagrect': {
      // Two snipped diagonal corners (top-right + bottom-left)
      const a = Math.min(50000, Math.max(0, adj ?? 16667));
      const s = Math.min(w, h) * a / 100000;
      ctx.moveTo(x + s, y);
      ctx.lineTo(x + w - s, y);
      ctx.lineTo(x + w, y + s);
      ctx.lineTo(x + w, y + h - s);
      ctx.lineTo(x + w - s, y + h);
      ctx.lineTo(x + s, y + h);
      ctx.lineTo(x, y + h - s);
      ctx.lineTo(x, y + s);
      ctx.closePath();
      break;
    }
    case 'snipRoundRect':
    case 'sniproundrect': {
      // One snipped + one rounded corner
      const a = Math.min(50000, Math.max(0, adj ?? 16667));
      const s = Math.min(w, h) * a / 100000;
      ctx.moveTo(x + s, y);
      ctx.lineTo(x + w - s, y);
      ctx.lineTo(x + w, y + s);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x, y + h);
      ctx.quadraticCurveTo(x, y, x + s, y);
      ctx.closePath();
      break;
    }
    case 'round1rect': {
      // One rounded corner (top-left); adj = corner size
      const a = Math.min(50000, Math.max(0, adj ?? 16667));
      const r = Math.min(w, h) * a / 100000;
      ctx.moveTo(x + r, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x, y + h);
      ctx.lineTo(x, y + r);
      ctx.quadraticCurveTo(x, y, x + r, y);
      ctx.closePath();
      break;
    }
    case 'round2samerect': {
      // Two rounded corners on same side (top); adj = corner size
      const a = Math.min(50000, Math.max(0, adj ?? 16667));
      const r = Math.min(w, h) * a / 100000;
      ctx.moveTo(x + r, y);
      ctx.lineTo(x + w - r, y);
      ctx.quadraticCurveTo(x + w, y, x + w, y + r);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x, y + h);
      ctx.lineTo(x, y + r);
      ctx.quadraticCurveTo(x, y, x + r, y);
      ctx.closePath();
      break;
    }
    case 'round2diagrect': {
      // Two rounded diagonal corners (top-left + bottom-right)
      const a = Math.min(50000, Math.max(0, adj ?? 16667));
      const r = Math.min(w, h) * a / 100000;
      ctx.moveTo(x + r, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + h - r);
      ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
      ctx.lineTo(x, y + h);
      ctx.lineTo(x, y + r);
      ctx.quadraticCurveTo(x, y, x + r, y);
      ctx.closePath();
      break;
    }

    // ── Misc shapes ───────────────────────────────────────────────────────────
    case 'plaque': {
      // Rectangle with concave quarter-circle corners
      const r = Math.min(w, h) * 0.25;
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
      break;
    }
    case 'can': {
      // Cylinder (top + body ellipse); approximate with top arc + rect sides
      const ry = h * 0.1;
      ctx.ellipse(cx, y + ry, w / 2, ry, 0, 0, Math.PI * 2);
      ctx.moveTo(x, y + ry);
      ctx.lineTo(x, y + h - ry);
      ctx.ellipse(cx, y + h - ry, w / 2, ry, 0, 0, Math.PI);
      ctx.lineTo(x + w, y + ry);
      break;
    }
    case 'cube': {
      const off = Math.min(w, h) * 0.2;
      ctx.moveTo(x + off, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + h - off);
      ctx.lineTo(x + w - off, y + h);
      ctx.lineTo(x, y + h);
      ctx.lineTo(x, y + off);
      ctx.closePath();
      ctx.moveTo(x + off, y);
      ctx.lineTo(x + off, y + off);
      ctx.lineTo(x + w, y + off);
      ctx.moveTo(x + off, y + off);
      ctx.lineTo(x, y + off);
      break;
    }
    case 'bevel': {
      // Beveled rectangle (inset rectangle + corner lines)
      const bev = Math.min(w, h) * 0.1;
      ctx.rect(x, y, w, h);
      ctx.moveTo(x, y);
      ctx.lineTo(x + bev, y + bev);
      ctx.lineTo(x + w - bev, y + bev);
      ctx.lineTo(x + w, y);
      ctx.moveTo(x + w - bev, y + bev);
      ctx.lineTo(x + w - bev, y + h - bev);
      ctx.lineTo(x + w, y + h);
      ctx.moveTo(x + w - bev, y + h - bev);
      ctx.lineTo(x + bev, y + h - bev);
      ctx.lineTo(x, y + h);
      ctx.moveTo(x + bev, y + h - bev);
      ctx.lineTo(x + bev, y + bev);
      break;
    }
    case 'halfframe': {
      // L-shaped half-frame
      const th = Math.min(w, h) * 0.25;
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + th);
      ctx.lineTo(x + th, y + th);
      ctx.lineTo(x + th, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;
    }
    case 'corner': {
      // L-shaped corner bracket
      const th = Math.min(w, h) * 0.25;
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + th);
      ctx.lineTo(x + th, y + th);
      ctx.lineTo(x + th, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;
    }
    case 'irregularSeal1':
    case 'irregularseal1':
    case 'irregularSeal2':
    case 'irregularseal2': {
      // Starburst / explosion shape (simplified)
      const pts = geom.includes('1') ? 6 : 8;
      const outerR = Math.min(w, h) / 2;
      const innerR = outerR * 0.5;
      const angleStep = Math.PI / pts;
      let first = true;
      for (let i = 0; i < pts * 2; i++) {
        const r = i % 2 === 0 ? outerR : innerR;
        const angle = i * angleStep - Math.PI / 2;
        const px = cx + r * Math.cos(angle);
        const py = cy + r * Math.sin(angle);
        if (first) { ctx.moveTo(px, py); first = false; }
        else ctx.lineTo(px, py);
      }
      ctx.closePath();
      break;
    }
    case 'flowchartalternateprocess':
    case 'flowchartprocess': {
      const a2 = Math.min(50000, Math.max(0, adj ?? 16667));
      const r2 = Math.min(w, h) * a2 / 100000;
      ctx.roundRect(x, y, w, h, [{ x: r2, y: r2 }]);
      break;
    }
    case 'flowchartdecision': {
      // Diamond
      ctx.moveTo(cx, y);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(cx, y + h);
      ctx.lineTo(x, cy);
      ctx.closePath();
      break;
    }
    case 'flowchartterminator': {
      // Stadium / pill shape
      const sr = Math.min(w, h) / 2;
      ctx.roundRect(x, y, w, h, [{ x: sr, y: sr }]);
      break;
    }
    case 'flowchartdocument': {
      // Rectangle with wavy bottom
      const waveH = h * 0.1;
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + h - waveH);
      ctx.bezierCurveTo(x + w * 0.75, y + h, x + w * 0.25, y + h - waveH * 2, x, y + h - waveH);
      ctx.closePath();
      break;
    }
    case 'flowchartpredefinedprocess': {
      const barW = w * 0.1;
      ctx.rect(x, y, w, h);
      ctx.moveTo(x + barW, y);
      ctx.lineTo(x + barW, y + h);
      ctx.moveTo(x + w - barW, y);
      ctx.lineTo(x + w - barW, y + h);
      break;
    }
    case 'flowchartsort': {
      // Diamond
      ctx.moveTo(cx, y);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(cx, y + h);
      ctx.lineTo(x, cy);
      ctx.closePath();
      ctx.moveTo(x, cy);
      ctx.lineTo(x + w, cy);
      break;
    }
    case 'flowchartmanualinput': {
      const sl = h * 0.2;
      ctx.moveTo(x, y + sl);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;
    }
    case 'moon': {
      // Crescent moon
      ctx.arc(cx, cy, Math.min(w, h) / 2, -Math.PI / 2, Math.PI / 2);
      ctx.arc(cx - w * 0.2, cy, Math.min(w, h) / 2, Math.PI / 2, -Math.PI / 2, true);
      ctx.closePath();
      break;
    }
    case 'arc': {
      // Open arc
      ctx.arc(cx, cy, Math.min(w, h) / 2, -Math.PI / 2, Math.PI);
      break;
    }

    // ── Math operator shapes ───────────────────────────────────────────────────
    case 'mathequal': {
      const barH = Math.max(1, (adj ?? 23520) / 100000 * h);
      const gap  = 17490 / 100000 * h;
      ctx.rect(x, cy - gap / 2 - barH, w, barH);
      ctx.rect(x, cy + gap / 2,        w, barH);
      break;
    }

    case 'mathmultiply': {
      // "×": a "+" shape rotated 45°
      const t  = (adj ?? 23520) / 100000 * Math.min(w, h) * 0.5;
      const hl = Math.max(w, h) * 0.72;
      ctx.save();
      ctx.translate(cx, cy);
      ctx.rotate(Math.PI / 4);
      ctx.rect(-t, -hl, 2 * t, 2 * hl);
      ctx.rect(-hl, -t, 2 * hl, 2 * t);
      ctx.restore();
      break;
    }

    case 'mathplus': {
      const t = (adj ?? 23520) / 100000 * Math.min(w, h) * 0.5;
      ctx.rect(cx - t, y, 2 * t, h);
      ctx.rect(x, cy - t, w, 2 * t);
      break;
    }

    case 'mathminus': {
      const barH = Math.max(1, (adj ?? 23520) / 100000 * h);
      ctx.rect(x, cy - barH / 2, w, barH);
      break;
    }

    case 'mathdivide': {
      const barH  = Math.max(1, (adj ?? 23520) / 100000 * h);
      const dotR  = barH * 1.1;
      const dotGap = h * 0.22;
      ctx.rect(x, cy - barH / 2, w, barH);
      ctx.arc(cx, cy - dotGap, dotR, 0, Math.PI * 2);
      ctx.arc(cx, cy + dotGap, dotR, 0, Math.PI * 2);
      break;
    }

    // ── 4-direction arrow ────────────────────────────────────────────────────
    case 'quadarrow': {
      const sw  = w * (adj  ?? 23000) / 100000;
      const ahw = w * (adj2 ?? 30000) / 100000;
      const sx  = x + (w - sw) / 2;
      const sy2 = y + (h - sw) / 2;
      ctx.moveTo(cx, y);
      ctx.lineTo(x + w - ahw, y + ahw);
      ctx.lineTo(x + w - ahw, sy2);
      ctx.lineTo(sx + sw, sy2);
      ctx.lineTo(sx + sw, y + ahw);
      ctx.lineTo(x + ahw, y + ahw);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(x + w - ahw, y + h - ahw);
      ctx.lineTo(sx + sw, y + h - ahw);
      ctx.lineTo(sx + sw, sy2 + sw);
      ctx.lineTo(x + w - ahw, sy2 + sw);
      ctx.lineTo(x + w - ahw, y + h - ahw);
      ctx.lineTo(cx, y + h);
      ctx.lineTo(x + ahw, y + h - ahw);
      ctx.lineTo(x + ahw, sy2 + sw);
      ctx.lineTo(sx, sy2 + sw);
      ctx.lineTo(sx, y + h - ahw);
      ctx.lineTo(x, cy);
      ctx.lineTo(x + ahw, y + ahw);
      ctx.lineTo(sx, y + ahw);
      ctx.lineTo(sx, sy2);
      ctx.lineTo(x + ahw, sy2);
      ctx.closePath();
      break;
    }

    // ── Quad-arrow callout ────────────────────────────────────────────────────
    case 'quadarrowcallout': {
      const t = Math.min(w, h) * 0.25;
      ctx.rect(x + t, y + t, w - 2 * t, h - 2 * t);
      ctx.moveTo(cx, y); ctx.lineTo(x + t, y + t); ctx.lineTo(x + w - t, y + t); ctx.closePath();
      ctx.moveTo(cx, y + h); ctx.lineTo(x + t, y + h - t); ctx.lineTo(x + w - t, y + h - t); ctx.closePath();
      ctx.moveTo(x, cy); ctx.lineTo(x + t, y + t); ctx.lineTo(x + t, y + h - t); ctx.closePath();
      ctx.moveTo(x + w, cy); ctx.lineTo(x + w - t, y + t); ctx.lineTo(x + w - t, y + h - t); ctx.closePath();
      break;
    }

    // ── Wave ──────────────────────────────────────────────────────────────────
    case 'wave': {
      const wAmp = h * (adj ?? 12500) / 100000;
      ctx.moveTo(x, cy - wAmp);
      ctx.bezierCurveTo(x + w * 0.25, cy - wAmp * 2, x + w * 0.25, cy, x + w * 0.5, cy);
      ctx.bezierCurveTo(x + w * 0.75, cy, x + w * 0.75, cy - wAmp * 2, x + w, cy - wAmp);
      ctx.lineTo(x + w, cy + wAmp);
      ctx.bezierCurveTo(x + w * 0.75, cy + wAmp * 2, x + w * 0.75, cy, x + w * 0.5, cy);
      ctx.bezierCurveTo(x + w * 0.25, cy, x + w * 0.25, cy + wAmp * 2, x, cy + wAmp);
      ctx.closePath();
      break;
    }

    // ── Sun ───────────────────────────────────────────────────────────────────
    case 'sun': {
      const outerR = Math.min(w, h) / 2;
      const innerR = outerR * 0.55;
      const rayLen = outerR * 0.35;
      const rayW   = outerR * 0.1;
      ctx.arc(cx, cy, innerR, 0, Math.PI * 2);
      for (let i = 0; i < 8; i++) {
        const angle = (i / 8) * Math.PI * 2;
        ctx.save();
        ctx.translate(cx, cy);
        ctx.rotate(angle);
        ctx.rect(innerR + 2, -rayW / 2, rayLen, rayW);
        ctx.restore();
      }
      break;
    }

    // ── Lightning bolt ────────────────────────────────────────────────────────
    case 'lightningbolt': {
      ctx.moveTo(cx + w * 0.1, y);
      ctx.lineTo(x, cy - h * 0.05);
      ctx.lineTo(cx + w * 0.05, cy - h * 0.05);
      ctx.lineTo(cx - w * 0.1, y + h);
      ctx.lineTo(x + w, cy + h * 0.05);
      ctx.lineTo(cx - w * 0.05, cy + h * 0.05);
      ctx.closePath();
      break;
    }

    // ── Frame (hollow rectangle) ──────────────────────────────────────────────
    case 'frame': {
      const th = Math.min(w, h) * (adj ?? 12500) / 100000;
      ctx.rect(x, y, w, h);
      ctx.rect(x + th, y + th, w - 2 * th, h - 2 * th);
      break;
    }

    // ── Bracket pair [] ───────────────────────────────────────────────────────
    case 'bracketpair': {
      const a   = Math.min(50000, Math.max(0, adj ?? 8333));
      const arcH = h * a / 100000;
      ctx.moveTo(x + w * 0.4, y);
      ctx.quadraticCurveTo(x, y, x, y + arcH);
      if (h - 2 * arcH > 0) ctx.lineTo(x, y + h - arcH);
      ctx.quadraticCurveTo(x, y + h, x + w * 0.4, y + h);
      ctx.moveTo(x + w * 0.6, y);
      ctx.quadraticCurveTo(x + w, y, x + w, y + arcH);
      if (h - 2 * arcH > 0) ctx.lineTo(x + w, y + h - arcH);
      ctx.quadraticCurveTo(x + w, y + h, x + w * 0.6, y + h);
      break;
    }

    // ── Brace pair {} ─────────────────────────────────────────────────────────
    case 'bracepair': {
      const nb = w * 0.2;
      ctx.moveTo(x + w * 0.4, y);
      ctx.bezierCurveTo(x + w * 0.4 - nb, y, x + w * 0.4 - nb, cy - h * 0.08, x, cy);
      ctx.bezierCurveTo(x + w * 0.4 - nb, cy + h * 0.08, x + w * 0.4 - nb, y + h, x + w * 0.4, y + h);
      ctx.moveTo(x + w * 0.6, y);
      ctx.bezierCurveTo(x + w * 0.6 + nb, y, x + w * 0.6 + nb, cy - h * 0.08, x + w, cy);
      ctx.bezierCurveTo(x + w * 0.6 + nb, cy + h * 0.08, x + w * 0.6 + nb, y + h, x + w * 0.6, y + h);
      break;
    }

    // ── Chord (arc + closing line) ────────────────────────────────────────────
    case 'chord': {
      const startA = -Math.PI / 2 + (adj  ?? 2700000)  / 21600000 * Math.PI * 2;
      const endA   = -Math.PI / 2 + (adj2 ?? 16200000) / 21600000 * Math.PI * 2;
      ctx.arc(cx, cy, Math.min(w, h) / 2, startA, endA);
      ctx.closePath();
      break;
    }

    // ── Block arc ─────────────────────────────────────────────────────────────
    case 'blockarc': {
      const outerR   = Math.min(w, h) / 2;
      const innerFrac = (adj2 ?? 25000) / 100000;
      const innerR   = outerR * (1 - innerFrac);
      const startA   = -Math.PI / 2 + (adj ?? 10800000) / 21600000 * Math.PI * 2;
      const endA     = Math.PI / 2;
      ctx.arc(cx, cy, outerR, startA, endA);
      ctx.arc(cx, cy, innerR, endA, startA, true);
      ctx.closePath();
      break;
    }

    // ── Teardrop ──────────────────────────────────────────────────────────────
    case 'teardrop': {
      const r   = Math.min(w, h) * 0.4;
      const bCx = x + r;
      const bCy = y + h - r;
      ctx.arc(bCx, bCy, r, 0, Math.PI * 2 * 0.75);
      ctx.bezierCurveTo(bCx - r * 0.1, bCy - r, x + w - r, y + r, x + w, y);
      ctx.bezierCurveTo(x + w - r * 0.2, y + r * 0.5, bCx + r, bCy - r * 1.1, bCx + r, bCy);
      ctx.closePath();
      break;
    }

    // ── Diagonal stripe ───────────────────────────────────────────────────────
    case 'diagstripe': {
      const t = Math.min(w, h) * (adj ?? 50000) / 100000;
      ctx.moveTo(x, y + h - t);
      ctx.lineTo(x + t, y + h);
      ctx.lineTo(x + w, y + t);
      ctx.lineTo(x + w - t, y);
      ctx.closePath();
      break;
    }

    // ── Wedge round-rect callout ──────────────────────────────────────────────
    case 'wedgeroundrectcallout': {
      const r2 = Math.min(w, h) * 0.1;
      ctx.roundRect(x, y, w, h * 0.85, r2);
      ctx.moveTo(x + w * 0.1, y + h * 0.85);
      ctx.lineTo(x + w * 0.2, y + h);
      ctx.lineTo(x + w * 0.3, y + h * 0.85);
      ctx.closePath();
      break;
    }

    // ── Arrow callouts ────────────────────────────────────────────────────────
    case 'rightarrowcallout': {
      const shH = h * (adj  ?? 50000) / 100000;
      const shW = w * (adj2 ?? 50000) / 100000;
      const sy  = y + (h - shH) / 2;
      ctx.rect(x, sy, shW, shH);
      ctx.moveTo(x + shW, y); ctx.lineTo(x + w, cy); ctx.lineTo(x + shW, y + h); ctx.closePath();
      break;
    }
    case 'leftarrowcallout': {
      const shH = h * (adj  ?? 50000) / 100000;
      const shW = w * (adj2 ?? 50000) / 100000;
      const sy  = y + (h - shH) / 2;
      ctx.rect(x + w - shW, sy, shW, shH);
      ctx.moveTo(x + w - shW, y); ctx.lineTo(x, cy); ctx.lineTo(x + w - shW, y + h); ctx.closePath();
      break;
    }
    case 'uparrowcallout': {
      const shW = w * (adj  ?? 50000) / 100000;
      const shH = h * (adj2 ?? 50000) / 100000;
      const sx  = x + (w - shW) / 2;
      ctx.rect(sx, y + shH, shW, h - shH);
      ctx.moveTo(x, y + shH); ctx.lineTo(cx, y); ctx.lineTo(x + w, y + shH); ctx.closePath();
      break;
    }
    case 'downarrowcallout': {
      const shW = w * (adj  ?? 50000) / 100000;
      const shH = h * (adj2 ?? 50000) / 100000;
      const sx  = x + (w - shW) / 2;
      ctx.rect(sx, y, shW, h - shH);
      ctx.moveTo(x, y + h - shH); ctx.lineTo(cx, y + h); ctx.lineTo(x + w, y + h - shH); ctx.closePath();
      break;
    }
    case 'leftrightarrowcallout': {
      const shH = h * (adj  ?? 50000) / 100000;
      const shW = w * (adj2 ?? 25000) / 100000;
      const sy  = y + (h - shH) / 2;
      ctx.rect(x + shW, sy, w - 2 * shW, shH);
      ctx.moveTo(x + shW, y); ctx.lineTo(x, cy); ctx.lineTo(x + shW, y + h); ctx.closePath();
      ctx.moveTo(x + w - shW, y); ctx.lineTo(x + w, cy); ctx.lineTo(x + w - shW, y + h); ctx.closePath();
      break;
    }

    // ── Left-right-up arrow ───────────────────────────────────────────────────
    case 'leftrightuparrow': {
      const sw  = w * (adj  ?? 25000) / 100000;
      const ahh = h * (adj2 ?? 30000) / 100000;
      const sx  = x + (w - sw) / 2;
      ctx.moveTo(cx, y);
      ctx.lineTo(x + w, y + ahh);
      ctx.lineTo(sx + sw, y + ahh);
      ctx.lineTo(sx + sw, y + h);
      ctx.lineTo(sx, y + h);
      ctx.lineTo(sx, y + ahh);
      ctx.lineTo(x, y + ahh);
      ctx.closePath();
      break;
    }

    // ── Left-up arrow ─────────────────────────────────────────────────────────
    case 'leftuparrow': {
      const ahLen = Math.min(w, h) * (adj ?? 30000) / 100000;
      const sw    = Math.min(w, h) * (adj2 ?? 25000) / 100000;
      ctx.moveTo(x, cy);
      ctx.lineTo(x + ahLen, y);
      ctx.lineTo(x + ahLen, y + h - ahLen - sw);
      ctx.lineTo(cx + sw / 2, y + h - ahLen - sw);
      ctx.lineTo(cx + sw / 2, y + h);
      ctx.lineTo(x + w, y + h - ahLen);
      ctx.lineTo(cx + sw / 2 + sw, y + h - ahLen);
      ctx.lineTo(cx + sw / 2 + sw, y + ahLen);
      ctx.lineTo(x + ahLen, y + ahLen);
      ctx.lineTo(x + ahLen, cy + sw / 2);
      ctx.closePath();
      break;
    }

    // ── U-turn arrow ──────────────────────────────────────────────────────────
    // Spec (ECMA-376): outer half-arc on top, arrowhead on right side pointing down
    case 'uturnarrow': {
      const sw     = w * (adj ?? 25000) / 100000;   // shaft width
      const outerR = (w - sw) / 2;                   // outer bend radius
      const innerR = Math.max(0, outerR - sw);        // inner bend radius
      const arcCX  = x + sw + outerR;                // arc center X
      const arcCY  = y + sw + outerR;                // arc center Y
      const ahW    = sw * 2;                          // arrowhead full width
      const ahBase = y + h - sw * 2.5;               // where arrowhead base starts
      // shaft: left side down, U-arc across top, right side down to arrowhead
      ctx.moveTo(x, y + h);
      ctx.lineTo(x, arcCY);
      ctx.arc(arcCX, arcCY, outerR, Math.PI, 0);
      ctx.lineTo(x + w, ahBase);
      // arrowhead (pointing downward on right side)
      ctx.lineTo(x + w + (ahW - sw) / 2, ahBase);
      ctx.lineTo(arcCX + sw / 2, y + h);  // tip
      ctx.lineTo(x + w - (ahW - sw) / 2 - sw, ahBase);
      ctx.lineTo(x + w - sw, ahBase);
      ctx.lineTo(x + w - sw, arcCY);
      ctx.arc(arcCX, arcCY, innerR, 0, Math.PI, true);
      ctx.lineTo(x + sw, y + h);
      ctx.closePath();
      break;
    }

    // ── Bent arrow / bent-up arrow ────────────────────────────────────────────
    case 'bentarrow':
    case 'bentuparrow': {
      const t = Math.min(w, h) * 0.25;
      ctx.moveTo(x, cy - t / 2);
      ctx.lineTo(x + w - t * 2, cy - t / 2);
      ctx.lineTo(x + w - t * 2, y + t);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(x + w - t * 2, y + h - t);
      ctx.lineTo(x + w - t * 2, cy + t / 2);
      ctx.lineTo(x, cy + t / 2);
      ctx.closePath();
      break;
    }

    // ── Plus shape (non-math) ─────────────────────────────────────────────────
    case 'plus': {
      const t = Math.min(w, h) * (adj ?? 25000) / 100000;
      ctx.rect(cx - t, y, 2 * t, h);
      ctx.rect(x, cy - t, w, 2 * t);
      break;
    }

    // ── Math not-equal ────────────────────────────────────────────────────────
    case 'mathnotequal': {
      const barH = Math.max(1, (adj ?? 23520) / 100000 * h);
      const gap  = 17490 / 100000 * h;
      ctx.rect(x, cy - gap / 2 - barH, w, barH);
      ctx.rect(x, cy + gap / 2,        w, barH);
      ctx.moveTo(cx - w * 0.15, y + h * 0.1);
      ctx.lineTo(cx + w * 0.15, y + h * 0.9);
      break;
    }

    // ── Flowchart: connector (circle) ─────────────────────────────────────────
    case 'flowchartconnector': {
      ctx.ellipse(cx, cy, w / 2, h / 2, 0, 0, Math.PI * 2);
      break;
    }

    // ── Flowchart: delay (D-shape) ────────────────────────────────────────────
    case 'flowchartdelay': {
      const r = h / 2;
      ctx.moveTo(x, y);
      ctx.lineTo(x + w - r, y);
      ctx.arc(x + w - r, cy, r, -Math.PI / 2, Math.PI / 2);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;
    }

    // ── Flowchart: display (pentagon-like) ────────────────────────────────────
    case 'flowchartdisplay': {
      const lx = w * 0.2;
      const rx = w * 0.15;
      ctx.moveTo(x + lx, y);
      ctx.lineTo(x + w - rx, y);
      ctx.arc(x + w - rx, cy, h / 2, -Math.PI / 2, Math.PI / 2);
      ctx.lineTo(x + lx, y + h);
      ctx.lineTo(x, cy);
      ctx.closePath();
      break;
    }

    // ── Flowchart: input/output (parallelogram) ───────────────────────────────
    case 'flowchartinputoutput':
    case 'flowchartpunchedcard': {
      const sl = w * 0.2;
      ctx.moveTo(x + sl, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w - sl, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;
    }

    // ── Flowchart: merge (inverted triangle) ──────────────────────────────────
    case 'flowchartmerge': {
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(cx, y + h);
      ctx.closePath();
      break;
    }

    // ── Flowchart: extract (upward triangle) ─────────────────────────────────
    case 'flowchartextract': {
      ctx.moveTo(cx, y);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;
    }

    // ── Flowchart: off-page connector (pentagon pointing down) ────────────────
    case 'flowchartoffpageconnector': {
      const tipH = h * 0.3;
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + h - tipH);
      ctx.lineTo(cx, y + h);
      ctx.lineTo(x, y + h - tipH);
      ctx.closePath();
      break;
    }

    // ── Flowchart: online storage / manual label (rect fallback) ─────────────
    case 'flowchartonlinestorage':
    case 'flowchartmanuallabel':
    case 'flowchartpuncheddisk': {
      ctx.rect(x, y, w, h);
      break;
    }

    // ── Horizontal scroll ─────────────────────────────────────────────────────
    case 'horizontalscroll': {
      const r = Math.min(w, h) * 0.15;
      ctx.roundRect(x + r, y, w - r, h, r);
      ctx.moveTo(x + r, y + r * 2);
      ctx.arc(x + r, y + r, r, Math.PI / 2, Math.PI * 2.5);
      break;
    }

    // ── Vertical scroll ───────────────────────────────────────────────────────
    case 'verticalscroll': {
      const r = Math.min(w, h) * 0.15;
      ctx.roundRect(x, y + r, w, h - r, r);
      ctx.moveTo(x + r * 2, y + r);
      ctx.arc(x + r, y + r, r, 0, Math.PI * 2);
      break;
    }

    // ── Ribbon (fold at top) ──────────────────────────────────────────────────
    case 'ribbon': {
      const foldH = h * 0.25;
      const notch = w * 0.08;
      ctx.moveTo(x, y + foldH);
      ctx.lineTo(x + notch, y);
      ctx.lineTo(cx, y + foldH * 0.5);
      ctx.lineTo(x + w - notch, y);
      ctx.lineTo(x + w, y + foldH);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;
    }

    // ── Ribbon2 (fold at bottom) ──────────────────────────────────────────────
    case 'ribbon2': {
      const foldH = h * 0.25;
      const notch = w * 0.08;
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + h - foldH);
      ctx.lineTo(x + w - notch, y + h);
      ctx.lineTo(cx, y + h - foldH * 0.5);
      ctx.lineTo(x + notch, y + h);
      ctx.lineTo(x, y + h - foldH);
      ctx.closePath();
      break;
    }

    // ── Ellipse ribbon (curved bottom) ───────────────────────────────────────
    case 'ellipseribbon': {
      const adj1 = (adj ?? 25000) / 100000;
      const foldH = h * adj1;
      const notch = w * 0.08;
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + h - foldH);
      // curved bottom via ellipse arc
      ctx.ellipse(cx, y + h - foldH, w / 2, foldH, 0, 0, Math.PI);
      ctx.lineTo(x, y + h - foldH);
      ctx.closePath();
      // fold triangles
      ctx.moveTo(x, y + foldH * 0.5);
      ctx.lineTo(x + notch, y);
      ctx.lineTo(x + notch * 2, y + foldH * 0.5);
      ctx.moveTo(x + w - notch * 2, y + foldH * 0.5);
      ctx.lineTo(x + w - notch, y);
      ctx.lineTo(x + w, y + foldH * 0.5);
      break;
    }

    // ── Ellipse ribbon 2 (curved top) ────────────────────────────────────────
    case 'ellipseribbon2': {
      const adj1 = (adj  ?? 25000) / 100000;
      const notch = w * 0.08;
      const foldH = h * adj1;
      ctx.moveTo(x, y + foldH);
      // curved top via ellipse arc
      ctx.ellipse(cx, y + foldH, w / 2, foldH, 0, Math.PI, 0);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      // fold triangles
      ctx.moveTo(x, y + foldH * 1.5);
      ctx.lineTo(x + notch, y + h * 1.0);
      ctx.lineTo(x + notch * 2, y + foldH * 1.5);
      ctx.moveTo(x + w - notch * 2, y + foldH * 1.5);
      ctx.lineTo(x + w - notch, y + h);
      ctx.lineTo(x + w, y + foldH * 1.5);
      break;
    }

    // ── Circular arrow ────────────────────────────────────────────────────────
    case 'circulararrow': {
      const outerR = Math.min(w, h) / 2;
      const innerR = outerR * 0.5;
      ctx.arc(cx, cy, outerR, -Math.PI * 0.8, Math.PI * 0.4);
      ctx.arc(cx, cy, innerR, Math.PI * 0.4, -Math.PI * 0.8, true);
      ctx.closePath();
      break;
    }

    // ── Curved directional arrows (simplified) ────────────────────────────────
    case 'curveduparrow': {
      ctx.moveTo(cx, y);
      ctx.lineTo(x + w, y + h * 0.45);
      ctx.lineTo(x + w * 0.65, y + h * 0.45);
      ctx.quadraticCurveTo(x + w * 0.65, y + h, cx, y + h);
      ctx.quadraticCurveTo(x + w * 0.35, y + h, x + w * 0.35, y + h * 0.45);
      ctx.lineTo(x, y + h * 0.45);
      ctx.closePath();
      break;
    }
    case 'curveddownarrow': {
      ctx.moveTo(cx, y + h);
      ctx.lineTo(x + w, y + h * 0.55);
      ctx.lineTo(x + w * 0.65, y + h * 0.55);
      ctx.quadraticCurveTo(x + w * 0.65, y, cx, y);
      ctx.quadraticCurveTo(x + w * 0.35, y, x + w * 0.35, y + h * 0.55);
      ctx.lineTo(x, y + h * 0.55);
      ctx.closePath();
      break;
    }
    case 'curvedleftarrow': {
      ctx.moveTo(x, cy);
      ctx.lineTo(x + w * 0.45, y);
      ctx.lineTo(x + w * 0.45, y + h * 0.35);
      ctx.quadraticCurveTo(x + w, y + h * 0.35, x + w, cy);
      ctx.quadraticCurveTo(x + w, y + h * 0.65, x + w * 0.45, y + h * 0.65);
      ctx.lineTo(x + w * 0.45, y + h);
      ctx.closePath();
      break;
    }
    case 'curvedrightarrow': {
      ctx.moveTo(x + w, cy);
      ctx.lineTo(x + w * 0.55, y);
      ctx.lineTo(x + w * 0.55, y + h * 0.35);
      ctx.quadraticCurveTo(x, y + h * 0.35, x, cy);
      ctx.quadraticCurveTo(x, y + h * 0.65, x + w * 0.55, y + h * 0.65);
      ctx.lineTo(x + w * 0.55, y + h);
      ctx.closePath();
      break;
    }

    // ── Striped right arrow (3 stripes + arrowhead) ───────────────────────────
    // Spec: ssd = min(w,h), ssd32=ssd/32, ssd8=ssd/8 etc. adj=arrowhead length
    case 'stripedrightarrow': {
      const ssd   = Math.min(w, h);
      const ssd32 = ssd / 32;
      const ssd16 = ssd / 16;
      const ssd8  = ssd / 8;
      const shH   = ssd * (adj ?? 50000) / 100000;  // shaft height
      const ahW   = w * (adj2 ?? 50000) / 100000;   // arrowhead width
      const y1    = cy - shH / 2;
      const y2    = cy + shH / 2;
      const x4    = x + w - ahW;
      // stripe 1
      ctx.rect(x, y1, ssd32, shH);
      // stripe 2
      ctx.rect(x + ssd16, y1, ssd16, shH);
      // stripe 3 (narrower, bridging to arrowhead)
      ctx.rect(x + ssd8, y1, ssd8, shH);
      // arrow body + head
      ctx.moveTo(x4, y1);
      ctx.lineTo(x4, y);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(x4, y + h);
      ctx.lineTo(x4, y2);
      ctx.lineTo(x + ssd8 * 2, y2);
      ctx.lineTo(x + ssd8 * 2, y1);
      ctx.closePath();
      break;
    }

    // ── Flowchart: preparation (hexagon with angled sides) ────────────────────
    case 'flowchartpreparation': {
      const sl = w * 0.2;
      ctx.moveTo(x + sl, y);
      ctx.lineTo(x + w - sl, y);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(x + w - sl, y + h);
      ctx.lineTo(x + sl, y + h);
      ctx.lineTo(x, cy);
      ctx.closePath();
      break;
    }

    // ── Flowchart: collate (hourglass) ────────────────────────────────────────
    case 'flowchartcollate': {
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x, y + h);
      ctx.lineTo(x + w, y + h);
      ctx.closePath();
      break;
    }

    // ── Flowchart: magnetic disk (vertical cylinder) ──────────────────────────
    case 'flowchartmagneticdisk': {
      const ry = h * 0.15;
      ctx.moveTo(x, y + ry);
      ctx.ellipse(cx, y + ry, w / 2, ry, 0, Math.PI, 0);
      ctx.lineTo(x + w, y + h - ry);
      ctx.ellipse(cx, y + h - ry, w / 2, ry, 0, 0, Math.PI);
      ctx.lineTo(x, y + ry);
      ctx.closePath();
      // top cap stroke line
      ctx.moveTo(x + w, y + ry);
      ctx.ellipse(cx, y + ry, w / 2, ry, 0, 0, Math.PI);
      break;
    }

    // ── Flowchart: internal storage (rect with two inner lines) ───────────────
    case 'flowchartinternalstorage': {
      ctx.rect(x, y, w, h);
      const bw = w * 0.15;
      const bh = h * 0.15;
      ctx.moveTo(x + bw, y);
      ctx.lineTo(x + bw, y + h);
      ctx.moveTo(x, y + bh);
      ctx.lineTo(x + w, y + bh);
      break;
    }

    // ── Flowchart: magnetic drum (cylinder on its side with left cap) ─────────
    case 'flowchartmagneticdrum': {
      const rx = w * 0.15;
      ctx.moveTo(x + rx, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x + rx, y + h);
      ctx.ellipse(x + rx, cy, rx, h / 2, 0, Math.PI / 2, -Math.PI / 2, true);
      ctx.closePath();
      // right cap open arc
      ctx.moveTo(x + w, y);
      ctx.ellipse(x + w, cy, rx, h / 2, 0, -Math.PI / 2, Math.PI / 2);
      break;
    }

    // ── Flowchart: summing junction (circle + X) ──────────────────────────────
    case 'flowchartsumingjunction': {
      ctx.ellipse(cx, cy, w / 2, h / 2, 0, 0, Math.PI * 2);
      const r = Math.min(w, h) / 2 * 0.65;
      ctx.moveTo(cx - r, cy - r);
      ctx.lineTo(cx + r, cy + r);
      ctx.moveTo(cx + r, cy - r);
      ctx.lineTo(cx - r, cy + r);
      break;
    }

    // ── Flowchart: magnetic tape (circle with tail) ───────────────────────────
    case 'flowchartmagnetictape': {
      // circle from bottom going around, with a small tail at bottom-right
      const r = Math.min(w, h) / 2;
      const tailX = cx + r * 0.5;
      ctx.moveTo(cx, y + h);
      ctx.arc(cx, cy, r, Math.PI / 2, Math.PI / 2 + Math.PI * 2 * 0.875);
      ctx.lineTo(tailX, cy + r * 0.5);
      ctx.lineTo(tailX, y + h);
      ctx.closePath();
      break;
    }

    // ── Flowchart: punched tape (wave bottom) ─────────────────────────────────
    case 'flowchartpunchedtape': {
      const waveH = h * 0.12;
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + h - waveH);
      ctx.bezierCurveTo(x + w * 0.75, y + h, x + w * 0.25, y + h - waveH * 2, x, y + h - waveH);
      ctx.closePath();
      // second wave on top for symmetry
      ctx.moveTo(x, y + waveH);
      ctx.bezierCurveTo(x + w * 0.25, y, x + w * 0.75, y + waveH * 2, x + w, y + waveH);
      break;
    }

    // ── Flowchart: manual operation (inverted trapezoid) ─────────────────────
    case 'flowchartmanualoperation': {
      const sl = w * 0.15;
      ctx.moveTo(x + sl, y);
      ctx.lineTo(x + w - sl, y);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;
    }

    // ── Flowchart: multidocument (stacked wave documents) ────────────────────
    case 'flowchartmultidocument': {
      const waveH = h * 0.1;
      const shiftX = w * 0.04;
      // back documents (offset rects)
      ctx.rect(x + shiftX * 2, y - h * 0.08, w - shiftX * 2, h * 0.1);
      ctx.rect(x + shiftX, y - h * 0.04, w - shiftX, h * 0.06);
      // main document with wave bottom
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + h - waveH);
      ctx.bezierCurveTo(x + w * 0.75, y + h, x + w * 0.25, y + h - waveH * 2, x, y + h - waveH);
      ctx.closePath();
      break;
    }

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
  shapeDefaultTextColor: string | null = null,
  shapeRotation = 0,
  shapeFlipH = false,
  shapeFlipV = false,
  themeDefaultColor = '#000000'
) {
  // Vertical text: rotate rendering context so text flows top-to-bottom.
  // "vert" and "eaVert" both approximate to 90° clockwise rotation.
  // "vert270" rotates 270° (= 90° counterclockwise).
  const isVert    = body.vert === 'vert' || body.vert === 'eaVert';
  const isVert270 = body.vert === 'vert270';

  if (isVert || isVert270) {
    // Set up a rotated coordinate space:
    // Centre of the bounding box remains fixed; swap w and h for the text layout.
    const cx = bx + bw / 2;
    const cy = by + bh / 2;
    ctx.save();
    ctx.translate(cx, cy);
    ctx.rotate(isVert270 ? -Math.PI / 2 : Math.PI / 2);
    // After rotation the "width" direction of the new frame is the original height
    renderTextBody(ctx, { ...body, vert: 'horz' }, -bh / 2, -bw / 2, bh, bw, scale, shapeDefaultTextColor, 0, false, false, themeDefaultColor);
    ctx.restore();
    return;
  }
  const lPad = emuToPx(body.lIns, scale);
  const rPad = emuToPx(body.rIns, scale);
  const tPad = emuToPx(body.tIns, scale);
  const bPad = emuToPx(body.bIns, scale);
  const doWrap = body.wrap !== 'none';

  const bodyDefaultBold   = body.defaultBold   ?? false;
  const bodyDefaultItalic = body.defaultItalic ?? false;
  const bodyDefaultColor = shapeDefaultTextColor ?? themeDefaultColor;

  // ── Pass 1: lay out all paragraphs ──────────────────────────────────────

  interface LineEntry {
    line: LayoutLine;
    linePx: number;       // spacing advancement (lineHeight + spaceAfter for last line)
    lineHeight: number;   // pure line height used for baseline positioning (without spaceAfter)
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

  // buildLayout runs Pass 1 at a given font scale (1.0 = normal; <1 = normAutoFit shrink)
  const buildLayout = (fontScale: number): { allLines: LineEntry[], totalHeight: number } => {
  const bodyDefaultFontSizePx = (body.defaultFontSize ?? 18) * PT_TO_EMU * scale * fontScale;
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
      ? para.defFontSize * PT_TO_EMU * scale * fontScale : bodyDefaultFontSizePx;
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
      bulletLabel = applySymbolFont(b.char, b.fontFamily ?? '');
      bulletFont  = buildFont(false, false, bSizePx, normalizeFontFamily(b.fontFamily ?? 'sans-serif'));
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
    const lines = layoutParagraph(ctx, para, maxW, paraDefaultFontSizePx, paraDefaultColor, scale, marLPx, bodyDefaultBold, bodyDefaultItalic, fontScale);

    // spaceBefore/After are in hundredths of a point → convert to canvas px
    const spaceBeforePx = para.spaceBefore != null ? (para.spaceBefore / 100) * PT_TO_EMU * scale * fontScale : 0;
    const spaceAfterPx  = para.spaceAfter  != null ? (para.spaceAfter  / 100) * PT_TO_EMU * scale * fontScale : 0;

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
        line, linePx, lineHeight, topGapPx: topGap,
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

  return { allLines, totalHeight };
  }; // end buildLayout

  let { allLines, totalHeight } = buildLayout(1.0);

  // ── normAutoFit: shrink font until text fits ─────────────────────────────
  if (body.autoFit === 'norm') {
    const maxContentH = bh - tPad - bPad;
    if (totalHeight > maxContentH && maxContentH > 0) {
      let lo = 0.1, hi = 1.0;
      for (let i = 0; i < 6; i++) {
        const mid = (lo + hi) / 2;
        if (buildLayout(mid).totalHeight <= maxContentH) lo = mid; else hi = mid;
      }
      ({ allLines, totalHeight } = buildLayout(lo));
    }
  }

  // ── anchor="b" with bh=0: auto-height growing upward from by ────────────
  // When cy=0 and anchor="b", off_y is the bottom anchor; shape grows upward.
  const anchor = body.verticalAnchor ?? 't';
  let effectiveBy = by;
  let effectiveBh: number;
  if (bh === 0 && anchor === 'b') {
    effectiveBh = tPad + totalHeight + bPad;
    effectiveBy = by - effectiveBh;
  } else {
    // ── Effective height (spAutoFit: shape expands to fit text) ─────────────
    const isSpAutoFit = body.autoFit === 'sp';
    effectiveBh = isSpAutoFit
      ? Math.max(bh, tPad + totalHeight + bPad)
      : bh;
  }

  // ── Vertical anchor ─────────────────────────────────────────────────────
  let cursorY: number;
  if (anchor === 'ctr') {
    cursorY = effectiveBy + (effectiveBh - totalHeight) / 2;
  } else if (anchor === 'b') {
    cursorY = effectiveBy + effectiveBh - totalHeight - bPad;
  } else {
    cursorY = effectiveBy + tPad;
  }

  // ── Pass 2: render ───────────────────────────────────────────────────────
  ctx.save();
  ctx.beginPath();
  ctx.rect(bx, effectiveBy, bw, effectiveBh);
  ctx.clip();

  for (const entry of allLines) {
    const { line, linePx, lineHeight, topGapPx, textXOffset, bulletLabel, bulletFont, bulletColor, bulletX, textX, textMaxW, alignment } = entry;
    cursorY += topGapPx;

    const baseline = cursorY + lineHeight * 0.8;

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

      ctx.font = seg.font;
      const segW = ctx.measureText(seg.text).width;

      if (seg.underline) {
        ctx.beginPath();
        ctx.moveTo(penX, baseline + 2);
        ctx.lineTo(penX + segW, baseline + 2);
        ctx.strokeStyle = seg.color;
        ctx.lineWidth = Math.max(1, seg.sizePx * 0.05);
        ctx.stroke();
      }

      if (seg.strikethrough) {
        ctx.beginPath();
        ctx.moveTo(penX, baseline - seg.sizePx * 0.32);
        ctx.lineTo(penX + segW, baseline - seg.sizePx * 0.32);
        ctx.strokeStyle = seg.color;
        ctx.lineWidth = Math.max(1, seg.sizePx * 0.05);
        ctx.stroke();
      }

      penX += segW;
    }

    // ── Tab-stop segments (right-aligned or centred at tab stop position) ──
    if (line.tabStop && line.tabStop.segments.length > 0) {
      const tabAbsX = bx + lPad + line.tabStop.px;
      let totalTabW = 0;
      for (const seg of line.tabStop.segments) {
        ctx.font = seg.font;
        totalTabW += ctx.measureText(seg.text).width;
      }
      let tabPenX: number;
      if (line.tabStop.algn === 'r') {
        tabPenX = tabAbsX - totalTabW;
      } else if (line.tabStop.algn === 'ctr') {
        tabPenX = tabAbsX - totalTabW / 2;
      } else {
        tabPenX = tabAbsX;
      }
      for (const seg of line.tabStop.segments) {
        ctx.font = seg.font;
        ctx.fillStyle = seg.color;
        ctx.fillText(seg.text, tabPenX, baseline);
        ctx.font = seg.font;
        tabPenX += ctx.measureText(seg.text).width;
      }
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
      if (el.rotation !== 0 || el.flipH || el.flipV) {
        ctx.translate(x + w / 2, y + h / 2);
        ctx.rotate((el.rotation * Math.PI) / 180);
        if (el.flipH) ctx.scale(-1, 1);
        if (el.flipV) ctx.scale(1, -1);
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

// ─── Chart rendering ────────────────────────────────────────────────────────

function niceStep(range: number, targetSteps = 5): number {
  const raw = range / targetSteps;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const normed = raw / mag;
  const nice = normed < 1.5 ? 1 : normed < 3.5 ? 2 : normed < 7.5 ? 5 : 10;
  return nice * mag;
}

function renderStackedBarChart(
  ctx: CanvasRenderingContext2D, el: ChartElement, scale: number
) {
  const x = emuToPx(el.x, scale);
  const y = emuToPx(el.y, scale);
  const w = emuToPx(el.width, scale);
  const h = emuToPx(el.height, scale);

  const hasTitle = el.title && el.title.trim().length > 0;
  const titleH = hasTitle ? h * 0.09 : 0;
  const padL = w * 0.11;
  const padR = w * 0.04;
  const padT = h * 0.06 + titleH;
  const padB = h * 0.13;
  const px0 = x + padL;
  const py0 = y + padT;
  const pw  = w - padL - padR;
  const ph  = h - padT - padB;

  const cats = el.categories;
  const n = cats.length;
  if (n === 0) return;

  // Compute stacked totals to find axis scale
  const totals = Array.from({ length: n }, (_, i) =>
    el.series.reduce((s, ser) => s + (ser.values[i] ?? 0), 0)
  );
  const dataMax = el.valMax ?? Math.max(...totals) * 1.05;
  if (dataMax <= 0) return;

  ctx.save();

  // Chart title
  if (hasTitle) {
    ctx.font = `bold ${Math.round(h * 0.062)}px sans-serif`;
    ctx.fillStyle = '#222';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(el.title!, x + w / 2, y + titleH / 2 + h * 0.015);
  }

  // Grid lines + Y-axis labels
  const step = niceStep(dataMax);
  const labelFontSize = Math.round(h * 0.045);
  ctx.font = `${labelFontSize}px sans-serif`;
  ctx.textAlign = 'right';
  ctx.textBaseline = 'middle';
  ctx.fillStyle = '#666';
  ctx.strokeStyle = '#e8e8e8';
  ctx.lineWidth = 0.7;
  for (let v = 0; v <= dataMax + step * 0.5; v += step) {
    const gy = py0 + ph * (1 - v / dataMax);
    ctx.beginPath();
    ctx.moveTo(px0, gy);
    ctx.lineTo(px0 + pw, gy);
    ctx.stroke();
    ctx.fillText(v.toLocaleString(), px0 - 4, gy);
  }

  // Y-axis line and X-axis baseline
  ctx.strokeStyle = '#bbb';
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(px0, py0);
  ctx.lineTo(px0, py0 + ph);
  ctx.lineTo(px0 + pw, py0 + ph);
  ctx.stroke();

  // Bars
  const barW = (pw / n) * 0.55;
  const gapW = pw / n;

  el.series.forEach((ser, si) => {
    ctx.fillStyle = ser.color ? `#${ser.color}` : `hsl(${210 + si * 40}, 60%, ${50 - si * 8}%)`;
    for (let i = 0; i < n; i++) {
      const v = ser.values[i] ?? 0;
      if (v === 0) continue;
      const base = el.series.slice(0, si).reduce((s, ps) => s + (ps.values[i] ?? 0), 0);
      const barH = ph * (v / dataMax);
      const bx = px0 + gapW * i + (gapW - barW) / 2;
      const by = py0 + ph * (1 - (base + v) / dataMax);
      ctx.fillRect(bx, by, barW, barH);
    }
  });

  // X-axis category labels
  ctx.textAlign = 'center';
  ctx.textBaseline = 'top';
  ctx.fillStyle = '#666';
  ctx.font = `${Math.round(h * 0.042)}px sans-serif`;
  for (let i = 0; i < n; i++) {
    const cx = px0 + gapW * i + gapW / 2;
    ctx.fillText(cats[i], cx, py0 + ph + 4);
  }

  // Legend (top, below title)
  if (el.series.length > 1) {
    const legY = y + titleH + h * 0.03;
    const legFontSize = Math.round(h * 0.042);
    ctx.font = `${legFontSize}px sans-serif`;
    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    let legX = px0;
    el.series.forEach((ser, si) => {
      const col = ser.color ? `#${ser.color}` : `hsl(${210 + si * 40}, 60%, ${50 - si * 8}%)`;
      ctx.fillStyle = col;
      ctx.fillRect(legX, legY - legFontSize / 2, legFontSize, legFontSize);
      ctx.fillStyle = '#444';
      ctx.fillText(ser.name, legX + legFontSize + 4, legY);
      legX += legFontSize + 4 + ctx.measureText(ser.name).width + 16;
    });
  }

  ctx.restore();
}

function renderStackedBarHChart(
  ctx: CanvasRenderingContext2D, el: ChartElement, scale: number
) {
  const x = emuToPx(el.x, scale);
  const y = emuToPx(el.y, scale);
  const w = emuToPx(el.width, scale);
  const h = emuToPx(el.height, scale);

  const padL = w * 0.22;  // category labels on left
  const padR = w * 0.04;
  const padT = h * 0.12;  // legend + top labels
  const padB = h * 0.10;  // x-axis labels
  const px0 = x + padL;
  const py0 = y + padT;
  const pw  = w - padL - padR;
  const ph  = h - padT - padB;

  const cats = el.categories;
  const n = cats.length;
  if (n === 0) return;

  const totals = Array.from({ length: n }, (_, i) =>
    el.series.reduce((s, ser) => s + (ser.values[i] ?? 0), 0)
  );
  const dataMax = el.valMax ?? Math.max(...totals) * 1.05;
  if (dataMax <= 0) return;

  ctx.save();

  // Grid lines + X-axis (value) labels
  const step = niceStep(dataMax);
  ctx.font = `${Math.round(h * 0.045)}px sans-serif`;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'top';
  ctx.fillStyle = '#666';
  ctx.strokeStyle = '#e8e8e8';
  ctx.lineWidth = 0.7;
  for (let v = 0; v <= dataMax + step * 0.5; v += step) {
    const gx = px0 + pw * (v / dataMax);
    ctx.beginPath(); ctx.moveTo(gx, py0); ctx.lineTo(gx, py0 + ph); ctx.stroke();
    ctx.fillText(v.toLocaleString(), gx, py0 + ph + 4);
  }

  // X-axis line and Y-axis baseline
  ctx.strokeStyle = '#bbb';
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(px0, py0);
  ctx.lineTo(px0, py0 + ph);
  ctx.lineTo(px0 + pw, py0 + ph);
  ctx.stroke();

  // Bars
  const barH = (ph / n) * 0.55;
  const gapH = ph / n;

  el.series.forEach((ser, si) => {
    ctx.fillStyle = ser.color ? `#${ser.color}` : `hsl(${210 + si * 40}, 60%, ${50 - si * 8}%)`;
    for (let i = 0; i < n; i++) {
      const v = ser.values[i] ?? 0;
      if (v === 0) continue;
      const base = el.series.slice(0, si).reduce((s, ps) => s + (ps.values[i] ?? 0), 0);
      const segW = pw * (v / dataMax);
      const bx = px0 + pw * (base / dataMax);
      const by = py0 + gapH * i + (gapH - barH) / 2;
      ctx.fillRect(bx, by, segW, barH);
    }
  });

  // Y-axis (category) labels
  ctx.textAlign = 'right';
  ctx.textBaseline = 'middle';
  ctx.fillStyle = '#444';
  ctx.font = `${Math.round(h * 0.042)}px sans-serif`;
  for (let i = 0; i < n; i++) {
    const cy = py0 + gapH * i + gapH / 2;
    ctx.fillText(cats[i], px0 - 6, cy);
  }

  // Legend (top)
  if (el.series.length > 1) {
    const legY = y + padT * 0.4;
    const legFontSize = Math.round(h * 0.042);
    ctx.font = `${legFontSize}px sans-serif`;
    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    let legX = px0;
    el.series.forEach((ser, si) => {
      const col = ser.color ? `#${ser.color}` : `hsl(${210 + si * 40}, 60%, ${50 - si * 8}%)`;
      ctx.fillStyle = col;
      ctx.fillRect(legX, legY - legFontSize / 2, legFontSize, legFontSize);
      ctx.fillStyle = '#444';
      ctx.fillText(ser.name, legX + legFontSize + 4, legY);
      legX += legFontSize + 4 + ctx.measureText(ser.name).width + 16;
    });
  }

  ctx.restore();
}

function renderWaterfallChart(
  ctx: CanvasRenderingContext2D, el: ChartElement, scale: number
) {
  const x = emuToPx(el.x, scale);
  const y = emuToPx(el.y, scale);
  const w = emuToPx(el.width, scale);
  const h = emuToPx(el.height, scale);

  const padL = w * 0.11;
  const padR = w * 0.04;
  const padT = h * 0.08;
  const padB = h * 0.18;
  const px0 = x + padL;
  const py0 = y + padT;
  const pw  = w - padL - padR;
  const ph  = h - padT - padB;

  const vals = el.series[0]?.values ?? [];
  const cats = el.categories;
  const n = cats.length;
  if (n === 0) return;

  const subSet = new Set(el.subtotalIndices);

  // Build bar segments: {start, end, isSub, isPos}
  let running = 0;
  const bars: Array<{ start: number; end: number; isSub: boolean; isPos: boolean }> = [];
  for (let i = 0; i < n; i++) {
    const v = vals[i] ?? 0;
    const isSub = i === 0 || subSet.has(i);
    if (isSub) {
      bars.push({ start: 0, end: v, isSub: true, isPos: true });
      running = v;
    } else {
      const start = v >= 0 ? running : running + v;
      const end   = v >= 0 ? running + v : running;
      bars.push({ start, end, isSub: false, isPos: v >= 0 });
      running += v;
    }
  }

  // Axis range
  const allEnds = bars.map(b => b.end);
  const allStarts = bars.map(b => b.start);
  const rawMax = Math.max(...allEnds, ...allStarts);
  const rawMin = Math.min(...allStarts, 0);
  const dataRange = rawMax - rawMin;
  if (dataRange <= 0) return;
  const padded = dataRange * 1.1;
  const dataMin = rawMin - dataRange * 0.05;
  const dataMax = dataMin + padded;

  // Grid
  const step = niceStep(padded);
  ctx.save();
  const fontSize = Math.round(h * 0.042);
  ctx.font = `${fontSize}px sans-serif`;
  ctx.strokeStyle = '#e8e8e8';
  ctx.lineWidth = 0.7;
  ctx.fillStyle = '#666';
  ctx.textAlign = 'right';
  ctx.textBaseline = 'middle';
  for (let v = Math.ceil(dataMin / step) * step; v <= dataMax; v += step) {
    const gy = py0 + ph * (1 - (v - dataMin) / padded);
    ctx.beginPath(); ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy); ctx.stroke();
    ctx.fillText(v.toLocaleString(), px0 - 4, gy);
  }

  // Y-axis line and X-axis baseline (at dataMin level = y-axis origin)
  ctx.strokeStyle = '#bbb';
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(px0, py0);
  ctx.lineTo(px0, py0 + ph);
  ctx.lineTo(px0 + pw, py0 + ph);
  ctx.stroke();

  // Colors
  const colorSub = '#196ECA';
  const colorPos = '#5BA4E6';
  const colorNeg = '#E46970';

  const barW = (pw / n) * 0.55;
  const gapW = pw / n;

  bars.forEach((bar, i) => {
    const bx = px0 + gapW * i + (gapW - barW) / 2;
    const yTop = py0 + ph * (1 - (bar.end - dataMin) / padded);
    const yBot = py0 + ph * (1 - (bar.start - dataMin) / padded);
    const bh = Math.max(1, yBot - yTop);

    if (bar.isSub) {
      ctx.fillStyle = colorSub;
      ctx.fillRect(bx, yTop, barW, bh);
    } else {
      // Delta bars: outlined (hollow) style
      ctx.strokeStyle = bar.isPos ? colorPos : colorNeg;
      ctx.lineWidth = 1.5;
      ctx.strokeRect(bx + 0.75, yTop + 0.75, barW - 1.5, bh - 1.5);
    }

    // Connector line to next bar
    if (i < n - 1) {
      const nextBx = px0 + gapW * (i + 1) + (gapW - barW) / 2;
      const connY = bar.isPos ? yTop : yBot;
      ctx.strokeStyle = '#ccc';
      ctx.lineWidth = 0.8;
      ctx.setLineDash([3, 3]);
      ctx.beginPath();
      ctx.moveTo(bx + barW, connY);
      ctx.lineTo(nextBx, connY);
      ctx.stroke();
      ctx.setLineDash([]);
    }

    // Value label above each bar
    const rawVal = vals[i] ?? 0;
    const labelText = rawVal < 0
      ? `△ ${Math.abs(rawVal).toLocaleString()}`
      : rawVal.toLocaleString();
    ctx.fillStyle = '#595959';
    ctx.font = `bold ${Math.round(h * 0.044)}px sans-serif`;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'bottom';
    ctx.fillText(labelText, bx + barW / 2, yTop - 3);
  });

  // X-axis labels (may contain newlines → split)
  ctx.textAlign = 'center';
  ctx.textBaseline = 'top';
  ctx.fillStyle = '#666';
  ctx.font = `${Math.round(h * 0.038)}px sans-serif`;
  const labelY = py0 + ph + 4;
  for (let i = 0; i < n; i++) {
    const cx = px0 + gapW * i + gapW / 2;
    const lines = cats[i].split(/\s+/);
    lines.forEach((line, li) => ctx.fillText(line, cx, labelY + li * (fontSize + 2)));
  }

  ctx.restore();
}

function renderChart(ctx: CanvasRenderingContext2D, el: ChartElement, scale: number) {
  if (el.chartType === 'stackedBar' || el.chartType === 'clusteredBar') {
    renderStackedBarChart(ctx, el, scale);
  } else if (el.chartType === 'stackedBarH' || el.chartType === 'clusteredBarH') {
    renderStackedBarHChart(ctx, el, scale);
  } else if (el.chartType === 'waterfall') {
    renderWaterfallChart(ctx, el, scale);
  }
}

// ─── Table rendering ─────────────────────────────────────────────────────────

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
  /** Theme default text color (dk1), used as fallback when shapes have no explicit color */
  defaultTextColor?: string | null;
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

  const themeDefaultColor = opts.defaultTextColor
    ? `#${opts.defaultTextColor}`
    : '#000000';

  for (const el of slide.elements) {
    if (el.type === 'shape') {
      renderShape(ctx, el, scale, themeDefaultColor);
    } else if (el.type === 'picture') {
      await renderPicture(ctx, el, scale);
    } else if (el.type === 'table') {
      renderTable(ctx, el, scale);
    } else if (el.type === 'chart') {
      renderChart(ctx, el, scale);
    }
  }

  return canvas;
}

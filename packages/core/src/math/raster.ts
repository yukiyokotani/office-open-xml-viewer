import { recolorSvg, type MathSvg } from './mathjax.js';
import { createAuxCanvas, type AuxContext } from '../canvas/aux-canvas.js';
import { clampCanvasSize } from '../canvas/clamp.js';

export const MATH_RASTER_PX_PER_EM = 256;
const MATH_RASTER_CACHE_PX_PER_EM = 64;

export interface RasterizedMathSvg {
  readonly source: CanvasImageSource;
  readonly widthPx: number;
  readonly heightPx: number;
}

/** Give the standalone SVG deterministic intrinsic pixels for its auxiliary
 * Canvas backing store. MathJax otherwise emits `ex` dimensions whose raster
 * resolution can vary by realm and then blur on HiDPI output. */
export function sizeMathSvgForRaster(
  svg: string,
  widthEm: number,
  heightEm: number,
): string {
  const widthPx = Math.max(1, Math.round(widthEm * MATH_RASTER_PX_PER_EM));
  const heightPx = Math.max(1, Math.round(heightEm * MATH_RASTER_PX_PER_EM));
  return svg.replace(/<svg([^>]*?)>/, (_match, attributes: string) => {
    const clean = attributes.replace(/\s(?:width|height)="[^"]*"/g, '');
    return `<svg${clean} width="${widthPx}" height="${heightPx}">`;
  });
}

interface PaintState {
  fill: boolean;
  stroke: boolean;
  fillRule: CanvasFillRule;
}

function attribute(source: string, name: string): string | undefined {
  const match = new RegExp(`(?:^|\\s)${name}="([^"]*)"`).exec(source);
  return match?.[1];
}

function numericAttribute(source: string, name: string, fallback = 0): number {
  const value = Number.parseFloat(attribute(source, name) ?? '');
  return Number.isFinite(value) ? value : fallback;
}

function decodeXmlText(value: string): string {
  return value.replace(/&(#(?:x[\da-fA-F]+|\d+)|amp|lt|gt|quot|apos);/g, (entity, code: string) => {
    if (code[0] === '#') {
      const point = code[1] === 'x'
        ? Number.parseInt(code.slice(2), 16)
        : Number.parseInt(code.slice(1), 10);
      return point <= 0x10ffff ? String.fromCodePoint(point) : entity;
    }
    return { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }[code as 'amp' | 'lt' | 'gt' | 'quot' | 'apos'];
  });
}

function presentationAttributes(source: string): Map<string, string> {
  const properties = new Map<string, string>();
  for (const name of [
    'fill',
    'stroke',
    'stroke-width',
    'stroke-linecap',
    'stroke-linejoin',
    'fill-rule',
    'opacity',
  ]) {
    const value = attribute(source, name);
    if (value !== undefined) properties.set(name, value);
  }
  const style = attribute(source, 'style');
  if (style) {
    for (const declaration of style.split(';')) {
      const separator = declaration.indexOf(':');
      if (separator < 0) continue;
      const name = declaration.slice(0, separator).trim();
      const value = declaration.slice(separator + 1).trim();
      if (name && value) properties.set(name, value);
    }
  }
  return properties;
}

function applyPresentation(
  context: AuxContext,
  source: string,
  inherited: PaintState,
): PaintState {
  const properties = presentationAttributes(source);
  const state = { ...inherited };
  const fill = properties.get('fill');
  if (fill === 'none') state.fill = false;
  else if (fill) {
    state.fill = true;
    context.fillStyle = fill;
  }
  const stroke = properties.get('stroke');
  if (stroke === 'none') state.stroke = false;
  else if (stroke) {
    state.stroke = true;
    context.strokeStyle = stroke;
  }
  const strokeWidth = Number.parseFloat(properties.get('stroke-width') ?? '');
  if (Number.isFinite(strokeWidth)) context.lineWidth = strokeWidth;
  const lineCap = properties.get('stroke-linecap');
  if (lineCap === 'butt' || lineCap === 'round' || lineCap === 'square') {
    context.lineCap = lineCap;
  }
  const lineJoin = properties.get('stroke-linejoin');
  if (lineJoin === 'bevel' || lineJoin === 'round' || lineJoin === 'miter') {
    context.lineJoin = lineJoin;
  }
  const fillRule = properties.get('fill-rule');
  if (fillRule === 'evenodd' || fillRule === 'nonzero') state.fillRule = fillRule;
  const opacity = Number.parseFloat(properties.get('opacity') ?? '');
  if (Number.isFinite(opacity)) context.globalAlpha *= Math.max(0, Math.min(1, opacity));
  return state;
}

function applyTransform(context: AuxContext, transform: string | undefined): void {
  if (!transform) return;
  const operations = /([a-zA-Z]+)\(([^)]*)\)/g;
  for (let match = operations.exec(transform); match; match = operations.exec(transform)) {
    const values = match[2]
      .trim()
      .split(/[\s,]+/)
      .filter(Boolean)
      .map(Number);
    if (values.some((value) => !Number.isFinite(value))) continue;
    switch (match[1]) {
      case 'matrix':
        if (values.length >= 6) context.transform(...values.slice(0, 6) as [number, number, number, number, number, number]);
        break;
      case 'translate':
        context.translate(values[0] ?? 0, values[1] ?? 0);
        break;
      case 'scale': {
        const x = values[0] ?? 1;
        context.scale(x, values[1] ?? x);
        break;
      }
      case 'rotate': {
        const radians = ((values[0] ?? 0) * Math.PI) / 180;
        if (values.length >= 3) {
          context.translate(values[1], values[2]);
          context.rotate(radians);
          context.translate(-values[1], -values[2]);
        } else {
          context.rotate(radians);
        }
        break;
      }
    }
  }
}

function paintPath(context: AuxContext, path: Path2D, state: PaintState): void {
  if (state.fill) context.fill(path, state.fillRule);
  if (state.stroke && context.lineWidth > 0) context.stroke(path);
}

/** Draw the deliberately small SVG vocabulary emitted by the bundled MathJax
 * SVG output. This avoids `createImageBitmap(svgBlob)`, which Chromium workers
 * do not decode, and keeps Window/Worker pixels on the same Canvas path.
 * Unsupported STIX2 ranges use MathJax's system-font <text> fallback, which
 * must be painted here as well or the glyph silently disappears in workers. */
export function drawMathJaxSvg(
  context: AuxContext,
  svg: string,
  widthPx: number,
  heightPx: number,
): void {
  const viewBox = attribute(svg.slice(0, svg.indexOf('>') + 1), 'viewBox')
    ?.trim()
    .split(/[\s,]+/)
    .map(Number);
  if (!viewBox || viewBox.length !== 4 || viewBox.some((value) => !Number.isFinite(value))) {
    throw new TypeError('MathJax SVG must contain a finite viewBox');
  }
  const [minX, minY, viewWidth, viewHeight] = viewBox;
  if (!(viewWidth > 0) || !(viewHeight > 0)) {
    throw new TypeError('MathJax SVG viewBox must have positive dimensions');
  }

  context.save();
  context.setTransform(
    widthPx / viewWidth,
    0,
    0,
    heightPx / viewHeight,
    (-minX * widthPx) / viewWidth,
    (-minY * heightPx) / viewHeight,
  );
  context.fillStyle = '#000000';
  context.strokeStyle = '#000000';
  context.lineWidth = 1;
  let state: PaintState = { fill: true, stroke: false, fillRule: 'nonzero' };
  const stack: PaintState[] = [];
  const tags = /<\/?([A-Za-z][\w:-]*)([^>]*)>/g;
  for (let match = tags.exec(svg); match; match = tags.exec(svg)) {
    const [tag, rawName, attributes = ''] = match;
    const name = rawName.toLowerCase();
    const closing = tag.startsWith('</');
    if (name === 'svg') continue;
    if (closing) {
      if (name === 'g') {
        context.restore();
        state = stack.pop() ?? state;
      }
      continue;
    }
    if (name === 'g') {
      stack.push(state);
      context.save();
      state = applyPresentation(context, attributes, state);
      applyTransform(context, attribute(attributes, 'transform'));
      continue;
    }
    if (name === 'text') {
      const end = svg.indexOf('</text>', tags.lastIndex);
      if (end < 0) throw new TypeError('MathJax SVG text must have a closing tag');
      context.save();
      const textState = applyPresentation(context, attributes, state);
      applyTransform(context, attribute(attributes, 'transform'));
      const style = attribute(attributes, 'font-style') === 'italic' ? 'italic ' : '';
      const weight = attribute(attributes, 'font-weight') === 'bold' ? 'bold ' : '';
      const size = numericAttribute(attributes, 'font-size', 1000);
      const family = attribute(attributes, 'font-family') ?? 'serif';
      context.font = `${style}${weight}${size}px ${family}`;
      context.textBaseline = 'alphabetic';
      const content = decodeXmlText(svg.slice(tags.lastIndex, end));
      const x = numericAttribute(attributes, 'x');
      const y = numericAttribute(attributes, 'y');
      if (textState.fill) context.fillText(content, x, y);
      if (textState.stroke && context.lineWidth > 0) context.strokeText(content, x, y);
      context.restore();
      tags.lastIndex = end + '</text>'.length;
      continue;
    }
    if (name !== 'path' && name !== 'rect' && name !== 'line') continue;
    context.save();
    const shapeState = applyPresentation(context, attributes, state);
    applyTransform(context, attribute(attributes, 'transform'));
    const path = name === 'path'
      ? new Path2D(attribute(attributes, 'd') ?? '')
      : new Path2D();
    if (name === 'rect') {
      path.rect(
        numericAttribute(attributes, 'x'),
        numericAttribute(attributes, 'y'),
        numericAttribute(attributes, 'width'),
        numericAttribute(attributes, 'height'),
      );
    } else if (name === 'line') {
      path.moveTo(numericAttribute(attributes, 'x1'), numericAttribute(attributes, 'y1'));
      path.lineTo(numericAttribute(attributes, 'x2'), numericAttribute(attributes, 'y2'));
    }
    paintPath(context, path, shapeState);
    context.restore();
  }
  context.restore();
}

/** Rasterize one MathJax SVG in Window or Worker without DOM-only image
 * decoding. Both realms use the same Canvas path renderer for pixel parity. */
export async function rasterizeMathSvg(
  output: MathSvg,
  color: string,
): Promise<RasterizedMathSvg> {
  const requestedWidth = output.widthEm * MATH_RASTER_PX_PER_EM;
  const requestedHeight = (output.ascentEm + output.descentEm) * MATH_RASTER_PX_PER_EM;
  const { width: widthPx, height: heightPx } = clampCanvasSize(
    requestedWidth,
    requestedHeight,
  );
  const canvas = createAuxCanvas(widthPx, heightPx);
  const context = canvas?.getContext('2d');
  if (!canvas || !context || typeof Path2D === 'undefined') {
    throw new Error('Math SVG rasterization requires Canvas 2D and Path2D support');
  }
  const svg = sizeMathSvgForRaster(
    recolorSvg(output.svg, color),
    output.widthEm,
    output.ascentEm + output.descentEm,
  );
  drawMathJaxSvg(context, svg, widthPx, heightPx);
  let source: CanvasImageSource = canvas;
  let cachedWidthPx = widthPx;
  let cachedHeightPx = heightPx;
  // Chromium does not materially distinguish `low` and `high` quality when a
  // very dense CanvasImageSource is reduced directly to typical text size.
  // Build a small mip level once, while the resource is prepared, so every
  // later Window/Worker paint performs only a modest final reduction. Two 2x
  // passes retain the 256 px/em path detail without paying that 16x reduction
  // on every frame.
  const targetScale = MATH_RASTER_CACHE_PX_PER_EM / MATH_RASTER_PX_PER_EM;
  const targetWidthPx = Math.max(1, Math.round(widthPx * targetScale));
  const targetHeightPx = Math.max(1, Math.round(heightPx * targetScale));
  while (cachedWidthPx > targetWidthPx || cachedHeightPx > targetHeightPx) {
    const nextWidthPx = Math.max(targetWidthPx, Math.ceil(cachedWidthPx / 2));
    const nextHeightPx = Math.max(targetHeightPx, Math.ceil(cachedHeightPx / 2));
    const reduced = createAuxCanvas(nextWidthPx, nextHeightPx);
    const reducedContext = reduced?.getContext('2d');
    if (!reduced || !reducedContext) break;
    reducedContext.imageSmoothingEnabled = true;
    reducedContext.imageSmoothingQuality = 'high';
    reducedContext.drawImage(source, 0, 0, nextWidthPx, nextHeightPx);
    source = reduced;
    cachedWidthPx = nextWidthPx;
    cachedHeightPx = nextHeightPx;
  }
  return { source, widthPx: cachedWidthPx, heightPx: cachedHeightPx };
}

/** Recolor a cached black equation raster synchronously for the authored run
 * color. Uses OffscreenCanvas in workers and a detached canvas fallback on
 * Window. */
export function tintMathRaster(
  raster: RasterizedMathSvg,
  color: string,
): CanvasImageSource {
  const canvas = createAuxCanvas(raster.widthPx, raster.heightPx);
  if (!canvas) return raster.source;
  const context = canvas.getContext('2d');
  if (!context) return raster.source;
  context.drawImage(raster.source, 0, 0, raster.widthPx, raster.heightPx);
  context.globalCompositeOperation = 'source-in';
  context.fillStyle = color;
  context.fillRect(0, 0, raster.widthPx, raster.heightPx);
  return canvas;
}

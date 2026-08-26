import { intersectElementRects } from './dom-geometry.js';

/** Internal geometry used by the built-in DOCX/PPTX comment connectors. */
export interface ReadOnlyCommentRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export interface ReadOnlyCommentThreadGeometry {
  readonly occurrenceKey: string;
  readonly active: boolean;
  readonly anchorRects: readonly ReadOnlyCommentRect[];
  readonly cardRect?: ReadOnlyCommentRect;
}

/** Geometry captured by a full comment projection. Card rectangles are stored
 * before margin clipping so a scroll-only update can reveal a previously
 * clipped card without rebuilding comment DOM. */
export interface ReadOnlyCommentMarginGeometry {
  readonly threads: readonly ReadOnlyCommentThreadGeometry[];
  readonly cardClipBounds?: ReadOnlyCommentRect;
  readonly scrollTop: number;
}

/** Project cached card geometry through the current margin scroll offset. */
export function projectReadOnlyCommentMarginScroll(
  geometry: ReadOnlyCommentMarginGeometry,
  scrollTop: number,
): readonly ReadOnlyCommentThreadGeometry[] {
  const deltaY = geometry.scrollTop - scrollTop;
  return Object.freeze(geometry.threads.map((thread) => {
    if (!thread.cardRect) return thread;
    const shifted = Object.freeze({
      ...thread.cardRect,
      y: thread.cardRect.y + deltaY,
    });
    const cardRect = geometry.cardClipBounds
      ? intersectElementRects(shifted, geometry.cardClipBounds)
      : shifted;
    const { cardRect: _cachedCardRect, ...rest } = thread;
    return Object.freeze({
      ...rest,
      ...(cardRect ? { cardRect } : {}),
    });
  }));
}

export interface ReadOnlyCommentDecorationSnapshot {
  readonly surfaceBounds: ReadOnlyCommentRect;
  readonly contentBounds: ReadOnlyCommentRect;
  readonly side: 'left' | 'right';
  readonly threads: readonly ReadOnlyCommentThreadGeometry[];
}

export type ReadOnlyCommentConnectorRoute = 'bezier' | 'orthogonal';
export type ReadOnlyCommentConnectorStroke = 'solid' | 'dashed';

export interface ReadOnlyCommentDecorationOptions {
  readonly route: ReadOnlyCommentConnectorRoute;
  readonly stroke: ReadOnlyCommentConnectorStroke;
  readonly color?: string;
  readonly activeColor?: string;
}

const SVG_NS = 'http://www.w3.org/2000/svg';

interface DecorationState {
  readonly svg: SVGSVGElement;
  readonly paths: Map<string, SVGPathElement>;
}

const stateByLayer = new WeakMap<HTMLDivElement, DecorationState>();

interface Point { readonly x: number; readonly y: number }

function n(value: number): string {
  const rounded = Math.round(value * 1000) / 1000;
  return Object.is(rounded, -0) ? '0' : String(rounded);
}

function bezierControls(start: Point, end: Point): readonly [Point, Point] {
  const control = (end.x - start.x) * 0.5;
  return [
    { x: start.x + control, y: start.y },
    { x: end.x - control, y: end.y },
  ];
}

/** Generate one connector path. Exported only for focused geometry tests. */
export function readOnlyCommentConnectorPath(
  start: Point,
  end: Point,
  route: ReadOnlyCommentConnectorRoute,
): string {
  const elbowX = start.x + (end.x - start.x) * 0.55;
  if (route === 'orthogonal') {
    return `M ${n(start.x)} ${n(start.y)} H ${n(elbowX)} V ${n(end.y)} H ${n(end.x)}`;
  }
  const [control1, control2] = bezierControls(start, end);
  return `M ${n(start.x)} ${n(start.y)} C ${n(control1.x)} ${n(control1.y)}, ` +
    `${n(control2.x)} ${n(control2.y)}, ${n(end.x)} ${n(end.y)}`;
}

export function disposeReadOnlyCommentDecoration(layer: HTMLDivElement): void {
  stateByLayer.delete(layer);
  layer.replaceChildren();
}

/** Draw the built-in page/slide-to-card connectors in surface CSS pixels. */
export function buildReadOnlyCommentDecoration(
  layer: HTMLDivElement,
  snapshot: ReadOnlyCommentDecorationSnapshot,
  options: ReadOnlyCommentDecorationOptions,
): void {
  layer.dataset.ooxmlCommentConnectors = '';
  let state = stateByLayer.get(layer);
  if (!state) {
    const svg = layer.ownerDocument.createElementNS(SVG_NS, 'svg');
    svg.setAttribute('aria-hidden', 'true');
    svg.style.cssText =
      'position:absolute;inset:0;width:100%;height:100%;overflow:visible;pointer-events:none;';
    state = { svg, paths: new Map() };
    stateByLayer.set(layer, state);
    layer.replaceChildren(svg);
  }
  state.svg.setAttribute(
    'viewBox',
    `${snapshot.surfaceBounds.x} ${snapshot.surfaceBounds.y} ` +
      `${snapshot.surfaceBounds.width} ${snapshot.surfaceBounds.height}`,
  );

  const orderedPaths: SVGPathElement[] = [];
  const desired = new Set<string>();
  for (const thread of snapshot.threads) {
    const anchor = thread.anchorRects.at(-1);
    const card = thread.cardRect;
    if (!anchor || !card) continue;
    desired.add(thread.occurrenceKey);

    const startX = snapshot.side === 'left' ? anchor.x : anchor.x + anchor.width;
    const startY = anchor.y + anchor.height / 2;
    const endX = snapshot.side === 'left' ? card.x + card.width : card.x;
    const endY = card.y + Math.min(card.height / 2, 25);
    let path = state.paths.get(thread.occurrenceKey);
    if (!path) {
      path = layer.ownerDocument.createElementNS(SVG_NS, 'path');
      state.paths.set(thread.occurrenceKey, path);
    }
    path.dataset.ooxmlCommentConnector = thread.occurrenceKey;
    path.dataset.active = String(thread.active);
    path.setAttribute(
      'd',
      readOnlyCommentConnectorPath(
        { x: startX, y: startY },
        { x: endX, y: endY },
        options.route,
      ),
    );
    path.style.cssText =
      'fill:none;vector-effect:non-scaling-stroke;' +
      `stroke:${thread.active
        ? options.activeColor ?? options.color ?? '#2563eb'
        : options.color ?? '#94a3b8'};` +
      `stroke-width:${thread.active ? '1.5px' : '1px'};` +
      `stroke-dasharray:${options.stroke === 'dashed' ? '4 4' : 'none'};` +
      `opacity:${thread.active ? '.9' : '.45'};`;
    orderedPaths.push(path);
  }

  for (const [key, path] of [...state.paths]) {
    if (desired.has(key)) continue;
    state.paths.delete(key);
    path.remove();
  }
  const orderChanged = orderedPaths.length !== state.svg.children.length ||
    orderedPaths.some((path, index) => state.svg.children[index] !== path);
  if (orderChanged) state.svg.replaceChildren(...orderedPaths);
}

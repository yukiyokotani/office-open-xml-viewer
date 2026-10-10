import type { ShadeBox } from './path-gradient';

import { appendPathStateCommand, strokePathStateBounds, type PathState } from './path-enclosure';
const states = new WeakMap<CanvasRenderingContext2D, PathState>();
const sources = new WeakMap<CanvasRenderingContext2D, CanvasRenderingContext2D>();
export const paintPathSource = (ctx: CanvasRenderingContext2D): CanvasRenderingContext2D => sources.get(ctx) ?? ctx;
const wrappers = new WeakMap<CanvasRenderingContext2D, CanvasRenderingContext2D>();

/** Observe the actual Canvas path without changing its commands or state.
 * Canvas exposes no current-path bounds API. Keep only the current path,
 * recording device-space control hulls and exact join tangents. Curves use
 * their conservative convex hull; straight segments, caps and miters use
 * the actual stroked geometry. No line-width/decoration-size guess is used. */
export function trackPaintPath(ctx: CanvasRenderingContext2D): CanvasRenderingContext2D {
  if (states.has(ctx)) return ctx;
  const existing = wrappers.get(ctx);
  if (existing) return existing;
  const state: PathState = { paths: [] };
  const transform = () => typeof ctx.getTransform === 'function' ? ctx.getTransform() : { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
  const observe = (name: string, args: number[]) => appendPathStateCommand(state, transform(), name, args);
  const proxy = new Proxy(ctx, {
    get(target, property) {
      const value = Reflect.get(target, property, target);
      if (typeof value !== 'function') return value;
      return (...args: number[]) => {
        if (typeof property === 'string') observe(property, args);
        return Reflect.apply(value, target, args);
      };
    },
    set(target, property, value) { return Reflect.set(target, property, value, target); },
  });
  states.set(proxy, state); sources.set(proxy, ctx); wrappers.set(ctx, proxy);
  return proxy;
}

/** Conservative device bounds of the recorded path's stroked outline.
 * Every dash has caps, including interior dashes on closed curves. A square
 * cap's corners are p + r(±t ±n), with unit tangent t and normal n. For a
 * device-axis row v of the affine transform, their support is
 * r(|v·t| + |v·n|) <= r sqrt(2) |v|. This envelopes every possible end tangent
 * on a curved control hull without flattening or approximating dash lengths.
 * Straight segments use their exact tangent. Round caps use circular support;
 * flat caps add no tangent extension. Neither envelope changes the shade frame. */
export function currentStrokeBounds(ctx: CanvasRenderingContext2D): ShadeBox | undefined {
  const state = states.get(ctx);
  if (!state || typeof ctx.getTransform !== 'function') return undefined;
  const matrix = ctx.getTransform();
  const determinant = matrix.a * matrix.d - matrix.b * matrix.c;
  // Preserve the observational adapter's ordinary singular-transform fallback.
  if (!Number.isFinite(determinant) || determinant === 0) return undefined;
  return strokePathStateBounds(state, matrix, {
    lineWidth: ctx.lineWidth, lineCap: ctx.lineCap, lineJoin: ctx.lineJoin,
    miterLimit: ctx.miterLimit, dash: typeof ctx.getLineDash === 'function' ? ctx.getLineDash() : [],
  });
}

/** Public resolver compatibility when no current path was supplied: coverage
 * of a stroked host rectangle, transformed as geometry (including its joins).
 * Custom strokes must pass their recorded current-path bounds. */
export function hostStrokeBounds(ctx: CanvasRenderingContext2D, box: ShadeBox): ShadeBox | undefined {
  if (typeof ctx.getTransform !== 'function') return undefined;
  const m = ctx.getTransform(); const r = (ctx.lineWidth ?? 1) / 2;
  const miter = ctx.lineJoin === 'miter' && ctx.miterLimit >= Math.SQRT2;
  const hx = r * (miter ? Math.abs(m.a) + Math.abs(m.c) : Math.hypot(m.a, m.c));
  const hy = r * (miter ? Math.abs(m.b) + Math.abs(m.d) : Math.hypot(m.b, m.d));
  const points = [[box.x, box.y], [box.x + box.w, box.y], [box.x, box.y + box.h], [box.x + box.w, box.y + box.h]];
  const xs = points.map(([x, y]) => m.a * x + m.c * y + m.e);
  const ys = points.map(([x, y]) => m.b * x + m.d * y + m.f);
  return { x: Math.min(...xs) - hx, y: Math.min(...ys) - hy,
    w: Math.max(...xs) - Math.min(...xs) + 2 * hx, h: Math.max(...ys) - Math.min(...ys) + 2 * hy };
}

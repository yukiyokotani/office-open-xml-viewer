/**
 * Trailing-whitespace-robust text advance measurement.
 *
 * `CanvasRenderingContext2D.measureText(s).width` for a string that INCLUDES
 * trailing (or lone) whitespace is engine-dependent:
 *
 *   - Chrome / Blink and skia-canvas INCLUDE the trailing-space advance.
 *   - Firefox and WebKit / Safari EXCLUDE it — a run's trailing whitespace (and a
 *     string that is ENTIRELY whitespace, e.g. a lone " ") contributes 0 advance.
 *
 * The HTML spec left this ambiguous for years and the engines never converged.
 * Every OOXML renderer here stores a laid-out text advance as this width and then
 * advances the draw pen by the same amount, so on a trimming engine an inter-word
 * space that sits at a segment's edge loses its advance and the next word
 * collides ("Other countries" → "Othercountries"). The same divergence bites all
 * three formats — docx (word segments keep their trailing space), pptx (a lone
 * " " token measures 0) and xlsx (a wrap buffer flushed at a space) — so the
 * correction lives here, in core, as a single shared source of truth.
 *
 * The fix detects ONCE per (context, font) whether the engine trims, and — only
 * when it does — adds back `trailingSpaceCount × spaceAdvance`, where
 * `spaceAdvance` is derived from an INTERIOR space ("m m" − "mm") that no engine
 * trims. On a non-trimming engine the detector reports `trims=false`, so the
 * helper is a pure pass-through (`ctx.measureText(text).width`) with zero
 * behaviour change — and for the common no-trailing-space string it never even
 * runs the detector probe.
 */

type AnyCtx = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

/** Per-context detector cache. Keyed by the 2D context, holding one entry per
 *  resolved font string. A WeakMap lets a context be GC'd with its cache. */
const detectorCache = new WeakMap<AnyCtx, Map<string, TrailingSpaceInfo>>();

interface TrailingSpaceInfo {
  /** True when this engine excludes trailing whitespace from measureText width. */
  trims: boolean;
  /** The interior (never-trimmed) advance of one space in the current font, px. */
  spaceAdvance: number;
}

/** Detect (memoized per context+font) whether the engine trims trailing
 *  whitespace from `measureText` width, and the true per-space advance. */
function trailingSpaceInfo(ctx: AnyCtx): TrailingSpaceInfo {
  let perFont = detectorCache.get(ctx);
  if (!perFont) {
    perFont = new Map();
    detectorCache.set(ctx, perFont);
  }
  const font = ctx.font;
  let info = perFont.get(font);
  if (!info) {
    // Interior space advance (never trimmed): width('m m') − width('mm').
    const interior = ctx.measureText('m m').width - ctx.measureText('mm').width;
    // Trailing space advance: width('m ') − width('m'). ~0 on a trimming engine.
    const trailing = ctx.measureText('m ').width - ctx.measureText('m').width;
    // Trims when a genuine interior advance exists but the trailing one vanished.
    const trims = interior > 0.01 && trailing < interior * 0.5;
    info = { trims, spaceAdvance: interior };
    perFont.set(font, info);
  }
  return info;
}

/**
 * Advance of `text` under the font ALREADY set on `ctx`, corrected so trailing
 * (and lone) spaces keep their advance on engines whose `measureText` trims
 * trailing whitespace. On non-trimming engines this equals
 * `ctx.measureText(text).width` exactly.
 *
 * Note the empty-string / all-space case: on a trimming engine
 * `measureText("  ")` returns 0, and this helper adds back
 * `spaceCount × spaceAdvance`, so a lone-space token (pptx) or an all-space
 * segment measures its true width.
 */
export function measureAdvanceWithTrailingSpace(ctx: AnyCtx, text: string): number {
  if (text.length === 0) return 0;
  const raw = ctx.measureText(text).width;
  // Count trailing spaces; this also covers the wholly-space string (a lone " "
  // token: raw is 0 on a trimming engine, so every space must be added back).
  const trail = /( +)$/.exec(text);
  if (!trail) return raw;
  const info = trailingSpaceInfo(ctx);
  if (!info.trims) return raw;
  return raw + trail[1].length * info.spaceAdvance;
}

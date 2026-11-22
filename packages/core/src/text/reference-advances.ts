/**
 * Size-proportional glyph advances for Canvas text.
 *
 * Word lays out text from design advances: a glyph's advance is its font-unit
 * advance width (OpenType `hmtx`) scaled linearly by the em size, so a run is
 * exactly twice as wide at twice the size. A Canvas implementation may instead
 * report grid-fitted advances at small sizes and shape `fillText` from those
 * same advances. Observed in Chromium-based browsers on Windows (Edge on
 * Windows 11): a face with embedded bitmap strikes (OpenType `EBLC`/`EBDT`,
 * which many CJK system faces carry) is measured in GDI-compatible mode at the
 * pixel sizes its strikes cover, returning whole-pixel advances; for one such
 * face U+3000 measured 16px at 14.667px, 14px at 13.333px and exactly 100px at
 * 100px, independent of `textRendering`. Faces without strikes measured
 * proportionally at every size, and macOS measured every tested face
 * proportionally. Integer strike advances are a rasterizer artefact, not font
 * data: laying out from them makes text wider than Word's on that host.
 *
 * Chrome on macOS shows the same artefact for an `sbix` colour bitmap face:
 * Apple Color Emoji's digit, `#` and emoji glyphs measured 13px at 10px and
 * 14px at 11px (proportional: 10 and 11).
 *
 * Library policy: a font shorthand whose host advances at its own size are
 * grid-fitted, i.e. depart from the advances at {@link TEXT_ADVANCE_REFERENCE_PX}
 * scaled to that size and land on whole pixels, is measured and painted at
 * that reference em under a linear scale. Advances that depart by a fractional
 * amount are font data, not grid fitting: an optical-size axis (`opsz`, e.g.
 * macOS system-ui, whose 11px 'a' measures 6.13 against 5.52 scaled from
 * 256px) or a size-dependent tracking table selects a different design at a
 * different size, so such a face keeps its own size. The device raster
 * size is unchanged by the scale, so the host still selects its glyph images
 * (a bitmap strike included) for the displayed size; only their positions
 * follow the proportional advances, as in Word. Shorthands the host measures
 * proportionally keep measuring and painting at their own size: drawing at
 * the reference em is not raster-neutral (Chrome on macOS, measured: synthetic
 * bold loses about a tenth of its ink because the emboldening stroke follows
 * the user-space size, and antialiasing shifts slightly), so it is used only
 * where the advances require it. No face name or per-face constant takes part
 * in the decision; it is a probe of the measuring context itself.
 */

import { saveCanvasState } from '../internal/canvas-language.js';

/**
 * Reference em in CSS px. Embedded bitmap strikes are indexed by a uint8 ppem
 * (OpenType `EBLC`/`CBLC` BitmapSize.ppemX/ppemY) and `hdmx` device records by
 * a uint8 pixel size, so no face can carry a strike or device advance at 256px
 * or above. As a power of two, the factor mapping reference metrics onto a
 * requested size only shifts binary exponents and adds a single rounding.
 */
export const TEXT_ADVANCE_REFERENCE_PX = 256;

/** A font shorthand rewritten to the reference em, with the factor that maps
 * its metrics (and drawing) back onto the requested size. */
export interface ReferenceTextFont {
  /** The same CSS font shorthand at {@link TEXT_ADVANCE_REFERENCE_PX}. */
  readonly font: string;
  /** Requested px per reference px. */
  readonly scale: number;
}

/** The absolute `<number>px` size of a CSS font shorthand. The size is the
 * first token followed by `px` (an optional `/line-height` may follow);
 * style/variant/weight/stretch tokens before it never carry a unit. */
const FONT_SIZE = /(^|\s)((?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?)px(\/\S+)?(?=\s)/i;

interface ParsedFontShorthand {
  readonly before: string;
  readonly sizePx: number;
  readonly lineHeight: string;
  readonly family: string;
}

function parseFontShorthand(font: string): ParsedFontShorthand | null {
  const match = FONT_SIZE.exec(font);
  if (!match) return null;
  const sizePx = Number(match[2]);
  if (!Number.isFinite(sizePx) || sizePx <= 0) return null;
  const family = font.slice(match.index + match[0].length).trim();
  if (!family) return null;
  return {
    before: font.slice(0, match.index + match[1]!.length).trim(),
    sizePx,
    lineHeight: match[3] ?? '',
    family,
  };
}

function shorthandAt(parsed: ParsedFontShorthand, sizePx: number): string {
  return `${parsed.before ? `${parsed.before} ` : ''}${sizePx}px${parsed.lineHeight} ${parsed.family}`;
}

/** `font` at the reference em, or `null` when it has no absolute px size
 * (keywords and relative units already scale with their context). */
export function referenceTextFont(font: string): ReferenceTextFont | null {
  const parsed = parseFontShorthand(font);
  if (!parsed) return null;
  return Object.freeze({
    font: shorthandAt(parsed, TEXT_ADVANCE_REFERENCE_PX),
    scale: parsed.sizePx / TEXT_ADVANCE_REFERENCE_PX,
  });
}

/** Font identity for planning: equal for shorthands that select the same
 * face at the same size regardless of keyword spelling (`bold` ≡ `700`,
 * `normal` ≡ absent) or token order before the size. */
function fontPlanKey(parsed: ParsedFontShorthand): string {
  const tokens = parsed.before.toLowerCase().split(/\s+/u)
    .filter((token) => token !== '' && token !== 'normal' && token !== '400')
    .map((token) => (token === 'bold' ? '700' : token))
    .sort();
  return `${tokens.join(' ')}|${parsed.sizePx}|${parsed.lineHeight}|${parsed.family}`;
}

/** Per-cluster deviation from the proportional advance that still counts as
 * proportional, in two parts. An absolute part far above fixed-point glyph
 * positioning (2^-16 px). A relative part for an engine that resolves the
 * requested size to finite precision, which scales every advance uniformly:
 * Blink keys faces by size in 1/100 px, under 2^-9 of any advance from 5.12px
 * up, so fractional sizes such as a 6.825px superscript stay at their own size.
 * Whole-pixel grid fitting moves at least one probe cluster by several
 * percent. */
const POSITION_ALLOWANCE_PX = 2 ** -10;
const SIZE_PRECISION_ALLOWANCE = 2 ** -9;
/** Distance from a whole pixel within which a measured advance counts as
 * grid-fitted: grid-fitting hosts report exact integers, while an advance
 * from a design (any optical-size instance included) is a fixed-point value
 * that lands this close to an integer only by coincidence. */
const WHOLE_PIXEL_TOLERANCE_PX = 2 ** -10;
/** Relative agreement required between the reference em and twice it before
 * a context is trusted to scale metrics with the font size at all. */
const SIZE_PROPORTIONAL_TOLERANCE = 2 ** -12;

/** The measuring surface a planner probes: a Canvas 2D context or the
 * structural subset of one that text layout uses. */
export interface TextAdvanceMeasurer {
  font: string;
  letterSpacing: string;
  fontKerning?: CanvasFontKerning;
  measureText(text: string): Pick<TextMetrics, 'width'>;
}

/** Read-only view of planned shorthands, for consumers that must not measure
 * (retained paint): `undefined` means the shorthand was never planned. */
export interface TextAdvanceLookup {
  planned(font: string): ReferenceTextFont | null | undefined;
}

export interface TextAdvancePlanner extends TextAdvanceLookup {
  /** Whether `font` is measured and painted at the reference em (its
   * {@link ReferenceTextFont}) or at its own size (`null`). Probes the
   * measuring context once per distinct shorthand; the answer is fixed for
   * the planner's lifetime so every consumer of one layout agrees. */
  plan(font: string): ReferenceTextFont | null;
}

export interface TextAdvancePlannerOptions {
  /** Probe clusters, measured one at a time so opposite rounding errors
   * cannot cancel. Grid fitting is a per-face, per-size measuring mode, so a
   * few clusters of differing proportional widths that each face of a route
   * draws (Latin letters and a digit; a kana and an ideograph for East Asian
   * faces) expose it: a grid-fitting host moves at least one of them to a
   * whole pixel. They must be scalars the format already measures and
   * demands, so probing never touches a font subset the document does not
   * load. */
  readonly probeClusters: readonly string[];
}

/** Creates the planner for one measuring context. A context whose metrics do
 * not scale with the font size at large sizes (a structural test double) is
 * never re-sized: its measurements are used as given. */
export function createTextAdvancePlanner(
  context: TextAdvanceMeasurer,
  options: TextAdvancePlannerOptions,
): TextAdvancePlanner {
  const probeClusters = Object.freeze(options.probeClusters.filter((cluster) => cluster !== ''));
  if (probeClusters.length === 0) throw new RangeError('Text advance planning requires probe clusters');
  const probeText = probeClusters.join('');
  /** Decisions by face and size; the reference shorthand returned for a
   * spelling is always that spelling at the reference em. */
  const byKey = new Map<string, boolean>();
  const byFont = new Map<string, ReferenceTextFont | null>();
  const probe = (parsed: ParsedFontShorthand, reference: ReferenceTextFont): boolean => {
    const previousFont = context.font;
    const previousSpacing = context.letterSpacing;
    const previousKerning = context.fontKerning;
    try {
      if (typeof previousSpacing === 'string') context.letterSpacing = '0px';
      if (previousKerning !== undefined) context.fontKerning = 'none';
      context.font = reference.font;
      const referenceWidths = probeClusters.map((cluster) => context.measureText(cluster).width);
      const referenceTotal = context.measureText(probeText).width;
      context.font = shorthandAt(parsed, TEXT_ADVANCE_REFERENCE_PX * 2);
      const doubleTotal = context.measureText(probeText).width;
      if (!(Number.isFinite(referenceTotal) && referenceTotal > 0 && Number.isFinite(doubleTotal))
        || Math.abs(doubleTotal - 2 * referenceTotal) > 2 * referenceTotal * SIZE_PROPORTIONAL_TOLERANCE) {
        return false;
      }
      context.font = shorthandAt(parsed, parsed.sizePx);
      // Grid fitting is proven only when clusters depart from their
      // proportional advance and every departing cluster measures a whole
      // number of pixels. One fractional departure means a size-dependent
      // design (opsz, tracking) shapes part of this route, which the
      // reference em would replace with its display design: keep the size.
      let departed = false;
      for (const [index, cluster] of probeClusters.entries()) {
        const width = context.measureText(cluster).width;
        const linear = referenceWidths[index]! * reference.scale;
        if (Math.abs(width - linear) <= POSITION_ALLOWANCE_PX + Math.abs(linear) * SIZE_PRECISION_ALLOWANCE) {
          continue;
        }
        if (!(Math.abs(width - Math.round(width)) <= WHOLE_PIXEL_TOLERANCE_PX)) return false;
        departed = true;
      }
      return departed;
    } finally {
      context.font = previousFont;
      if (typeof previousSpacing === 'string') context.letterSpacing = previousSpacing;
      if (previousKerning !== undefined) context.fontKerning = previousKerning;
    }
  };
  const resolve = (font: string, probeMissing: boolean): ReferenceTextFont | null | undefined => {
    const known = byFont.get(font);
    if (known !== undefined || byFont.has(font)) return known;
    const parsed = parseFontShorthand(font);
    const reference = parsed && parsed.sizePx !== TEXT_ADVANCE_REFERENCE_PX ? referenceTextFont(font) : null;
    if (!parsed || !reference) {
      byFont.set(font, null);
      return null;
    }
    const key = fontPlanKey(parsed);
    let planned = byKey.get(key);
    if (planned === undefined) {
      if (!probeMissing) return undefined;
      planned = probe(parsed, reference);
      byKey.set(key, planned);
    }
    const result = planned ? reference : null;
    byFont.set(font, result);
    return result;
  };
  return Object.freeze({
    planned: (font: string) => resolve(font, false),
    plan: (font: string) => resolve(font, true) ?? null,
  });
}

const SCALED_TEXT_METRICS = [
  'width',
  'actualBoundingBoxLeft',
  'actualBoundingBoxRight',
  'actualBoundingBoxAscent',
  'actualBoundingBoxDescent',
  'fontBoundingBoxAscent',
  'fontBoundingBoxDescent',
  'emHeightAscent',
  'emHeightDescent',
  'hangingBaseline',
  'alphabeticBaseline',
  'ideographicBaseline',
] as const;

/** Reference-em metrics mapped onto the requested size. Every TextMetrics
 * length is proportional to the em; metrics the engine omits stay absent. */
export function scaleTextMetrics(metrics: TextMetrics, scale: number): TextMetrics {
  const scaled: Partial<Record<(typeof SCALED_TEXT_METRICS)[number], number>> = {};
  for (const key of SCALED_TEXT_METRICS) {
    const value = metrics[key];
    if (typeof value === 'number') scaled[key] = value * scale;
  }
  return Object.freeze(scaled) as TextMetrics;
}

function scaledLetterSpacing(spacing: string, factor: number): string {
  if (typeof spacing !== 'string' || factor === 1) return spacing;
  const match = /^\s*(-?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?)px\s*$/i.exec(spacing);
  if (!match) return spacing;
  const px = Number(match[1]);
  return px === 0 ? spacing : `${px * factor}px`;
}

/** The text state of a Canvas 2D context a reference-advance view forwards. */
export interface ReferenceAdvanceTarget extends TextAdvanceMeasurer {
  fontKerning: CanvasFontKerning;
  textAlign: CanvasTextAlign;
  textBaseline: CanvasTextBaseline;
  direction: CanvasDirection;
  measureText(text: string): TextMetrics;
  readonly canvas?: unknown;
}

/** A measuring view of `target` whose `font` and `letterSpacing` are the
 * logical (requested) values: a planned shorthand is applied to the target at
 * the reference em with letter spacing in reference px, and measurements come
 * back in requested px. One target measurement per `measureText`. */
export interface ReferenceAdvanceContext {
  font: string;
  letterSpacing: string;
  fontKerning: CanvasFontKerning;
  textAlign: CanvasTextAlign;
  textBaseline: CanvasTextBaseline;
  direction: CanvasDirection;
  readonly canvas: unknown;
  measureText(text: string): TextMetrics;
}

export function createReferenceAdvanceContext(
  target: ReferenceAdvanceTarget,
  planner: TextAdvancePlanner,
): ReferenceAdvanceContext {
  let font = target.font;
  let letterSpacing = target.letterSpacing;
  /** Requested px per target px of the shorthand applied to the target. A
   * shorthand measured at its own size leaves the target's spacing as set. */
  let appliedScale = 1;
  return {
    get font() { return font; },
    set font(value: string) {
      font = value;
      const reference = planner.plan(value);
      target.font = reference ? reference.font : value;
      const scale = reference ? reference.scale : 1;
      if (scale !== appliedScale) {
        appliedScale = scale;
        target.letterSpacing = scaledLetterSpacing(letterSpacing, 1 / scale);
      }
    },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(value: string) {
      letterSpacing = value;
      target.letterSpacing = appliedScale === 1 ? value : scaledLetterSpacing(value, 1 / appliedScale);
    },
    get fontKerning() { return target.fontKerning; },
    set fontKerning(value: CanvasFontKerning) { target.fontKerning = value; },
    get textAlign() { return target.textAlign; },
    set textAlign(value: CanvasTextAlign) { target.textAlign = value; },
    get textBaseline() { return target.textBaseline; },
    set textBaseline(value: CanvasTextBaseline) { target.textBaseline = value; },
    get direction() { return target.direction; },
    set direction(value: CanvasDirection) { target.direction = value; },
    get canvas() { return target.canvas; },
    measureText(text: string): TextMetrics {
      const metrics = target.measureText(text);
      return appliedScale === 1 ? metrics : scaleTextMetrics(metrics, appliedScale);
    },
  };
}

/** `target` itself with its `font`, `letterSpacing` and `measureText` served
 * by `view` (whose target it is): APIs typed for a concrete Canvas context
 * measure through the reference-advance view while every other member, the
 * `canvas` element included, stays the target's own. */
export function withReferenceAdvances<T extends ReferenceAdvanceTarget>(
  target: T,
  view: ReferenceAdvanceContext,
): T {
  const measureText = (text: string): TextMetrics => view.measureText(text);
  return new Proxy(target, {
    get(object, key) {
      if (key === 'font') return view.font;
      if (key === 'letterSpacing') return view.letterSpacing;
      if (key === 'measureText') return measureText;
      const value: unknown = Reflect.get(object, key, object);
      return typeof value === 'function' ? value.bind(object) : value;
    },
    set(object, key, value) {
      if (key === 'font') view.font = value as string;
      else if (key === 'letterSpacing') view.letterSpacing = value as string;
      else return Reflect.set(object, key, value, object);
      return true;
    },
  });
}

/** The drawing subset {@link fillTextAtReference} uses. */
export interface ReferenceTextPaintTarget {
  font: string;
  letterSpacing: string;
  save(): void;
  restore(): void;
  translate(x: number, y: number): void;
  scale(x: number, y: number): void;
  fillText(text: string, x: number, y: number): void;
}

/** Draws `text` with its anchor at (`x`, `y`) in the current frame using the
 * reference em under a linear scale, so the shaped advances are the
 * reference-em advances a planned layout measured. The caller's current
 * `letterSpacing` (requested px) is converted to reference px; alignment,
 * baseline, direction, kerning and fill state are scale invariant. The font
 * and spacing are restored with the drawing state. */
export function fillTextAtReference(
  ctx: ReferenceTextPaintTarget,
  reference: ReferenceTextFont,
  text: string,
  x: number,
  y: number,
): void {
  const spacing = ctx.letterSpacing;
  saveCanvasState(ctx);
  try {
    ctx.translate(x, y);
    ctx.scale(reference.scale, reference.scale);
    ctx.font = reference.font;
    ctx.letterSpacing = scaledLetterSpacing(spacing, 1 / reference.scale);
    ctx.fillText(text, 0, 0);
  } finally {
    ctx.restore();
  }
}

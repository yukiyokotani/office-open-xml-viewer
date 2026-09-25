export interface CanvasFontBoxProbeContext {
  font: string;
  measureText(text: string): TextMetrics;
}

export interface CanvasFontBoxProbeOptions {
  /** A glyph that the requested face must itself cover. */
  readonly text: string;
  /** Probe size. A large em reduces integer rounding in Canvas metrics. */
  readonly emPx?: number;
  readonly weight?: number;
  readonly style?: 'normal' | 'italic';
}

const MISSING_CONTROL_FAMILY = '__ooxml_missing_font_control_6f33c9b4__';

function quoteCssFamily(family: string): string {
  let escaped = '';
  for (const scalar of family) {
    const codePoint = scalar.codePointAt(0) ?? 0;
    if (scalar === '"' || scalar === '\\' || codePoint <= 0x1f || codePoint === 0x7f) {
      escaped += `\\${codePoint.toString(16)} `;
    } else {
      escaped += scalar;
    }
  }
  return `"${escaped}"`;
}

function fontString(
  family: string,
  emPx: number,
  weight: number,
  style: 'normal' | 'italic',
): string {
  // Use the same explicit control + generic tail for both measurements. If a
  // hostile or coincidental global FontFace owns the control name, a missing
  // requested family still resolves through exactly the same chain and cannot
  // be misidentified as present.
  return `${style} ${weight} ${emPx}px ${quoteCssFamily(family)}, ${quoteCssFamily(MISSING_CONTROL_FAMILY)}, serif`;
}

function glyphGeometryTuple(metrics: TextMetrics): readonly number[] {
  return [
    metrics.width,
    metrics.actualBoundingBoxAscent,
    metrics.actualBoundingBoxDescent,
  ];
}

/**
 * Measure a browser-selected family only when it demonstrably supplies the
 * requested glyph. CSS silently substitutes a missing family, so the same
 * glyph is also measured through a deliberately absent control family. Equal
 * glyph geometry means the selected face cannot be distinguished from fallback
 * and no metric is claimed. The font bounding box is deliberately excluded
 * from that identity check: browsers may report the primary Latin face's box
 * even when the requested glyph came from an East Asian fallback.
 *
 * This discovers no alternate name and chooses no substitute. It reports the
 * concrete Canvas face selected for the authored family on this browser.
 */
export function measureResolvedCanvasFontBoxRatio(
  context: CanvasFontBoxProbeContext,
  familyValue: string,
  options: CanvasFontBoxProbeOptions,
): number | null {
  const family = familyValue.trim();
  const text = options.text;
  const emPx = options.emPx ?? 100;
  const weight = options.weight ?? 400;
  const style = options.style ?? 'normal';
  if (
    !family
    || family === MISSING_CONTROL_FAMILY
    || !text
    || !Number.isFinite(emPx)
    || emPx <= 0
    || !Number.isInteger(weight)
    || weight < 1
    || weight > 1000
  ) return null;

  const previousFont = context.font;
  try {
    context.font = fontString(family, emPx, weight, style);
    const selected = context.measureText(text);
    context.font = fontString(MISSING_CONTROL_FAMILY, emPx, weight, style);
    const missing = context.measureText(text);
    const selectedTuple = glyphGeometryTuple(selected);
    const missingTuple = glyphGeometryTuple(missing);
    if (selectedTuple.every((value, index) => Object.is(value, missingTuple[index]))) {
      return null;
    }
    const ascent = selected.fontBoundingBoxAscent;
    const descent = selected.fontBoundingBoxDescent;
    if (!(Number.isFinite(ascent) && Number.isFinite(descent) && ascent + descent > 0)) {
      return null;
    }
    return (ascent + descent) / emPx;
  } finally {
    context.font = previousFont;
  }
}

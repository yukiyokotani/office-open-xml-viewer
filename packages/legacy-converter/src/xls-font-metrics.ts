/** The actual BIFF Normal-style font, resolved through its style XF. */
export interface LegacyXlsNormalFont {
  readonly family: string;
  readonly sizePoints: number;
  readonly bold: boolean;
  readonly italic: boolean;
}

/**
 * Measure digits 0–9 in this font at 96 dpi, then return the rounded maximum
 * advance in pixels (integer 1–4096). Load the intended font before measuring;
 * return undefined if unavailable. No fallback font width is assumed.
 * The signal is aborted if conversion is cancelled. Never fetch a font URL
 * supplied by the document: family is an untrusted name, not a resource URL.
 */
export type LegacyXlsFontMeasurement = (
  font: Readonly<LegacyXlsNormalFont>,
  signal: AbortSignal,
) => number | undefined | Promise<number | undefined>;

/** The callback cannot keep a prepared WASM model alive after cancellation. */
export function measureXlsFont(
  measure: LegacyXlsFontMeasurement,
  font: Readonly<LegacyXlsNormalFont>,
  signal: AbortSignal,
): Promise<number | undefined> {
  return new Promise((resolve, reject) => {
    const aborted = () => reject(new Error('XLS font measurement aborted'));
    if (signal.aborted) { aborted(); return; }
    signal.addEventListener('abort', aborted, { once: true });
    Promise.resolve().then(() => {
      if (signal.aborted) throw new Error('XLS font measurement aborted');
      return measure(Object.freeze(font), signal);
    }).then((width) => {
      if (width !== undefined && (!Number.isInteger(width) || width < 1 || width > 4096)) {
        throw new Error('invalid XLS maximum digit width');
      }
      resolve(width);
    }).catch(reject).finally(() => signal.removeEventListener('abort', aborted));
  });
}

/**
 * Default browser measurement for the direct XLS source. Excel column widths
 * are expressed in the Normal font's maximum digit width in whole pixels
 * (ECMA-376 §18.3.1.13), so this loads the named font through the document's
 * FontFaceSet and measures digits 0–9 at 96 dpi. It honors the measurement
 * contract: when the named font cannot be loaded it returns undefined rather
 * than measuring a fallback face, and callers then omit geometry-dependent
 * drawings. Once shared reference font metrics can supply Office font
 * advances without an installed face, this default should use them instead.
 */
export async function measureLegacyXlsNormalFontInDocument(
  font: Readonly<LegacyXlsNormalFont>,
  signal: AbortSignal,
): Promise<number | undefined> {
  if (typeof document === 'undefined' || !document.fonts) return undefined;
  const px = font.sizePoints * 96 / 72;
  const family = `"${font.family.replace(/["\\]/g, '')}"`;
  const spec = `${font.italic ? 'italic ' : ''}${font.bold ? 'bold ' : ''}${px}px ${family}`;
  try {
    await document.fonts.load(spec, '0123456789');
  } catch {
    return undefined;
  }
  if (signal.aborted || !document.fonts.check(spec, '0123456789')) return undefined;
  const context = document.createElement('canvas').getContext('2d');
  if (!context) return undefined;
  context.font = spec;
  let widest = 0;
  for (const digit of '0123456789') widest = Math.max(widest, context.measureText(digit).width);
  const width = Math.round(widest);
  return width >= 1 && width <= 4096 ? width : undefined;
}

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

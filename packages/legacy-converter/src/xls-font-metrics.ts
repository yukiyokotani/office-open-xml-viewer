import { fontStackFor } from '@silurus/ooxml-core/internal/spreadsheet-font-stack';
export type {
  LegacyXlsFontMeasurement,
  LegacyXlsNormalFont,
} from '@silurus/ooxml-core/internal/legacy-xls-source';
import type {
  LegacyXlsFontMeasurement,
  LegacyXlsNormalFont,
} from '@silurus/ooxml-core/internal/legacy-xls-source';

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
 * (ECMA-376 §18.3.1.13). This loads the named face through the document's
 * FontFaceSet, then measures digits 0-9 at 96 dpi through the same cell font
 * stack the spreadsheet renderer paints with. When the authored face is
 * unavailable the painted fallback face is measured, so column geometry,
 * drawing anchors and painted text agree, as for an XLSX workbook whose
 * Normal font is missing. Returns undefined only without a DOM or canvas.
 */
export async function measureLegacyXlsNormalFontInDocument(
  font: Readonly<LegacyXlsNormalFont>,
  signal: AbortSignal,
): Promise<number | undefined> {
  if (typeof document === 'undefined') return undefined;
  const px = font.sizePoints * 96 / 72;
  const style = `${font.italic ? 'italic ' : ''}${font.bold ? 'bold ' : ''}${px}px`;
  const named = `${style} "${font.family.replace(/["\\]/g, '')}"`;
  try {
    await document.fonts?.load(named, '0123456789');
  } catch {
    // An unavailable authored face is measured through the painted fallback.
  }
  if (signal.aborted) return undefined;
  const context = document.createElement('canvas').getContext('2d');
  if (!context) return undefined;
  context.font = `${style} ${fontStackFor(font.family)}`;
  let widest = 0;
  for (const digit of '0123456789') widest = Math.max(widest, context.measureText(digit).width);
  const width = Math.round(widest);
  return width >= 1 && width <= 4096 ? width : undefined;
}

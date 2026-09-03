/** Shared first-paint helpers for embedded and provider-backed fonts. */

const HARD_CEILING_MS = 15000;

/** Bound a font operation so a wedged network or browser font loader cannot hang. */
export function withFontCeiling<T>(operation: Promise<T>): Promise<T | void> {
  return Promise.race([
    operation,
    new Promise<void>((resolve) => setTimeout(resolve, HARD_CEILING_MS)),
  ]);
}

/** FontFaceSet owned by the current document or render worker. */
export function activeFontSet(): FontFaceSet | null {
  if (typeof document !== 'undefined' && document?.fonts) return document.fonts;
  if (typeof self !== 'undefined' && self && 'fonts' in self) {
    return Reflect.get(self, 'fonts') as FontFaceSet;
  }
  return null;
}

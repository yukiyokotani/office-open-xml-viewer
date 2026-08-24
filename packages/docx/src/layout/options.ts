import type { LayoutServices } from './types.js';
import { stableFingerprint } from './fingerprint.js';

export interface LayoutOptions {
  readonly currentDateMs: number;
}

export interface LayoutRenderSelectionInput {
  readonly currentDate?: Date | number;
  readonly defaultCurrentDateMs: number;
}

export function normalizeLayoutOptions(
  currentDate: Date | number | undefined,
  defaultCurrentDateMs: number,
): LayoutOptions {
  const currentDateMs = currentDate == null
    ? defaultCurrentDateMs
    : typeof currentDate === 'number' ? currentDate : currentDate.getTime();
  if (!Number.isFinite(currentDateMs)) throw new RangeError('currentDate must resolve to finite epoch milliseconds');
  return Object.freeze({ currentDateMs });
}

export function layoutOptionsForRender(input: LayoutRenderSelectionInput): LayoutOptions {
  return normalizeLayoutOptions(input.currentDate, input.defaultCurrentDateMs);
}

export function layoutOptionsKey(options: LayoutOptions, services: LayoutServices): string {
  return stableFingerprint('layout', {
    currentDateMs: options.currentDateMs,
    text: services.text.fingerprint,
    images: services.images.fingerprint,
    math: services.math.fingerprint,
    verticalGlyphs: services.verticalGlyphFingerprint ?? null,
  });
}

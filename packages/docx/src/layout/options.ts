import type { LayoutServices } from './types.js';
import { stableFingerprint } from './fingerprint.js';

export interface LayoutOptions {
  readonly currentDateMs: number;
  /** ECMA-376 §17.13.5 tracked-change view. `false` (the default) lays the
   * document out in its final state: deleted (`w:del`) and moved-away
   * (`w:moveFrom`) runs are hidden. `true` selects the markup view: revision
   * content stays visible and receives author-coloured decoration. A
   * geometry-selecting acquisition input — hiding deletions changes line
   * breaking and pagination — so it participates in the variant key. Optional
   * so a bare `{ currentDateMs }` stays a valid literal; absent means the
   * default final view. */
  readonly showTrackedChanges?: boolean;
}

export interface LayoutRenderSelectionInput {
  readonly currentDate?: Date | number;
  readonly defaultCurrentDateMs: number;
  readonly showTrackedChanges?: boolean;
}

export function normalizeLayoutOptions(
  currentDate: Date | number | undefined,
  defaultCurrentDateMs: number,
  showTrackedChanges = false,
): LayoutOptions {
  const currentDateMs = currentDate == null
    ? defaultCurrentDateMs
    : typeof currentDate === 'number' ? currentDate : currentDate.getTime();
  if (!Number.isFinite(currentDateMs)) throw new RangeError('currentDate must resolve to finite epoch milliseconds');
  // The final-view default omits the key entirely so normalized default
  // options keep their historical `{ currentDateMs }` shape (and the default
  // variant's options object stays deep-equal to pre-axis builds).
  return Object.freeze({
    currentDateMs,
    ...(showTrackedChanges === true ? { showTrackedChanges: true as const } : {}),
  });
}

export function layoutOptionsForRender(input: LayoutRenderSelectionInput): LayoutOptions {
  return normalizeLayoutOptions(
    input.currentDate,
    input.defaultCurrentDateMs,
    input.showTrackedChanges,
  );
}

export function layoutOptionsKey(options: LayoutOptions, services: LayoutServices): string {
  return stableFingerprint('layout', {
    currentDateMs: options.currentDateMs,
    showTrackedChanges: options.showTrackedChanges === true,
    text: services.text.fingerprint,
    images: services.images.fingerprint,
    math: services.math.fingerprint,
    verticalGlyphs: services.verticalGlyphFingerprint ?? null,
  });
}

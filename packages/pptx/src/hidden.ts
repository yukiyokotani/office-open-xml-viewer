import type { DimOptions } from './types';

// Generic index-navigation helpers live in core (shared with xlsx hidden sheets).
export { nextVisibleIndex, resolveVisibleIndex, countVisible } from '@silurus/ooxml-core';

/** How `PptxViewer` and `PptxScrollViewer` present hidden slides
 *  (`<p:sld show="0">`). */
export type HiddenSlideMode = 'show' | 'skip' | 'dim';

/** Default `'dim'` overlay: 60% white (hidden content shows at 40%). */
const DEFAULT_HIDDEN_DIM: DimOptions = { color: '#ffffff', opacity: 0.6 };

/** The `'dim'` overlay with a viewer's `hiddenSlideDim` overrides merged over
 *  the default, so both viewers dim a hidden slide identically. */
export function resolveHiddenSlideDim(overrides: Partial<DimOptions> | undefined): DimOptions {
  return {
    color: overrides?.color ?? DEFAULT_HIDDEN_DIM.color,
    opacity: overrides?.opacity ?? DEFAULT_HIDDEN_DIM.opacity,
  };
}

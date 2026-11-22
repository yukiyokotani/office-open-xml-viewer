import type { LayoutServices } from './types.js';
import { stableFingerprint } from './fingerprint.js';

/** Geometry-selecting render inputs, normalized before layout. Only inputs
 * that can change layout belong here: paint-only scale, DPR, target width and
 * colour never enter a layout variant key. Service identity enters the key
 * only through each injected service's own fingerprint; there is no separate
 * caller-supplied fingerprint. */
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

/** Keys one geometry variant among the layouts of a single variant store. */
export type LayoutOptionsKeyer = (options: LayoutOptions, services: LayoutServices) => string;

/**
 * Create a variant keyer whose keys are exact within its owner.
 *
 * A service fingerprint is an exact canonical identity, so it grows with the
 * data it identifies: the text-service fingerprint embeds the document's
 * complete font-metric snapshot (tens of KB). Spelling it into the key
 * re-encoded that whole string on every variant lookup, which runs on every
 * page render. Each distinct fingerprint string instead gets an ordinal, the
 * same interning the paragraph acquisition keys use: the map keeps the string
 * itself (no copy) and assigns ordinals one-to-one, so two keys from one keyer
 * are equal exactly when the date, tracked-change view and every service
 * fingerprint are equal. Keys from different keyers are not comparable; a
 * variant store owns one keyer for its fixed services.
 */
export function createLayoutOptionsKeyer(): LayoutOptionsKeyer {
  const fingerprintOrdinals = new Map<string, number>();
  const ordinal = (fingerprint: string | undefined): number | null => {
    if (fingerprint === undefined) return null;
    let value = fingerprintOrdinals.get(fingerprint);
    if (value === undefined) {
      value = fingerprintOrdinals.size;
      fingerprintOrdinals.set(fingerprint, value);
    }
    return value;
  };
  return (options, services) => stableFingerprint('layout', {
    currentDateMs: options.currentDateMs,
    showTrackedChanges: options.showTrackedChanges === true,
    text: ordinal(services.text.fingerprint),
    images: ordinal(services.images.fingerprint),
    math: ordinal(services.math.fingerprint),
    verticalGlyphs: ordinal(services.verticalGlyphFingerprint),
  });
}

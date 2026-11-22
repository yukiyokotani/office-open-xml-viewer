import {
  resolveOpcPartName,
  parseRelativeSlideJump,
  resolveRelativeSlideJump,
} from '@silurus/ooxml-core';

/**
 * PPTX internal-navigation resolution (IX-nav).
 *
 * An internal hyperlink in a deck jumps to another slide rather than opening a
 * URL (ECMA-376 §21.1.2.3.5, `<a:hlinkClick @action>`):
 *   - `ppaction://hlinksldjump` names a specific slide via its `r:id` — the rel
 *     Target (e.g. `../slides/slide3.xml`) resolves to a slide *part name*.
 *   - `ppaction://hlinkshowjump?jump=firstslide|lastslide|nextslide|previousslide`
 *     names a slide *relative* to the current one, with no rel.
 *
 * This module turns either into a 0-based slide index. The part-name path uses a
 * `partName → index` map built from the parsed slides (whose `partName` the
 * parser stamped in `sldIdLst` order); the relative path uses pure arithmetic
 * from core. The map + these resolvers are the piece the viewers were missing to
 * make an internal slide-jump click actually navigate.
 */

/** The slide `partName` per index (`sldIdLst` order), from either the parsed
 *  model (main mode) or the worker meta (worker mode). Index i's entry is that
 *  slide's normalized OPC part name, or `undefined` when unrecorded. */
export type SlidePartNames = readonly (string | undefined)[];

/**
 * Build a `partName → 0-based index` map from the per-slide part names. When two
 * slides somehow share a part name (malformed deck), the FIRST wins — an
 * internal jump lands on the earliest matching slide, matching the deterministic
 * `sldIdLst` scan order.
 */
export function buildSlidePartIndex(partNames: SlidePartNames): Map<string, number> {
  const map = new Map<string, number>();
  for (let i = 0; i < partNames.length; i++) {
    const name = partNames[i];
    if (name !== undefined && name !== '' && !map.has(name)) map.set(name, i);
  }
  return map;
}

/**
 * Resolve a `ppaction://hlinksldjump` slide-part target to a 0-based slide index.
 *
 * `target` is the run/shape's resolved hyperlink string — for a slide jump this
 * is the rel Target verbatim (e.g. `../slides/slide3.xml`), authored relative to
 * the slide's own directory `ppt/slides`. It is normalized to a part name
 * (`ppt/slides/slide3.xml`) via the shared OPC resolver and looked up in
 * `partIndex`. Returns `undefined` when the target doesn't resolve to a known
 * slide (e.g. it's actually an external URL, or names a non-slide part).
 *
 * @param target    the raw hyperlink target string from the run/shape.
 * @param partIndex the map from {@link buildSlidePartIndex}.
 */
export function resolveSlidePartTarget(
  target: string,
  partIndex: Map<string, number>,
): number | undefined {
  if (target === '') return undefined;
  // Master/layout-inherited hyperlink relationships are resolved relative to
  // their owning OPC part by the parser. Their canonical part name is already
  // rooted at ppt/, whereas a slide-local target is still relative to slides.
  if (target.startsWith('ppt/')) return partIndex.get(target);
  // Slide-rel targets are authored relative to the slide part's directory.
  const partName = resolveOpcPartName('ppt/slides', target);
  return partIndex.get(partName);
}

/**
 * Resolve any pptx internal hyperlink to a 0-based slide index, or `undefined`
 * when it names no reachable slide. Handles both action classes:
 *
 *   - a relative show jump (`ppaction://hlinkshowjump?jump=…`) → arithmetic from
 *     `current` (needs `current`; clamps at the deck ends);
 *   - a specific slide-part jump (`hlinksldjump`, or any target that resolves to
 *     a known slide part) → the `partIndex` lookup.
 *
 * `ref` is the internal reference the viewer holds: either the raw
 * `ppaction://…` action string OR the resolved slide-part target string. Passing
 * the action string lets relative jumps resolve; passing the part target lets
 * specific jumps resolve; a caller that has both should prefer the action string
 * (a relative action has no part target).
 *
 * @param ref       the internal action/target string.
 * @param partIndex the map from {@link buildSlidePartIndex}.
 * @param current   the 0-based index the jump is relative to (for show jumps).
 */
export function resolveInternalSlideTarget(
  ref: string,
  partIndex: Map<string, number>,
  current: number,
): number | undefined {
  // A relative show-jump verb resolves without a part lookup.
  const jump = parseRelativeSlideJump(ref);
  if (jump !== null) {
    return resolveRelativeSlideJump(jump, current, partIndex.size);
  }
  // Otherwise treat `ref` as a slide-part target.
  return resolveSlidePartTarget(ref, partIndex);
}

/**
 * Whether an internal hyperlink `ref` is the relative `previousslide` show
 * jump. In `hiddenSlideMode: 'skip'` its hidden target gives way to the nearest
 * shown slide before it (as `prevSlide()` does), not the next one after it,
 * which could be the slide the jump started from.
 */
export function isBackwardSlideJump(ref: string): boolean {
  return parseRelativeSlideJump(ref) === 'previousslide';
}

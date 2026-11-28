import type { CanvasFontBoxProbeContext } from './canvas-font-box.js';
import { quoteCssFontFamily } from '../internal/css-font-family.js';
import { getGraphemeSegmenter } from '../text/sea-break.js';
import { scratchCanvas, type SymbolProbeCanvasFactory } from './symbol-face.js';
import {
  PROBE_EM, SENTINELS, SENTINEL_FAMILIES, SENTINEL_UPM, parseFamilyList, sentinelsWork, type FamilyEntry,
} from './route-sentinels.js';

export { loadNativeRouteSentinels } from './route-sentinels.js';

/** One CSS family list exactly as a format's resolver hands it to Canvas, with
 * the requested weight/style tuple. */
export interface NativeRouteLineRequest {
  readonly familyList: string;
  readonly weight: number;
  readonly style: 'normal' | 'italic';
}

/** The normal line of a native route's first available face, measured in the
 * DOM realm. Plain data: it crosses the worker boundary and the layout realm
 * re-verifies it ({@link verifyNativeRouteNormalLine}) before admitting it. */
export interface NativeRouteNormalLine extends NativeRouteLineRequest {
  /** Index in the family list of the first available family. */
  readonly primaryIndex: number;
  /** That family's name (unquoted). */
  readonly primaryFamily: string;
  /** That family as written in the list (CSS syntax). */
  readonly primaryFamilyText: string;
  /** Probe em whose measurement scaled linearly. */
  readonly emPx: number;
  readonly ascentRatio: number;
  readonly descentRatio: number;
  /** CSS normal line minus the Canvas font box; signed. */
  readonly lineGapRatio: number;
}

/** Format policy: whether an identified first available family may be
 * measured and admitted. */
export type NativeRoutePrimaryAdmission = (
  primaryFamily: string, weight: number, style: 'normal' | 'italic',
) => boolean;

// The route sentinels (route-sentinels.ts) make observable which entry of a
// CSS family list Canvas uses as the first available face, and which faces
// of the list draw a cluster.

/** The strut starts at twice the probe size, where engines round the normal
 * line of a 2048-unit face to whole design units; an engine that clamps that
 * size (Firefox, near 2000 px) halves it through the linear-scale check. */
const STRUT_EM = 2 * PROBE_EM;

function newProbeContext(owner?: Document | null): CanvasFontBoxProbeContext | null {
  if (owner) return owner.createElement('canvas').getContext('2d');
  return typeof OffscreenCanvas !== 'undefined'
    ? new OffscreenCanvas(1, 1).getContext('2d')
    : typeof document !== 'undefined' ? document.createElement('canvas').getContext('2d') : null;
}

/** The sentinel geometry is a design contract: verify that this realm loaded
 * and applies both registrations before any result depends on them. */
function sameFamilyList(a: readonly FamilyEntry[], b: readonly FamilyEntry[]): boolean {
  return a.length === b.length && a.every((entry, index) =>
    entry.name === b[index]!.name && entry.generic === b[index]!.generic);
}

function registeredFamilies(fontSet: FontFaceSet | null | undefined): Set<string> {
  const names = new Set<string>();
  if (!fontSet || typeof fontSet[Symbol.iterator] !== 'function') return names;
  for (const face of fontSet) {
    const family = parseFamilyList(face.family)?.[0]?.name ?? face.family;
    names.add(family.trim().toLowerCase());
  }
  return names;
}

/** Index of the first available family, judged by the Canvas font box (see
 * the sentinel notes above). The first available face of a list is that of
 * its first family that resolves to any face. A system family resolves for
 * every weight and style once it resolves at all (CSS Fonts 4 §5.2 matches
 * the nearest face and synthesizes the rest), and registered faces are
 * declined separately, so availability is judged once per family at the
 * regular tuple and memoized: routes share most of their families. */
function primaryIndex(
  context: CanvasFontBoxProbeContext,
  entries: readonly FamilyEntry[],
  memo: Map<string, boolean>,
): number | undefined {
  for (let index = 0; index < entries.length; index += 1) {
    const entry = entries[index]!;
    // A generic keyword always resolves; callers decline it unprobed.
    if (entry.generic) return index;
    const key = entry.text;
    let available = memo.get(key);
    if (available === undefined) {
      const boxes = SENTINEL_FAMILIES.map((family) => {
        context.font = `${PROBE_EM}px ${entry.text}, ${quoteCssFontFamily(family)}`;
        const box = context.measureText('');
        return [box.fontBoundingBoxAscent, box.fontBoundingBoxDescent] as const;
      });
      available = boxes[0]![0] === boxes[1]![0] && boxes[0]![1] === boxes[1]![1];
      memo.set(key, available);
    }
    if (available) return index;
  }
  return undefined;
}

/** First available family of a route in this realm, or undefined when the
 * route cannot be judged: a generic keyword (an unnamed face), or a registered
 * FontFace at or before it (an application or web resource whose unicode-range
 * coverage, not installation, decides which clusters it draws). */
function identifiedPrimary(
  context: CanvasFontBoxProbeContext,
  entries: readonly FamilyEntry[],
  registered: ReadonlySet<string>,
  memo: Map<string, boolean>,
): number | undefined {
  const index = primaryIndex(context, entries, memo);
  if (index === undefined || entries[index]!.generic) return undefined;
  for (let entry = 0; entry <= index; entry += 1) {
    if (registered.has(entries[entry]!.name.toLowerCase())) return undefined;
  }
  return index;
}

interface NormalBox { readonly height: number; readonly ascent: number; readonly descent: number }

/** offsetHeight is an integer CSS pixel and Canvas boxes carry sub-pixel
 * quantization (same bounds as the exact-local measurement). */
function scalesLinearly(full: NormalBox, half: NormalBox): boolean {
  return Math.abs(full.height - 2 * half.height) <= 3
    && Math.abs(full.ascent - 2 * half.ascent) <= 1
    && Math.abs(full.descent - 2 * half.descent) <= 1;
}

const LINEAR_PROBE_ATTEMPTS = 4;

/**
 * Measure the CSS normal line of each route's first available face, in the
 * document realm, before layout. Canvas exposes the font box but no leading;
 * a `line-height: normal` strut (U+200B, a closed shadow tree, computed family
 * list checked entry by entry) supplies the missing leading, exactly as the
 * exact-local measurement does for one face. Every strut of a batch is
 * appended once and read once, so a document costs one style/layout pass per
 * linear-scale attempt, never one per route or per segment.
 *
 * Requires the sentinels to be loaded in `owner.fonts`. A request whose first
 * available family cannot be identified, that `admitted` declines, whose
 * computed list differs, or whose box does not scale linearly yields null.
 */
export function measureNativeRouteNormalLines(
  requests: readonly NativeRouteLineRequest[],
  owner: Document,
  admitted: NativeRoutePrimaryAdmission,
): (NativeRouteNormalLine | null)[] {
  const results: (NativeRouteNormalLine | null)[] = requests.map(() => null);
  const context = owner.body ? newProbeContext(owner) : null;
  const view = owner.defaultView;
  if (!context || !view || !sentinelsWork(context)) return results;
  const registered = registeredFamilies(owner.fonts);
  const memo = new Map<string, boolean>();
  // The strut holds only U+200B, a default ignorable that engines render
  // invisibly without font fallback, so its line box is the root inline
  // box's strut alone (CSS 2 §10.8.1; CSS Inline 3 §5.3), whose
  // `line-height: normal` metrics come from the first available font, the
  // face Canvas also reports as the font box. (Fallback glyphs of real text
  // can enlarge a normal line; the strut deliberately has none.) Families
  // before the first available one resolve to no face, so measuring that
  // family alone gives the same strut, shares one strut among every route it
  // paints, and resolves no absent family at the strut sizes.
  interface Pending {
    readonly requests: number[];
    readonly primary: FamilyEntry;
    readonly weight: number;
    readonly style: 'normal' | 'italic';
    em: number;
  }
  const byFace = new Map<string, Pending>();
  const primaries = new Map<number, number>();
  requests.forEach((request, index) => {
    const entries = parseFamilyList(request.familyList);
    if (!entries) return;
    const primary = identifiedPrimary(context, entries, registered, memo);
    if (primary === undefined) return;
    if (!admitted(entries[primary]!.name, request.weight, request.style)) return;
    primaries.set(index, primary);
    const key = JSON.stringify([entries[primary]!.text, request.weight, request.style]);
    const face = byFace.get(key);
    if (face) face.requests.push(index);
    else byFace.set(key, { requests: [index], primary: entries[primary]!, weight: request.weight, style: request.style, em: STRUT_EM });
  });
  let pending = [...byFace.values()];
  for (let attempt = 0; attempt < LINEAR_PROBE_ATTEMPTS && pending.length > 0; attempt += 1) {
    const host = owner.createElement('div');
    host.style.cssText = 'all:initial;position:fixed;visibility:hidden;contain:strict;';
    if (typeof host.attachShadow !== 'function') return results;
    const root = host.attachShadow({ mode: 'closed' });
    const struts = pending.map((item) => [item.em, item.em / 2].map((em) => {
      const strut = owner.createElement('span');
      strut.style.cssText = `all:initial;display:block;white-space:pre;font:${item.style} ${item.weight} ${em}px ${item.primary.text};line-height:normal;`;
      strut.textContent = '\u200b';
      root.append(strut);
      return strut;
    }));
    const next: Pending[] = [];
    owner.body.append(host);
    try {
      pending.forEach((item, index) => {
        const boxes = struts[index]!.map((strut, half): NormalBox | null => {
          const computed = parseFamilyList(view.getComputedStyle(strut).fontFamily);
          if (!computed || !sameFamilyList(computed, [item.primary])) return null;
          const em = half ? item.em / 2 : item.em;
          context.font = `${item.style} ${item.weight} ${em}px ${item.primary.text}`;
          const box = context.measureText('');
          const height = strut.offsetHeight;
          return Number.isFinite(height) && height > 0
            && Number.isFinite(box.fontBoundingBoxAscent) && box.fontBoundingBoxAscent >= 0
            && Number.isFinite(box.fontBoundingBoxDescent) && box.fontBoundingBoxDescent >= 0
            ? { height, ascent: box.fontBoundingBoxAscent, descent: box.fontBoundingBoxDescent }
            : null;
        });
        const [full, half] = boxes;
        if (!full || !half) return;
        if (!scalesLinearly(full, half)) {
          // Engines may clamp large sizes: halve and retry, else decline.
          item.em /= 2;
          next.push(item);
          return;
        }
        for (const requestIndex of item.requests) {
          const request = requests[requestIndex]!;
          results[requestIndex] = Object.freeze({
            familyList: request.familyList,
            weight: request.weight,
            style: request.style,
            primaryIndex: primaries.get(requestIndex)!,
            primaryFamily: item.primary.name,
            primaryFamilyText: item.primary.text,
            emPx: item.em,
            ascentRatio: full.ascent / item.em,
            descentRatio: full.descent / item.em,
            lineGapRatio: (full.height - full.ascent - full.descent) / item.em,
          });
        }
      });
    } catch {
      // An unavailable measurement declines the remaining routes.
      next.length = 0;
    } finally {
      host.remove();
    }
    pending = next;
  }
  return results;
}

/** Admission in the realm that lays out (a worker has no DOM, so its lines are
 * measured by the owning document). Each route must have the same first
 * available family here, identified under the same rules, with the same
 * Canvas font box at the measured em; otherwise this realm draws with a
 * different face and the measurement does not describe it. */
export function verifyNativeRouteNormalLines(
  lines: readonly NativeRouteNormalLine[],
  fontSet: FontFaceSet | null,
): boolean[] {
  const context = newProbeContext();
  if (!context || !sentinelsWork(context)) return lines.map(() => false);
  const registered = registeredFamilies(fontSet);
  const memo = new Map<string, boolean>();
  return lines.map((line) => {
    const entries = parseFamilyList(line.familyList);
    if (!entries || entries[line.primaryIndex]?.name !== line.primaryFamily) return false;
    if (entries[line.primaryIndex]!.text !== line.primaryFamilyText
      || identifiedPrimary(context, entries, registered, memo) !== line.primaryIndex) {
      return false;
    }
    context.font = `${line.style} ${line.weight} ${line.emPx}px ${line.primaryFamilyText}`;
    const box = context.measureText('');
    return Math.abs(box.fontBoundingBoxAscent - line.ascentRatio * line.emPx) < 1e-3
      && Math.abs(box.fontBoundingBoxDescent - line.descentRatio * line.emPx) < 1e-3;
  });
}

// ── Ownership ───────────────────────────────────────────────────────────────

/** The two font strings whose geometry is compared for one route: its first
 * available family, then one sentinel. The families before it resolve to no
 * face and draw nothing, so leaving them out changes no cluster's face. */
export function nativeRouteOwnershipFonts(
  primaryFamilyText: string, weight: number, style: 'normal' | 'italic',
): readonly [string, string] {
  return [
    `${style} ${weight} ${PROBE_EM}px ${primaryFamilyText}, ${quoteCssFontFamily(SENTINEL_FAMILIES[0])}`,
    `${style} ${weight} ${PROBE_EM}px ${primaryFamilyText}, ${quoteCssFontFamily(SENTINEL_FAMILIES[1])}`,
  ];
}

const SENTINEL_A_INK_PX = (SENTINELS[0].inkUp + SENTINELS[0].inkDown) * PROBE_EM / SENTINEL_UPM;

const SENTINEL_ONLY_FONTS = [
  `${PROBE_EM}px ${quoteCssFontFamily(SENTINEL_FAMILIES[0])}`,
  `${PROBE_EM}px ${quoteCssFontFamily(SENTINEL_FAMILIES[1])}`,
] as const;

/** UTF-16 units of the longest grapheme cluster an ownership cache keeps
 * (the control results here and a {@link NamedFamilyOracle}'s results): a
 * longer cluster is proved on every request and never retained. */
const OWNERSHIP_CACHED_CLUSTER_UNITS = 16;

/** Two probe contexts of one font set, the shorthand each currently holds,
 * and that set's route-independent control results. */
interface OwnershipProbe {
  readonly create: () => CanvasFontBoxProbeContext | null;
  readonly contexts: (CanvasFontBoxProbeContext | null | undefined)[];
  readonly fonts: (string | undefined)[];
  readonly controls: Map<string, boolean>;
}

const ownershipProbe = (create: () => CanvasFontBoxProbeContext | null): OwnershipProbe =>
  ({ create, contexts: [], fonts: [], controls: new Map() });

/** The realm's own font set (no owner canvas). */
const realmProbe = ownershipProbe(() => newProbeContext());

/** Advance and ink extents. A sentinel glyph is an ink bar 11 em (A) or
 * 13 em (B) tall; wherever a shaper places it (a zero-advance mark over a
 * wide base, raised or lowered by fallback mark positioning), the cluster's
 * ink box is then at least that tall. */
function geometry(context: CanvasFontBoxProbeContext, text: string): readonly number[] {
  const measured = context.measureText(text);
  return [
    measured.width, measured.actualBoundingBoxLeft, measured.actualBoundingBoxRight,
    measured.actualBoundingBoxAscent, measured.actualBoundingBoxDescent,
  ];
}

function probeContext(probe: OwnershipProbe, index: number): CanvasFontBoxProbeContext | null {
  if (probe.contexts[index] === undefined) probe.contexts[index] = probe.create();
  return probe.contexts[index] ?? null;
}

/** Measure `cluster` under both shorthands. `slot` 0 is the route pair and
 * 1 the sentinel-only control pair: each keeps its own two contexts, so a
 * run of cold clusters judged against one route assigns no font per cluster. */
function measurePair(
  fonts: readonly [string, string], cluster: string, probe: OwnershipProbe = realmProbe, slot = 0,
): (readonly number[])[] | undefined {
  const results: (readonly number[])[] = [];
  for (let pair = 0; pair < 2; pair += 1) {
    const index = slot * 2 + pair;
    const context = probeContext(probe, index);
    if (!context) return undefined;
    if (probe.fonts[index] !== fonts[pair]) {
      context.font = fonts[pair]!;
      probe.fonts[index] = fonts[pair];
    }
    results.push(geometry(context, cluster));
  }
  return results;
}

const sameGeometry = (a: readonly number[], b: readonly number[]) =>
  a.every((value, index) => Number.isFinite(value) && Object.is(value, b[index]));

/** Scalars for which CSS Fonts 4 §5.4 lets an engine choose a face outside
 * the family-list order: emoji presentation (Extended_Pictographic, emoji
 * modifiers, regional indicators, keycap and tag sequences, ZWJ sequences)
 * and the standardized/ideographic variation selectors (U+FE00..FE0F,
 * U+E0100..E01EF). Any part of such a cluster may bypass both sentinels
 * through presentation-based system fallback, which no sentinel can observe,
 * so these clusters are outside the evidence boundary. This set decides the
 * native route line metric for every document (tagged or not), so it stays
 * as it was; the language-admission oracle rejects more (see
 * {@link LANGUAGE_PROOF_EXCLUDED}). */
const PRESENTATION_SELECTED =
  /[\p{Extended_Pictographic}\p{Emoji_Modifier}\p{Regional_Indicator}\u200d\u20e3\ufe00-\ufe0f\u{e0020}-\u{e007f}\u{e0100}-\u{e01ef}]/u;

/** Additional exclusion of the language-admission proof only: every
 * Unicode Variation_Selector (the Mongolian free variation selectors
 * U+180B..180D and U+180F included). A variation sequence selects a glyph
 * variant that a face may or may not provide (Unicode 15 §23.4), so whether a
 * named family "draws" it is not settled by the sentinel comparison; such a
 * cluster never earns a Canvas language. Kept out of
 * {@link PRESENTATION_SELECTED} so that untagged runs keep their existing
 * native line-metric eligibility (library policy). */
const LANGUAGE_PROOF_EXCLUDED = /\p{Variation_Selector}/u;

/** Whether the route's first available family (or an absent family before
 * it) draws every part of `cluster`; undefined when the realm cannot judge.
 *
 * The comparison is evidence only for scalars the engine routes through the
 * CSS family list. Clusters with presentation-selected scalars (see above)
 * are not owned. Every other cluster must first pass a control: with a
 * sentinel as the only family, the two registrations must give different
 * geometry, i.e. a sentinel demonstrably draws the cluster when no listed
 * face does. A cluster failing either is not owned (the span keeps its
 * Canvas box). Each sentinel has its own private context, so a font string
 * is parsed only when the route changes. */
export function nativeRouteClusterOwned(
  fonts: readonly [string, string], cluster: string, probe: OwnershipProbe = realmProbe,
): boolean | undefined {
  if (PRESENTATION_SELECTED.test(cluster)) return false;
  // Oversized clusters are judged each time: neither cache retains their text.
  const cacheable = cluster.length <= OWNERSHIP_CACHED_CLUSTER_UNITS;
  let routed = cacheable ? probe.controls.get(cluster) : undefined;
  if (routed === undefined) {
    const control = measurePair(SENTINEL_ONLY_FONTS, cluster, probe, 1);
    // Invalid readings are no evidence: negating an equality over them would
    // pass the control, so non-finite geometry cannot judge (nor be cached).
    if (!control || !control.every((geometry) => geometry.every(Number.isFinite))) return undefined;
    routed = !sameGeometry(control[0]!, control[1]!);
    if (cacheable) {
      if (probe.controls.size >= 4096) probe.controls.delete(probe.controls.keys().next().value!);
      probe.controls.set(cluster, routed);
    }
  }
  if (!routed) return false;
  const measured = measurePair(fonts, cluster, probe);
  if (!measured) return undefined;
  // Equal geometry proves no sentinel glyph only while the cluster's own ink
  // could not mask one: an ink box at least as tall as sentinel A's bar
  // (stacked marks) may contain it, so such a cluster is not owned.
  const [, , , ascent, descent] = measured[0]!;
  return sameGeometry(measured[0]!, measured[1]!) && ascent! + descent! < SENTINEL_A_INK_PX;
}

/** Bounds on what a {@link NamedFamilyOracle} remembers: results and the
 * family lists they refer to (in UTF-16 units; a longer list is judged
 * without being remembered). */
const NAMED_FAMILY_RESULT_LIMIT = 8192;
const NAMED_FAMILY_LIST_UNITS = 65536;

/** Judges whether NAMED families of a Canvas route draw text, in the font
 * set of the canvas that measures and paints it. */
export interface NamedFamilyOracle {
  /** Whether the route sentinels are registered AND applied in this font
   * set: to the measuring context and to the two contexts that perform the
   * proof. Judged on first use and fixed for the oracle's life. */
  readonly sentinelsApply: boolean;
  /** Whether this font set can judge a run language: Canvas `lang` exists
   * and {@link sentinelsApply}. Judged on first use and fixed for the
   * oracle's life. */
  readonly available: boolean;
  /** Whether a named family of `familyList` draws every grapheme cluster of
   * `text`; false when this font set cannot judge. */
  draws(familyList: string, weight: number, style: 'normal' | 'italic', text: string): boolean;
}

/**
 * An oracle over the font set of `owner` (the measuring 2D context; its
 * canvas's document, else the realm), proving with the route sentinels
 * ({@link nativeRouteClusterOwned}) which family draws each cluster.
 *
 * Only named families are judged: the list is cut at its first generic
 * keyword. A generic resolves to faces chosen per content language (CSS
 * Fonts 4 §5.1), and so does the engine's last-resort fallback (§5.4); a
 * named family is matched in list order whatever the language. Families
 * after a generic could still be reached when it lacks a glyph, so the cut
 * is a conservative evidence boundary, not a claim that they never draw.
 * Text drawn entirely by named families keeps its faces under any Canvas
 * `lang`: the language cannot replace a face by a fallback one (with other
 * advances, line metrics or weights). It can still select forms (locl) and
 * positioning inside those faces, which may change advances; measurement and
 * paint both apply it.
 *
 * The capability is judged once, by this oracle, and shared by every consumer: a
 * worker admits a document only where `sentinelsApply` holds (as it does in
 * the main realm, which retains its registrations), and the text service's
 * fingerprint records `available`. The oracle belongs to one font snapshot of
 * a layout session, which retains its sentinel registrations for its life.
 */
export function createNamedFamilyOracle(
  owner?: unknown, factory?: SymbolProbeCanvasFactory,
): NamedFamilyOracle {
  const probe = ownershipProbe(() => {
    try {
      return (scratchCanvas(owner, factory, 1)?.getContext('2d') as CanvasFontBoxProbeContext | null | undefined) ?? null;
    } catch {
      return null;
    }
  });
  const measuring = owner as Partial<CanvasFontBoxProbeContext> | null | undefined;
  // Judged on first use, once: a context that has no Canvas `lang` is never
  // measured for a run language, and no proof canvas exists before then.
  let applied: boolean | undefined;
  const sentinelsApply = (): boolean => applied ??= (() => {
    const context = measuring;
    if (!context || typeof context.measureText !== 'function' || typeof context.font !== 'string') return false;
    const previous = context.font;
    try {
      if (!sentinelsWork(context as CanvasFontBoxProbeContext)) return false;
    } finally {
      context.font = previous;
    }
    // The proof contexts keep the sentinel shorthand: measurePair assigns
    // each route before measuring (their recorded shorthand is still unset).
    // All four proof contexts: the route pair and the control pair.
    return [0, 1, 2, 3].every((index) => {
      const proof = probeContext(probe, index);
      return proof !== null && sentinelsWork(proof);
    });
  })();
  let admits: boolean | undefined;
  // Grapheme segmentation is part of the capability: clusters are what the
  // proof certifies (see draws), so without it nothing could be admitted.
  const available = (): boolean => admits ??= !!measuring && typeof measuring === 'object'
    && 'lang' in measuring && getGraphemeSegmenter() !== null && sentinelsApply();
  const listIds = new Map<string, number>();
  let listUnits = 0;
  const results = new Map<string, boolean>();
  return {
    get sentinelsApply() { return sentinelsApply(); },
    get available() { return available(); },
    draws(familyList, weight, style, text) {
      if (text.length === 0 || !available()) return false;
      const entries = parseFamilyList(familyList);
      if (!entries) return false;
      const firstGeneric = entries.findIndex((entry) => entry.generic);
      const named = (firstGeneric < 0 ? entries : entries.slice(0, firstGeneric)).map((entry) => entry.text);
      if (named.length === 0) return false;
      const fonts = nativeRouteOwnershipFonts(named.join(', '), weight, style);
      // A list longer than the whole budget is never interned: its clusters
      // are judged without result caching.
      let listId = fonts[0].length > NAMED_FAMILY_LIST_UNITS ? undefined : listIds.get(fonts[0]);
      if (listId === undefined && fonts[0].length <= NAMED_FAMILY_LIST_UNITS) {
        if (listIds.size >= NAMED_FAMILY_RESULT_LIMIT || listUnits + fonts[0].length > NAMED_FAMILY_LIST_UNITS) {
          listIds.clear();
          results.clear();
          listUnits = 0;
        }
        listId = listIds.size;
        listIds.set(fonts[0], listId);
        listUnits += fonts[0].length;
      }
      // CSS Fonts 4 §5.4 matches a face per grapheme cluster; scalar probes
      // cannot certify a cluster, so without a segmenter nothing is admitted.
      const segmenter = getGraphemeSegmenter();
      if (!segmenter) return false;
      const clusters: Iterable<string> = {
        *[Symbol.iterator]() { for (const part of segmenter.segment(text)) yield part.segment; },
      };
      for (const cluster of clusters) {
        const key = listId !== undefined && cluster.length <= OWNERSHIP_CACHED_CLUSTER_UNITS
          ? `${listId}\u0000${cluster}` : undefined;
        if (LANGUAGE_PROOF_EXCLUDED.test(cluster)) return false;
        let owned = key !== undefined ? results.get(key) : undefined;
        if (owned === undefined) {
          const judged = nativeRouteClusterOwned(fonts, cluster, probe);
          if (judged === undefined) return false;
          owned = judged;
          if (key !== undefined) {
            if (results.size >= NAMED_FAMILY_RESULT_LIMIT) results.delete(results.keys().next().value!);
            results.set(key, owned);
          }
        }
        if (!owned) return false;
      }
      return true;
    },
  };
}

/** The sentinels are registered and applied in this realm (once per prepared
 * route; an unavailable sentinel would let fallback faces pass as owned). */
export function nativeRouteSentinelsWork(): boolean {
  const context = newProbeContext();
  return context !== null && sentinelsWork(context);
}

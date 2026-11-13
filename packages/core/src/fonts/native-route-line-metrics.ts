import type { CanvasFontBoxProbeContext } from './canvas-font-box.js';
import { quoteCssFontFamily } from '../internal/css-font-family.js';
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

const ownershipContexts: (CanvasFontBoxProbeContext | null)[] = [];
const ownershipFonts: (string | undefined)[] = [];
/** Route-independent control results, bounded like the per-route caches. */
const sentinelControls = new Map<string, boolean>();

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

function measurePair(fonts: readonly [string, string], cluster: string): (readonly number[])[] | undefined {
  const results: (readonly number[])[] = [];
  for (let index = 0; index < 2; index += 1) {
    ownershipContexts[index] ??= newProbeContext();
    const context = ownershipContexts[index];
    if (!context) return undefined;
    if (ownershipFonts[index] !== fonts[index]) {
      context.font = fonts[index]!;
      ownershipFonts[index] = fonts[index];
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
 * and variation selectors. Any part of such a cluster may bypass both
 * sentinels through presentation-based system fallback, which no sentinel
 * can observe, so these clusters are outside the evidence boundary. */
const PRESENTATION_SELECTED =
  /[\p{Extended_Pictographic}\p{Emoji_Modifier}\p{Regional_Indicator}\u200d\u20e3\ufe00-\ufe0f\u{e0020}-\u{e007f}\u{e0100}-\u{e01ef}]/u;

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
  fonts: readonly [string, string], cluster: string,
): boolean | undefined {
  if (PRESENTATION_SELECTED.test(cluster)) return false;
  let routed = sentinelControls.get(cluster);
  if (routed === undefined) {
    const control = measurePair(SENTINEL_ONLY_FONTS, cluster);
    if (!control) return undefined;
    routed = !sameGeometry(control[0]!, control[1]!);
    if (sentinelControls.size >= 4096) sentinelControls.delete(sentinelControls.keys().next().value!);
    sentinelControls.set(cluster, routed);
  }
  if (!routed) return false;
  const measured = measurePair(fonts, cluster);
  if (!measured) return undefined;
  // Equal geometry proves no sentinel glyph only while the cluster's own ink
  // could not mask one: an ink box at least as tall as sentinel A's bar
  // (stacked marks) may contain it, so such a cluster is not owned.
  const [, , , ascent, descent] = measured[0]!;
  return sameGeometry(measured[0]!, measured[1]!) && ascent! + descent! < SENTINEL_A_INK_PX;
}

/** The sentinels are registered and applied in this realm (once per prepared
 * route; an unavailable sentinel would let fallback faces pass as owned). */
export function nativeRouteSentinelsWork(): boolean {
  const context = newProbeContext();
  return context !== null && sentinelsWork(context);
}

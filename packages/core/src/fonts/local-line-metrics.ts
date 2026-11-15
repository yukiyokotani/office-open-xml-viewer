import { retainFace, releaseFaces } from './font-registry.js';
import { withFontCeiling } from './font-registry.js';
import type { CanvasFontRoute } from './canvas-route.js';
import type { ResolvedFontMetric } from './resource-metrics.js';
import { quoteCssFontFamily } from '../internal/css-font-family.js';
import { getGraphemeSegmenter } from '../text/sea-break.js';
import type { CanvasFontBoxProbeContext } from './canvas-font-box.js';
import { nativeRouteClusterOwned, nativeRouteOwnershipFonts, nativeRouteSentinelsWork } from './native-route-line-metrics.js';

export interface LocalFontNormalLine {
  readonly ascentRatio: number;
  readonly descentRatio: number;
  readonly lineGapRatio: number;
}

// The FontFace is the retained resource, not an authored family or an alias
// shared across FontFaceSets. The weak cache retains no DOM or context.
const measuredLines = new WeakMap<FontFace, { emPx: number; metric: LocalFontNormalLine | null }>();

/** Measure the normal line of an already-loaded exact local() paint resource.
 * Canvas exposes its font box but omits leading. A CSS normal-line strut adds
 * that missing measurement; U+200B creates a line without fallback glyph ink.
 * This is browser geometry, NOT access to OpenType hhea/OS/2 tables. In
 * particular it cannot establish usWin, USE_TYPO_METRICS or code-page bits.
 * Callers must decline projections requiring those unavailable fields.
 *
 * The caller supplies the probe em (a known design UPM avoids CSS-pixel
 * quantization). This runs during resource preparation, never per segment or
 * during paint. A worker/foreign FontFaceSet has no matching DOM strut and
 * returns null, rather than silently measuring a different realm's face. */
export function measureLocalFontNormalLine(
  face: FontFace,
  fontSet: FontFaceSet,
  context: CanvasFontBoxProbeContext,
  emPx: number,
  ownerDocument: Document | null,
): LocalFontNormalLine | null {
  if (face.status !== 'loaded') return null;
  const cached = measuredLines.get(face);
  if (cached?.emPx === emPx) return cached.metric;
  const metric = measureFontFaceNormalLine(face, face.weight, face.style, fontSet, context, emPx, ownerDocument);
  measuredLines.set(face, { emPx, metric });
  return metric;
}

/** {@link measureLocalFontNormalLine} for an explicit weight/style tuple of
 * the face's family (a variable face measured at one instance, or a tuple
 * the family synthesizes). Not cached: callers own their reuse. The family
 * must resolve to `face` alone for the measurement to describe it. */
export function measureFontFaceNormalLine(
  face: FontFace,
  weight: string | number,
  style: string,
  fontSet: FontFaceSet,
  context: CanvasFontBoxProbeContext,
  emPx: number,
  ownerDocument: Document | null,
): LocalFontNormalLine | null {
  return measureFontFaceNormalLines([{ face, weight, style }], fontSet, context, emPx, ownerDocument)[0]!;
}

/** One face × tuple of {@link measureFontFaceNormalLines}. */
export interface FontFaceNormalLineRequest {
  readonly face: FontFace;
  readonly weight: string | number;
  readonly style: string;
}

/** {@link measureFontFaceNormalLine} for many faces × tuples at once, with
 * the same result for each request: every strut of a linear-scale attempt is
 * appended in one host and read after one style/layout pass, so a batch
 * costs at most {@link LINEAR_PROBE_ATTEMPTS} layouts, never one (or two)
 * per request. Each request keeps its own attempt sequence: it halves its em
 * only while its own boxes do not scale linearly, and declines once a
 * measurement fails. The struts are independent blocks (`display: block`,
 * one U+200B each), so sharing a host changes no strut's height. */
export function measureFontFaceNormalLines(
  requests: readonly FontFaceNormalLineRequest[],
  fontSet: FontFaceSet,
  context: CanvasFontBoxProbeContext,
  emPx: number,
  ownerDocument: Document | null,
): (LocalFontNormalLine | null)[] {
  const results: (LocalFontNormalLine | null)[] = requests.map(() => null);
  if (!Number.isFinite(emPx) || emPx <= 0 || !ownerDocument?.body || ownerDocument.fonts !== fontSet) return results;
  // Engines may clamp a large requested size, so a successful measurement does
  // not prove the requested scale. Accept an em only when the box measured at
  // half that size scales linearly; otherwise halve and retry, and decline
  // (catalog fallback) when no measured scale is linear.
  let pending = requests.flatMap((request, index) => request.face.status === 'loaded' ? [{ index, request, em: emPx }] : []);
  for (let attempt = 0; attempt < LINEAR_PROBE_ATTEMPTS && pending.length > 0; attempt++) {
    const boxes = measureNormalBoxes(pending.flatMap(({ request, em }) => [
      { family: request.face.family, weight: request.weight, style: request.style, em },
      { family: request.face.family, weight: request.weight, style: request.style, em: em / 2 },
    ]), context, ownerDocument);
    const next: typeof pending = [];
    pending.forEach((item, position) => {
      const full = boxes[position * 2];
      const half = boxes[position * 2 + 1];
      if (!full || !half) return;
      if (scalesLinearly(full, half)) {
        results[item.index] = Object.freeze({
          ascentRatio: full.ascent / item.em,
          descentRatio: full.descent / item.em,
          lineGapRatio: (full.height - full.ascent - full.descent) / item.em,
        });
        return;
      }
      item.em /= 2;
      next.push(item);
    });
    pending = next;
  }
  return results;
}

/** Halvings tried before declining; bounds probe work per resource. */
const LINEAR_PROBE_ATTEMPTS = 4;

interface NormalBox { readonly height: number; readonly ascent: number; readonly descent: number }

interface NormalBoxRequest {
  readonly family: string;
  readonly weight: string | number;
  readonly style: string;
  readonly em: number;
}

/** The CSS normal line (strut height) and Canvas font box of each request,
 * from one host appended once: one style/layout pass for the whole list. */
function measureNormalBoxes(
  requests: readonly NormalBoxRequest[], context: CanvasFontBoxProbeContext, ownerDocument: Document,
): (NormalBox | null)[] {
  const results: (NormalBox | null)[] = requests.map(() => null);
  // The struts live in a closed shadow tree so application stylesheets
  // (including !important rules) cannot restyle them; each computed family
  // is checked as well, so the DOM height belongs to the face Canvas measures.
  const host = ownerDocument.createElement('div');
  host.style.cssText = 'all:initial;position:fixed;visibility:hidden;contain:strict;';
  if (typeof host.attachShadow !== 'function') return results;
  const previousFont = context.font;
  try {
    const root = host.attachShadow({ mode: 'closed' });
    const struts = requests.map(({ family, weight, style, em }) => {
      const font = `${style} ${weight} ${em}px ${quoteCssFontFamily(family)}`;
      const probe = ownerDocument.createElement('span');
      probe.style.cssText = `all:initial;display:block;white-space:pre;font:${font};line-height:normal;`;
      probe.textContent = '\u200b';
      root.append(probe);
      return { probe, font, family };
    });
    ownerDocument.body.append(host);
    const view = ownerDocument.defaultView;
    if (!view) return results;
    struts.forEach(({ probe, font, family }, index) => {
      try {
        if (unquoted(view.getComputedStyle(probe).fontFamily) !== unquoted(quoteCssFontFamily(family))) return;
        context.font = font;
        const box = context.measureText('');
        const height = probe.offsetHeight;
        const ascent = box.fontBoundingBoxAscent;
        const descent = box.fontBoundingBoxDescent;
        results[index] = Number.isFinite(height) && height > 0 && Number.isFinite(ascent)
          && ascent >= 0 && Number.isFinite(descent) && descent >= 0
          ? { height, ascent, descent } : null;
      } catch {
        // An unavailable measurement leaves the caller's catalog fallback intact.
      }
    });
    return results;
  } catch {
    return requests.map(() => null);
  } finally {
    host.remove();
    context.font = previousFont;
  }
}

/** Engines differ in whether a computed single family keeps its quotes. */
function unquoted(family: string): string {
  return family.trim().replace(/^(["'])(.*)\1$/, '$2');
}

/** offsetHeight is an integer CSS pixel, so each reading carries up to one
 * pixel of rounding; Canvas font boxes carry sub-pixel quantization. */
function scalesLinearly(full: NormalBox, half: NormalBox): boolean {
  return Math.abs(full.height - 2 * half.height) <= 3
    && Math.abs(full.ascent - 2 * half.ascent) <= 1
    && Math.abs(full.descent - 2 * half.descent) <= 1;
}

/** A second registration of the exact source, with only size-adjust changed.
 * CSS Fonts 5 §4.4 (size-adjust) scales outlines and advances, not cmap.
 * Native cmap/shaping/fallback selection is unchanged: owned glyphs
 * have zero ink/advance; fallback glyphs retain their own ink/advance. This is
 * a controlled browser fallback test, not a comparison against a guessed
 * missing-font width or a script/family coverage table. Default ignorables
 * that draw no glyph contribute no geometry in either route.
 * Retain/release this registration with the paint face. */
export async function loadLocalFontCoverageProbe(
  face: FontFace, source: string, fontSet: FontFaceSet,
): Promise<FontFace | null> {
  const family = `${face.family}_coverage`;
  const { face: probe } = retainFace(`local-coverage:${source}:${face.weight}:${face.style}`, fontSet, () => {
    const created = new FontFace(family, source, {
      weight: face.weight, style: face.style, sizeAdjust: '0%',
    } as FontFaceDescriptors & { sizeAdjust: string });
    fontSet.add(created);
    return created;
  });
  try {
    const loaded = await withFontCeiling(probe.load());
    if (loaded && probe.status === 'loaded'
      && (probe as FontFace & { sizeAdjust?: string }).sizeAdjust === '0%'
      && coverageOracleWorks(probe, face)) return probe;
  } catch { /* Optional coverage unavailable: decline resource geometry. */ }
  releaseFaces([probe]);
  return null;
}

/** A readable size-adjust descriptor does not prove the engine applies it to
 * Canvas geometry. Before trusting zero geometry as ownership, the probe must
 * measure the face's own Latin glyph as zero in every dimension, while the
 * same glyph in the paint face keeps positive geometry (so the zero comes
 * from size-adjust, not from a degenerate measurement). An engine that only
 * zeroes advances (WebKit keeps ink extents) or a face without that glyph
 * declines the measured projection and keeps the catalog path. No
 * other-script fallback is measured here: that would start a system
 * fallback-font search on the first-display path. */
const OWNED_PROBE = 'A';

function coverageOracleWorks(probe: FontFace, paint: FontFace): boolean {
  // Once per loaded probe, so a fresh context costs nothing measurable.
  const context = newCoverageContext();
  if (!context) return false;
  try {
    context.font = `${probe.style} ${probe.weight} 100px ${quoteCssFontFamily(probe.family)}`;
    const owned = context.measureText(OWNED_PROBE);
    context.font = `${paint.style} ${paint.weight} 100px ${quoteCssFontFamily(paint.family)}`;
    const painted = context.measureText(OWNED_PROBE);
    return owned.width === 0 && owned.actualBoundingBoxAscent === 0
      && owned.actualBoundingBoxDescent === 0 && painted.width > 0;
  } catch {
    return false;
  }
}

function newCoverageContext(): CanvasFontBoxProbeContext | null {
  return typeof OffscreenCanvas !== 'undefined'
    ? new OffscreenCanvas(1, 1).getContext('2d')
    : typeof document !== 'undefined' ? document.createElement('canvas').getContext('2d') : null;
}

let coverageContext: CanvasFontBoxProbeContext | null | undefined;
interface CoverageRoute {
  /** Native judgement of one cluster; undefined when the realm cannot judge. */
  readonly measure: (cluster: string) => boolean | undefined;
  /** Text results (bounded) and the per-cluster answers they are built from. */
  readonly results: Map<string, boolean>;
  readonly clusters: Map<string, boolean>;
}

const PRINTABLE_ASCII = /^[\x20-\x7e]*$/;

/** Browser font fallback chooses a face per grapheme cluster, so ownership of
 * a text is ownership of each of its clusters. Printable ASCII clusters are
 * single characters. Without a segmenter the whole text is one unit. */
function clustersOf(text: string): Iterable<string> {
  if (PRINTABLE_ASCII.test(text)) return text;
  // Created on first non-ASCII text rather than at module evaluation: an ICU
  // break iterator costs several ms per realm and most texts never need one.
  const graphemes = getGraphemeSegmenter();
  if (!graphemes) return [text];
  return Array.from(graphemes.segment(text), (part) => part.segment);
}
const coverageResults = new WeakMap<ResolvedFontMetric, WeakMap<CanvasFontRoute, CoverageRoute>>();

/** Ask the same native shaper that paints the registered local source. No
 * per-span DOM reads or font loads; bounded weak caches die with the metric.
 * Any fallback ink/advance revokes the primary resource's vertical projection,
 * allowing the existing per-segment Canvas measurement to size that span.
 *
 * A native route metric (`nativeRoutePrimaryFamily`) describes the route's
 * first available face instead of a registered source; its clusters are judged
 * with the route sentinels (native-route-line-metrics.ts). */
export function localFontMetricCoversText(
  metric: ResolvedFontMetric, text: string, route: CanvasFontRoute | undefined,
): boolean {
  const family = metric.glyphCoverageProbeFamily;
  const nativeList = metric.nativeRoutePrimaryFamily;
  if ((!family && !nativeList) || !text) return !metric.requiresGlyphCoverage || text.length === 0;
  if (!route) return false;
  let byRoute = coverageResults.get(metric);
  if (!byRoute) { byRoute = new WeakMap(); coverageResults.set(metric, byRoute); }
  let prepared = byRoute.get(route);
  if (!prepared) {
    if (nativeList) {
      // The metric is attached to exactly this route (by its family list and
      // tuple); this realm's sentinels must work.
      if (!nativeRouteSentinelsWork()) return false;
      const fonts = nativeRouteOwnershipFonts(nativeList, metric.weight ?? 400, metric.style ?? 'normal');
      prepared = {
        measure: (cluster) => nativeRouteClusterOwned(fonts, cluster),
        results: new Map(),
        clusters: new Map(),
      };
    } else {
      const primary = quoteCssFontFamily(metric.family);
      if (route.familyList !== primary && !route.familyList.startsWith(`${primary},`)) return false;
      // Preserve the complete paint cascade. A caller-owned tail face may cover
      // glyphs that generic/system fallback cannot, including private-use scalars.
      const familyList = quoteCssFontFamily(family!) + route.familyList.slice(primary.length);
      const font = `${metric.style ?? 'normal'} ${metric.weight ?? 400} 100px ${familyList}`;
      prepared = {
        measure: (cluster) => measureCoverage(font, cluster),
        results: new Map(),
        clusters: new Map(),
      };
    }
    byRoute.set(route, prepared);
  }
  const cache = prepared.results;
  const cached = cache.get(text);
  if (cached !== undefined) return cached;
  // Native measurement happens once per distinct cluster, not per distinct
  // text, so a document costs about as many probes as it has distinct
  // clusters. Each cluster is measured on its own: a zero result for a whole
  // text does not prove every cluster is owned (a joiner can absorb an
  // unsupported combining mark).
  let covered = true;
  for (const cluster of clustersOf(text)) {
    let owned = prepared.clusters.get(cluster);
    if (owned === undefined) {
      const measured = prepared.measure(cluster);
      if (measured === undefined) return false;
      owned = measured;
      rememberCluster(prepared, cluster, owned);
    }
    if (!owned) { covered = false; break; }
  }
  if (cache.size >= 1024) cache.delete(cache.keys().next().value!);
  cache.set(text, covered);
  return covered;
}

function rememberCluster(prepared: CoverageRoute, cluster: string, owned: boolean): void {
  // Resource-governance bound, unrelated to layout or coverage decisions.
  if (prepared.clusters.size >= 4096) prepared.clusters.delete(prepared.clusters.keys().next().value!);
  prepared.clusters.set(cluster, owned);
}

// The coverage context is private to this module, so its font is switched
// only when the route changes instead of being restored after every probe.
let coverageFont: string | undefined;

function measureCoverage(font: string, text: string): boolean | undefined {
  coverageContext ??= newCoverageContext();
  if (!coverageContext) return undefined;
  if (coverageFont !== font) { coverageContext.font = font; coverageFont = font; }
  const result = coverageContext.measureText(text);
  return result.width === 0 && result.actualBoundingBoxAscent === 0
    && result.actualBoundingBoxDescent === 0;
}

/** Measure a worker-resolved exact local source in the DOM realm. The caller
 * must match the returned numbers to that source and tuple; aliases alone are
 * insufficient. The transient probe is refcounted and always released. */
export async function measureLocalFontSourceNormalLine(
  source: string, family: string, weight: number, style: 'normal' | 'italic',
  emPx: number, owner: Document,
): Promise<LocalFontNormalLine | null> {
  const { face } = retainFace(`local-normal:${source}:${weight}:${style}`, owner.fonts, () => {
    const created = new FontFace(`${family}_normal`, source, { weight: String(weight), style });
    owner.fonts.add(created);
    return created;
  });
  try {
    const loaded = await withFontCeiling(face.load());
    const context = owner.createElement('canvas').getContext('2d');
    return loaded && context ? measureLocalFontNormalLine(face, owner.fonts, context, emPx, owner) : null;
  } catch { return null; }
  finally { releaseFaces([face]); }
}

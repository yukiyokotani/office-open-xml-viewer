import { retainFace, releaseFaces } from './font-registry.js';
import { activeFontSet, withFontCeiling } from './preload.js';

/** A family-level local-font measurement request. This API is shared by all
 * OOXML formats; each format decides which Office behavior requires a measured
 * multiplier and supplies the evidence-backed value. Do not add family-specific
 * requests when the selected font resource can be measured directly. */
export interface LocalFontMetricRequest {
  /** Authored OOXML family name used as the lookup key. */
  family: string;
  /** Exact local() names, in preference order. No generic fallback is allowed. */
  localNames: readonly string[];
  /** Evidence-backed Office single-line multiplier applied to the local face's
   * design box. Omit when the face is loaded only as an exact geometry alias. */
  lineHeightMultiplier?: number;
  /** Descriptor of this exact full/PostScript face name. A styled face requires
   * its own request and exact local name; capabilities are never synthesized. */
  weight?: number;
  style?: 'normal' | 'italic';
}

export interface ResolvedLocalFontMetric {
  /** Isolated FontFace family registered for Canvas measure and paint. */
  family: string;
  /** Office single-line height divided by em, only for documented profiles. */
  lineHeightRatio?: number;
  /** Authored Canvas tuple isolated behind this alias. */
  requestedFamily?: string;
  weight?: number;
  style?: 'normal' | 'italic';
  /** Canonical exact local() source. This is a route identity, not a claim that
   * native Canvas geometry is portable across engines or machines. */
  sourceIdentity?: string;
  /** Explicit UA synthesis policy. Production exact-local records are false;
   * deterministic test fixtures may opt into and label synthesis explicitly. */
  synthesized?: boolean;
}

export interface LoadedLocalFontMetrics {
  /** Refcounted faces retained by this caller; release on document destroy. */
  faces: FontFace[];
  /** Normalized authored family → exact local face and measured line ratio. */
  metrics: Record<string, ResolvedLocalFontMetric>;
}

export function normalizeLocalFontMetricFamily(family: string): string {
  return family.trim().toLowerCase();
}

function quoteLocalName(name: string): string {
  return `local("${name.replaceAll('\\', '\\\\').replaceAll('"', '\\"')}")`;
}

function collisionFreeAlias(signature: string): string {
  // Fixed-width Unicode-scalar encoding is injective and CSS-family safe. A
  // short digest is not sufficient here: two sources sharing an alias would
  // make Canvas face selection ambiguous even if registry keys stayed distinct.
  return `__ooxml_local_${[...signature]
    .map((scalar) => (scalar.codePointAt(0) ?? 0).toString(16).padStart(6, '0'))
    .join('')}`;
}

function measureContext():
  | CanvasRenderingContext2D
  | OffscreenCanvasRenderingContext2D
  | null {
  if (typeof OffscreenCanvas !== 'undefined') {
    return new OffscreenCanvas(1, 1).getContext('2d');
  }
  if (typeof document !== 'undefined' && document?.createElement) {
    return document.createElement('canvas').getContext('2d');
  }
  return null;
}

/** Load positively identified local faces without accepting CSS fallback.
 * Every request names one exact full/PostScript face and one descriptor tuple.
 * CSS Fonts local() is not family discovery, so a regular face never implies a
 * bold/italic capability. Native geometry belongs to later retained placements;
 * this snapshot records only the immutable prepared route.
 */
export async function loadLocalFontMetrics(
  requests: readonly LocalFontMetricRequest[],
): Promise<LoadedLocalFontMetrics> {
  const set = activeFontSet();
  if (!set || typeof FontFace === 'undefined') return { faces: [], metrics: {} };

  const faces: FontFace[] = [];
  const metrics: Record<string, ResolvedLocalFontMetric> = {};
  type PreparedRequest = LocalFontMetricRequest & {
    family: string;
    normalizedFamily: string;
    source: string;
    weight: number;
    style: 'normal' | 'italic';
  };
  const grouped = new Map<string, { source: string; requests: PreparedRequest[] }>();
  for (const request of requests) {
    const family = request.family.trim();
    const localNames = request.localNames.map((name) => name.trim()).filter(Boolean);
    if (!family || localNames.length === 0) continue;
    if (request.lineHeightMultiplier != null && !(request.lineHeightMultiplier > 0)) continue;
    const weight = request.weight ?? 400;
    const style = request.style ?? 'normal';
    if (!(weight >= 100 && weight <= 900) || (style !== 'normal' && style !== 'italic')) continue;
    const source = localNames.map(quoteLocalName).join(', ');
    const normalizedFamily = normalizeLocalFontMetricFamily(family);
    const signature = `local-face:${source}:${weight}:${style}`;
    const group = grouped.get(signature) ?? { source, requests: [] };
    group.requests.push({ ...request, family, normalizedFamily, source, weight, style });
    grouped.set(signature, group);
  }

  for (const [signature, group] of grouped) {
    const alias = collisionFreeAlias(signature);
    const { face } = retainFace(signature, set, () => {
      const tuple = group.requests[0];
      const created = new FontFace(alias, group.source, {
        weight: String(tuple.weight),
        style: tuple.style,
      });
      set.add(created);
      return created;
    });

    try {
      // A second document can retain the shared face while the first holder's
      // load is still in flight. FontFace.load() is idempotent and joins that
      // pending load, so await it for every holder before measuring; `isNew`
      // alone is not a sufficient loaded-state guarantee under concurrent opens.
      const loaded = await withFontCeiling(face.load());
      if (!loaded || face.status !== 'loaded') throw new Error('local font load timed out');
      let hasRoute = false;
      for (const request of group.requests) {
        let lineHeightRatio: number | undefined;
        if (request.lineHeightMultiplier != null) {
          const ctx = measureContext();
          if (!ctx) continue;
          ctx.font = `${request.style} ${request.weight} 100px "${alias}"`;
          const measured = ctx.measureText('Hg国');
          const ascent = measured.fontBoundingBoxAscent;
          const descent = measured.fontBoundingBoxDescent;
          if (!(Number.isFinite(ascent) && Number.isFinite(descent) && ascent + descent > 0)) {
            continue;
          }
          lineHeightRatio = ((ascent + descent) / 100) * request.lineHeightMultiplier;
        }
        const key = request.weight === 400 && request.style === 'normal'
          ? request.normalizedFamily
          : `${request.normalizedFamily}:${request.weight}:${request.style}`;
        metrics[key] = {
          family: alias,
          ...(lineHeightRatio == null ? {} : { lineHeightRatio }),
          requestedFamily: request.family,
          weight: request.weight,
          style: request.style,
          sourceIdentity: request.source,
          synthesized: false,
        };
        hasRoute = true;
      }
      if (!hasRoute) throw new Error('exact local font route unavailable');
      faces.push(face);
    } catch {
      releaseFaces([face]);
    }
  }
  return { faces, metrics };
}

export function unloadLocalFontMetrics(faces: Iterable<FontFace>): void {
  releaseFaces(faces);
}

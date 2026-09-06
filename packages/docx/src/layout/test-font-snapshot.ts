import type { ResolvedFontMetric } from '@silurus/ooxml-core';

export interface TestFontFace {
  readonly family: string;
  readonly resolvedFamily?: string;
  readonly weight?: number;
  readonly style?: 'normal' | 'italic';
  /** Tests that intentionally model UA synthesis must say so explicitly. */
  readonly synthesized?: boolean;
  readonly lineHeightRatio?: number;
}

/** Explicit deterministic authored-face inventory for renderer unit fixtures.
 * Real document opens populate this map only through positive local()/embedded/
 * web-font loading. Tests use synthetic Canvas contexts, so they must state the
 * faces those contexts model instead of relying on host-local resolution. */
export function testFontSnapshot(
  faces: readonly TestFontFace[],
): Record<string, ResolvedFontMetric> {
  const metrics: Record<string, ResolvedFontMetric> = {};
  for (const face of faces) {
    const family = face.family;
    const normalized = family.trim().toLowerCase();
    if (!normalized) continue;
    const weight = face.weight ?? 400;
    const style = face.style ?? 'normal';
    const key = weight === 400 && style === 'normal'
      ? normalized
      : `${normalized}:${weight}:${style}`;
    metrics[key] = {
      family: face.resolvedFamily ?? family,
      requestedFamily: family,
      weight,
      style,
      sourceIdentity: `test-fixture:${normalized}:${weight}:${style}`,
      synthesized: face.synthesized ?? false,
      ...(face.lineHeightRatio === undefined ? {} : { lineHeightRatio: face.lineHeightRatio }),
    };
  }
  return metrics;
}

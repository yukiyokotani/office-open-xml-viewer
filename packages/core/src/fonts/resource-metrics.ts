/** Geometry and identity retained for one concrete font resource. The family
 * name selects the resource; numeric metrics come only from that resource's
 * bytes or another explicit resource owner, never from a family-name table. */
export interface ResolvedFontMetric {
  /** Concrete FontFace family registered for Canvas measure and paint. */
  family: string;
  /** General single-line height divided by em, when provided by the resource
   * owner for a documented format policy. */
  lineHeightRatio?: number;
  /** Format-owned East-Asian single-line height divided by em, derived from
   * this resolved face's OpenType tables rather than its family name. */
  eastAsianLineHeightRatio?: number;
  /** Selected face's Canvas font box divided by em. This is raw resource
   * geometry; a format package may project it into its own documented line
   * allocation rule. Present only when the requested probe glyph is proven to
   * come from this exact face rather than browser fallback. */
  fontBoxRatio?: number;
  /** Authored Canvas tuple associated with this resource. */
  requestedFamily?: string;
  weight?: number;
  style?: 'normal' | 'italic';
  /** Canonical resource source. This is a route identity, not a claim that
   * native Canvas geometry is portable across engines or machines. */
  sourceIdentity?: string;
  /** Explicit UA synthesis policy. Production resource records are false;
   * deterministic test fixtures may opt into and label synthesis explicitly. */
  synthesized?: boolean;
}

export function normalizeFontMetricFamily(family: string): string {
  return family.trim().toLowerCase();
}

/** Geometry and identity retained for one concrete font resource. The family
 * name selects the resource; numeric metrics come only from that resource's
 * bytes or another explicit resource owner, never from a family-name table. */
export interface ResolvedFontMetric {
  /** Concrete FontFace family registered for Canvas measure and paint. */
  family: string;
  /** Admitted allocation class accompanying this geometry. A format may use
   * catalog classification when browser measurement cannot expose OS/2 bits;
   * this does not assert that the resource bytes were inspected. */
  farEastCodePage?: boolean;
  /** General single-line height divided by em, when provided by the resource
   * owner for a documented format policy. */
  lineHeightRatio?: number;
  /** Design ascent divided by em, derived from the selected resource's
   * OpenType tables. Kept separate so inline objects can share the same
   * baseline without guessing from a family name. */
  designAscentRatio?: number;
  /** Design descent below the baseline divided by em. */
  designDescentRatio?: number;
  /** Inter-line leading divided by em. Kept out of the baseline descent so
   * inline objects and paragraph marks are not shifted by line gap. */
  lineGapRatio?: number;
  /** The face's own normal line divided by em, unprojected: the ascent above
   * the baseline, the descent below it (without the gap) and the signed line
   * gap. The hhea values (OS/2 typo when USE_TYPO_METRICS is set) for a
   * parsed face, the measured CSS normal line for a host-measured one. All
   * three are present together; a format may project them into its own line
   * allocation rule (the design sides above are Word's projection). */
  normalAscentRatio?: number;
  normalDescentRatio?: number;
  normalLineGapRatio?: number;
  /** Positive OS/2 xAvgCharWidth / em for the selected face. DOCX may use it
   * for Word-observed adjustable Latin spaces; it is never a glyph advance. */
  averageCharWidthRatio?: number;
  /** Format-owned East-Asian single-line height divided by em, derived from
   * this resolved face's OpenType tables rather than its family name. */
  eastAsianLineHeightRatio?: number;
  /** A local() face has no readable cmap. Host-derived geometry requires a
   * native fallback probe; unavailable proof must not imply full coverage. */
  requiresGlyphCoverage?: boolean;
  /** Exact-source size-adjust:0% registration retained beside the paint face. */
  glyphCoverageProbeFamily?: string;
  /** Native route metric: the route's first available family (CSS syntax),
   * whose measured normal line this is. Ownership of a cluster is judged with
   * the route sentinels (native-route-line-metrics.ts). */
  nativeRoutePrimaryFamily?: string;
  /** Unicode scalar ranges mapped by this exact resource's supported cmap.
   * Resource metrics are applied only when every scalar in a shaped span is
   * covered, so a subset font cannot lend its geometry to fallback glyphs. */
  unicodeRanges?: readonly (readonly [start: number, end: number])[];
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

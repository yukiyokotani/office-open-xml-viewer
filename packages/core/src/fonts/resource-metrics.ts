/** Geometry and identity retained for one concrete font resource. The family
 * name selects the resource; numeric metrics come only from that resource's
 * bytes or another explicit resource owner, never from a family-name table. */
export interface ResolvedFontMetric {
  /** Concrete FontFace family registered for Canvas measure and paint. */
  family: string;
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
  /** Positive OS/2 xAvgCharWidth / em for the selected face. DOCX may use it
   * for Word-observed adjustable Latin spaces; it is never a glyph advance. */
  averageCharWidthRatio?: number;
  /** Format-owned East-Asian single-line height divided by em, derived from
   * this resolved face's OpenType tables rather than its family name. */
  eastAsianLineHeightRatio?: number;
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

/** Project one concrete OpenType face into renderer-neutral design metrics.
 * Identity and face selection stay with the resource loader; this function
 * deliberately accepts no family name. */
export function openTypeDesignLineRatios(metrics: Readonly<{
  unitsPerEm: number;
  hheaAscent: number;
  hheaDescent: number;
  hheaLineGap: number;
}>): Readonly<{
  lineHeightRatio: number;
  designAscentRatio: number;
  designDescentRatio: number;
  lineGapRatio: number;
}> | null {
  if (!(Number.isFinite(metrics.unitsPerEm) && metrics.unitsPerEm > 0)) return null;
  const ascent = Math.max(0, metrics.hheaAscent);
  const descent = Math.max(0, -metrics.hheaDescent);
  const lineGap = Math.max(0, metrics.hheaLineGap);
  if (!(ascent + descent + lineGap > 0)) return null;
  return Object.freeze({
    lineHeightRatio: (ascent + descent + lineGap) / metrics.unitsPerEm,
    designAscentRatio: ascent / metrics.unitsPerEm,
    designDescentRatio: descent / metrics.unitsPerEm,
    lineGapRatio: lineGap / metrics.unitsPerEm,
  });
}

export function normalizeFontMetricFamily(family: string): string {
  return family.trim().toLowerCase();
}

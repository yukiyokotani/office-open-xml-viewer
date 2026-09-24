/**
 * Office-observed automatic single-line allocation from one static OpenType
 * face. Word for Mac controls isolated OS/2 code-page bits 17–20 and found a
 * 1.3× hhea glyph box for that class; other faces use signed hhea lineGap above
 * the baseline. Independent Excel DrawingML controls with omitted <a:lnSpc>
 * matched the same projection for Meiryo UI Bold and Arial Bold at 12/25 pt and
 * top/centre/bottom anchors. ECMA-376 defines those spacing/anchor attributes,
 * but does not select the font tables. Callers must gate the rule to their own
 * tested format and to a known font-metric source.
 */
export const OFFICE_FAR_EAST_SINGLE_LINE_FACTOR = 1.3;

export function officeOpenTypeAutoLineRatios(metrics: Readonly<{
  unitsPerEm: number;
  hheaAscent: number;
  hheaDescent: number;
  hheaLineGap: number;
  farEastCodePage: boolean;
}>): Readonly<{
  lineHeightRatio: number;
  designAscentRatio: number;
  designDescentRatio: number;
}> | null {
  const { unitsPerEm, hheaAscent, hheaDescent, hheaLineGap, farEastCodePage } = metrics;
  if (!(Number.isFinite(unitsPerEm) && unitsPerEm > 0
    && Number.isFinite(hheaAscent) && hheaAscent >= 0
    && Number.isFinite(hheaDescent) && hheaDescent <= 0
    && Number.isFinite(hheaLineGap))) return null;
  const glyphBox = hheaAscent - hheaDescent;
  if (!(glyphBox > 0)) return null;
  const farEastHalfLeading = ((OFFICE_FAR_EAST_SINGLE_LINE_FACTOR - 1) / 2) * glyphBox;
  const ascent = farEastCodePage
    ? hheaAscent + farEastHalfLeading
    : hheaAscent + hheaLineGap;
  const descent = farEastCodePage
    ? -hheaDescent + farEastHalfLeading
    : -hheaDescent;
  if (!(ascent >= 0 && descent >= 0 && ascent + descent > 0)) return null;
  return Object.freeze({
    lineHeightRatio: (ascent + descent) / unitsPerEm,
    designAscentRatio: ascent / unitsPerEm,
    designDescentRatio: descent / unitsPerEm,
  });
}

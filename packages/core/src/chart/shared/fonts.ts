// Classic chart fonts helpers.
import type { ChartModel } from '../../types/chart';


// ─── Font-face resolution (CH10) ─────────────────────────────────────────────
// Chart text elements draw with, in priority order: the element's own
// `<a:latin typeface>` (from its `<c:txPr>`), else the theme font-scheme face
// (heading `majorFont` for titles, body `minorFont` for tick labels / data
// labels / legend, ECMA-376 §20.1.4.2), else the built-in `sans-serif`. When
// neither a per-element face nor a theme face is present the result is exactly
// `sans-serif`, so charts that specify no faces render byte-identically to
// before. A resolved face is quoted and given the same Calibri/Arial fallback
// chain as the chart title, so a font the platform lacks still degrades to a
// sans-serif rather than a serif default.
export type ChartFontRole = 'major' | 'minor';


/** Resolve a DrawingML theme font-scheme reference (`+mj-lt` / `+mn-lt` etc.,
 *  ECMA-376 §20.1.4.1.16) to the concrete theme face. `+mj-*` = heading
 *  (majorFont), `+mn-*` = body (minorFont); the axis suffix (`-lt`/`-ea`/`-cs`)
 *  is ignored here — chart text is Latin. A non-reference face passes through.
 *  Returns null when a reference can't be resolved (theme not threaded). */
export function resolveThemeFontRef(chart: ChartModel, face: string | null | undefined): string | null | undefined {
  if (!face) return face;
  if (face.startsWith('+mj')) return chart.themeMajorFontLatin ?? null;
  if (face.startsWith('+mn')) return chart.themeMinorFontLatin ?? null;
  return face;
}


export function chartFontFamily(
  chart: ChartModel,
  elementFace: string | null | undefined,
  role: ChartFontRole,
): string {
  const themeFace = role === 'major' ? chart.themeMajorFontLatin : chart.themeMinorFontLatin;
  const face = resolveThemeFontRef(chart, elementFace) ?? themeFace;
  return face ? `"${face}", Calibri, Arial, sans-serif` : 'sans-serif';
}


export function chartFontCss(
  fontSizePx: number,
  fontFamily: string,
  bold = false,
  italic = false,
): string {
  return `${italic ? 'italic ' : ''}${bold ? 'bold ' : ''}${fontSizePx}px ${fontFamily}`;
}

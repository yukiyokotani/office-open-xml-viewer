// Classic chart fonts helpers.
import type { ChartModel } from '../../types/chart';
import { fontFamilyKey } from '../../fonts/substitutions.js';


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

/** Fallback families appended after a resolved chart face.
 *
 *  Library policy, not a spec rule: ECMA-376 names the face (element
 *  `<a:latin>` or the theme font scheme, §20.1.4.1.16/§20.1.4.2) but leaves
 *  the substitute for an unavailable face to the consumer. Calibri is the body
 *  face of the Office 2007–2019 default theme and Arial a widely installed
 *  metric-neutral sans-serif, so a missing face degrades to a sans-serif rather
 *  than the platform's serif default. No Office measurement backs the order. */
export const CHART_FALLBACK_FONT_STACK = 'Calibri, Arial, sans-serif';

type ChartFontSubstitutions = Pick<ChartModel, 'fontSubstitutions'>;

/** `"face"`, followed by the family a font source loaded for it
 *  ({@link ChartModel.fontSubstitutions}); the authored face stays first, so an
 *  installed authored face still wins. */
/** The family a font source loaded for `face`, read as an own string property
 *  only: chart faces are document text, so a face such as `constructor` or
 *  `__proto__` must never resolve through the object prototype. */
export function chartFontSubstitute(chart: ChartFontSubstitutions | undefined, face: string): string | undefined {
  const map = chart?.fontSubstitutions;
  if (!map) return undefined;
  const key = fontFamilyKey(face);
  if (!Object.prototype.hasOwnProperty.call(map, key)) return undefined;
  const substitute = (map as Record<string, unknown>)[key];
  return typeof substitute === 'string' && substitute ? substitute : undefined;
}

export function chartFaceList(face: string, chart?: ChartFontSubstitutions): string {
  const substitute = chartFontSubstitute(chart, face);
  return substitute ? `"${face}", "${substitute.replace(/["\\]/g, '')}"` : `"${face}"`;
}

/** {@link CHART_FALLBACK_FONT_STACK} with, after each of its named families,
 *  the family a font source loaded for it. */
export function chartFallbackFontStack(chart?: ChartFontSubstitutions): string {
  if (!chart?.fontSubstitutions) return CHART_FALLBACK_FONT_STACK;
  const named = (family: string) => {
    const substitute = chartFontSubstitute(chart, family);
    return substitute ? `${family}, "${substitute.replace(/["\\]/g, '')}"` : family;
  };
  return `${named('Calibri')}, ${named('Arial')}, sans-serif`;
}

/** CSS family list for a resolved chart face plus {@link CHART_FALLBACK_FONT_STACK}. */
export function chartFaceWithFallback(face: string, chart?: ChartFontSubstitutions): string {
  return `${chartFaceList(face, chart)}, ${chartFallbackFontStack(chart)}`;
}


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
  return face ? chartFaceWithFallback(face, chart) : 'sans-serif';
}


export function chartFontCss(
  fontSizePx: number,
  fontFamily: string,
  bold = false,
  italic = false,
): string {
  return `${italic ? 'italic ' : ''}${bold ? 'bold ' : ''}${fontSizePx}px ${fontFamily}`;
}

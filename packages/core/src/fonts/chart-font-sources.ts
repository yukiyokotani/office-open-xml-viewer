/**
 * Font sources (`LoadOptions.fonts`) for DrawingML charts, shared by DOCX,
 * XLSX and PPTX. Library resource policy, not OOXML behaviour.
 *
 * - Demand: the families a chart paints with, from one generic walk over the
 *   canonical {@link ChartModel}, so a face named only by a chart is demanded
 *   too.
 * - Paint: charts paint the faces a font source actually registered for the
 *   document, under its private family names, through
 *   {@link ChartModel.fontSubstitutions}. Chart labels are not split by
 *   script, so script-scoped substitutes are left out.
 */
import type { ChartModel } from '../types/chart.js';
import type { FontSubstituteScript } from './substitute-script.js';
import { fontFamilyKey } from './substitutions.js';

/** Keys of a chart model whose string value names a text face: every
 * `…FontFace` / `fontFace` (title, axes, axis titles, data labels, legend,
 * legend entries, rich runs, secondary axes) and the theme faces charts
 * default to (ECMA-376 §21.2.2.216 txPr, §20.1.4.2 font scheme). */
const CHART_FACE_KEY = /^(?:fontFace|[A-Za-z]+FontFace|themeMajorFontLatin|themeMinorFontLatin)$/;

/** The families a chart paints with, from one generic walk over its model.
 * Theme references (`+mj-*` / `+mn-*`, §20.1.4.1.16) resolve to the chart's
 * theme faces. The library fallback faces of chart stacks (Calibri, Arial)
 * are not demanded: they are fallbacks, painted with a source face only when
 * one loaded for them. */
export function chartFontFamilies(chart: unknown): string[] {
  const families = new Set<string>();
  const model = chart as { themeMajorFontLatin?: string | null; themeMinorFontLatin?: string | null } | null;
  const pending: unknown[] = [chart];
  const seen = new Set<object>();
  while (pending.length) {
    const value = pending.pop();
    if (value === null || typeof value !== 'object' || seen.has(value) || ArrayBuffer.isView(value)) continue;
    seen.add(value);
    for (const [key, child] of Object.entries(value)) {
      if (typeof child === 'string') {
        if (!CHART_FACE_KEY.test(key)) continue;
        const face = child.startsWith('+mj') ? model?.themeMajorFontLatin
          : child.startsWith('+mn') ? model?.themeMinorFontLatin : child;
        const name = face?.trim();
        if (name && !name.startsWith('+')) families.add(name);
      } else if (child !== null && typeof child === 'object') {
        pending.push(child);
      }
    }
  }
  return [...families];
}

/** What charts read of a font-source snapshot (`LoadedFontSources`):
 * lower-cased authored family → served family, and lower-cased served family
 * → the private family its faces are registered under. */
export interface LoadedChartFontSources {
  readonly substitutions: ReadonlyMap<string, { readonly family: string; readonly script?: FontSubstituteScript }>;
  readonly families: ReadonlyMap<string, string>;
}

const recordBySnapshot = new WeakMap<LoadedChartFontSources, Readonly<Record<string, string>> | null>();

/** One loaded snapshot as the {@link ChartModel.fontSubstitutions} record:
 * lower-cased chart face → the registered family a chart stack names after
 * it. A face a source served directly maps to its own private family; a face
 * with an unscoped loaded substitution maps to the substitute's private
 * family. `undefined` when there are none. The record has a null prototype
 * (keys are document font names, e.g. `__proto__`) and is memoised per
 * snapshot (snapshots are immutable). */
export function chartFontSubstitutions(
  sources: LoadedChartFontSources | undefined,
): Readonly<Record<string, string>> | undefined {
  if (!sources?.families.size) return undefined;
  let record = recordBySnapshot.get(sources);
  if (record === undefined) {
    const entries = Object.create(null) as Record<string, string>;
    for (const [family, alias] of sources.families) entries[family] = alias;
    for (const [family, substitution] of sources.substitutions) {
      const alias = substitution.script ? undefined : sources.families.get(fontFamilyKey(substitution.family));
      if (alias) entries[family] = alias;
    }
    record = Object.keys(entries).length ? Object.freeze(entries) : null;
    recordBySnapshot.set(sources, record);
  }
  return record ?? undefined;
}

/** The chart model to paint: `chart` with `fontSubstitutions`, or `chart`
 * unchanged without any (no font source, nothing loaded). */
export function withChartFontSubstitutions<T extends ChartModel>(
  chart: T,
  substitutions: Readonly<Record<string, string>> | undefined,
): T {
  return substitutions ? { ...chart, fontSubstitutions: substitutions } : chart;
}

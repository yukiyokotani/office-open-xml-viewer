import type { CjkLang, FontSubstituteScript } from '@silurus/ooxml-core';
import {
  classifyCjkFont,
  cjkFallbackForText,
  scriptPreloadNamesForText,
  substituteEntryCoversText,
  symbolFontToUnicode,
  type OfficeFontFallbackRequest,
} from '@silurus/ooxml-core';
import { formatAutoNum } from './autonum-label.js';
import { ScriptPreloadAccumulator } from '@silurus/ooxml-core/internal/script-preload-accumulator';
import {
  FONT_SUBSTITUTE_SCRIPTS,
  fontFamilyKey,
} from '@silurus/ooxml-core/internal/font-sources';
import { chartFontFamilies } from '@silurus/ooxml-core/internal/chart-font-sources';
import type { Slide, SlideElement, TextBody } from './types';
import { officeFaceWeight } from './powerpoint-line-metrics.js';

// PPTX font requests: the families a deck paints (`fontPreloadNames`), the
// exact local Office-face tuples one slide probes, and, when font sources are
// configured (`LoadOptions.fonts`), the remaining facts of the font-source
// demand. Library resource policy: DrawingML names faces (ECMA-376
// §21.1.2.3) and leaves their provision to the host. Demand is computed from
// the slide model alone; which families a source serves is the source's
// business, so no provider table appears here.

/** Font-source demand facts beyond the names, collected only when font
 * sources are configured (the default offline path neither computes nor
 * retains them). */
export interface PptxFontDemandFacts {
  /** The leading `fontPreloadNames` entries that are authored families (theme
   * major/minor and every run, paragraph-default and slot family); the rest
   * are library script fallbacks. */
  readonly familyCount: number;
  /** Lower-cased authored family → the scoped-substitute scripts its painted
   * text proves (`FontSessionDemand.substituteScripts`). Families absent here prove
   * none, so a script-scoped substitute is never requested for them. */
  readonly substituteScripts: Readonly<Record<string, readonly FontSubstituteScript[]>>;
}

/** Include authored equation previews in font preparation, so runtime recovery
 * uses the same fonts in main-thread and worker rendering. Iterators retain
 * only the branch stack rather than flattening a second slide model. */
function* fontElements(elements: readonly SlideElement[]): Generator<SlideElement> {
  const pending: Iterator<SlideElement>[] = [elements[Symbol.iterator]()];
  while (pending.length) {
    const next = pending[pending.length - 1].next();
    if (next.done) { pending.pop(); continue; }
    const element = next.value;
    yield element;
    if ((element.type === 'shape' || element.type === 'table') && element.mathFallback) {
      pending.push(element.mathFallback.elements[Symbol.iterator]());
    }
  }
}

/** Yield every painted text string in a text body (paragraph runs). */
function* textBodyRuns(body: TextBody | null | undefined): Generator<string> {
  for (const p of body?.paragraphs ?? []) {
    for (const fallback of p.mathFallbacks ?? []) yield fallback.text;
    for (const r of p.runs) {
      if (r.type === 'text') yield r.text;
      if (r.type === 'math' && r.fallbackText) yield r.fallbackText;
    }
  }
}

/** Yield every explicitly resolved family carried by one rendered text body. */
function* textBodyFontFamilies(body: TextBody | null | undefined): Generator<string> {
  for (const paragraph of body?.paragraphs ?? []) {
    if (paragraph.defFontFamily) yield paragraph.defFontFamily;
    for (const run of paragraph.runs) {
      if (run.type !== 'text') continue;
      if (run.fontFamily) yield run.fontFamily;
      if (run.fontFamilyEa) yield run.fontFamilyEa;
      // ECMA-376 §21.1.2.3.1: a:cs names an independently resolved run face.
      // The existing renderer consumes it for complex-script segments;
      // preload must retain that independently authored face.
      if (run.fontFamilyCs) yield run.fontFamilyCs;
      if (run.fontFamilySym) yield run.fontFamilySym;
    }
  }
}

/** Yield every rendered text string in the presentation: shape text, table
 *  cell text and chart labels across all slides. Speaker notes and comments are
 *  not painted on the slide, so they are excluded (the renderer ignores them). */
export function* pptxSlideTextRuns(slide: Slide): Generator<string> {
  for (const el of fontElements(slide.elements)) {
    if (el.type === 'shape') {
      yield* textBodyRuns(el.textBody);
    } else if (el.type === 'table') {
      for (const row of el.rows) {
        for (const cell of row.cells) yield* textBodyRuns(cell.textBody);
      }
    } else if (el.type === 'chart') {
      if (el.chart.title) yield el.chart.title;
      for (const c of el.chart.categories) yield c;
      for (const s of el.chart.series) if (s.name) yield s.name;
    }
  }
}

/** Mutable font-source demand facts of an accumulator (see
 * {@link PptxFontDemandFacts}). */
interface DemandFactState {
  readonly scripts: Map<string, Set<FontSubstituteScript>>;
}

function cloneDemandFacts(state: DemandFactState): DemandFactState {
  return {
    scripts: new Map([...state.scripts].map(([key, scripts]) => [key, new Set(scripts)])),
  };
}

/** Incremental PPTX adapter over core's single script-classification source. */
export class PptxFontPreloadAccumulator {
  private readonly scripts: ScriptPreloadAccumulator;
  private readonly families: Set<string>;
  private readonly demand: DemandFactState | null;

  constructor(
    private readonly majorFont: string | null,
    private readonly minorFont: string | null,
    scripts?: ScriptPreloadAccumulator,
    families?: Set<string>,
    private readonly fallback?: CjkLang,
    private readonly scriptNames = new Set<string>(),
    /** Collect {@link PptxFontDemandFacts} (font sources configured). */
    demand: boolean | DemandFactState = false,
  ) {
    const cjkLang = classifyCjkFont(majorFont) ?? classifyCjkFont(minorFont) ?? fallback ?? null;
    this.scripts = scripts ?? new ScriptPreloadAccumulator(cjkLang);
    this.families = families ?? new Set();
    this.demand = demand === false ? null : demand === true ? { scripts: new Map() } : demand;
    if (majorFont) this.addThemeFamily(majorFont);
    if (minorFont) this.addThemeFamily(minorFont);
  }

  /** Theme fonts govern every run that names no face of its own, whatever its
   * text, so every scoped-substitute script is admitted for them. */
  private addThemeFamily(family: string): void {
    this.families.add(family);
    if (this.demand) this.demand.scripts.set(fontFamilyKey(family), new Set(FONT_SUBSTITUTE_SCRIPTS));
  }

  addSlide(slide: Slide): void {
    if (slide.theme?.majorFont) this.addThemeFamily(slide.theme.majorFont);
    if (slide.theme?.minorFont) this.addThemeFamily(slide.theme.minorFont);
    this.scripts.addText(pptxSlideTextRuns(slide));
    // Keep the union of per-slide choices: later kana must not remove a Han-only
    // slide's SC preload after that slide has already been published.
    for (const name of scriptPreloadNamesForText(pptxSlideTextRuns(slide),
      classifyCjkFont(slide.theme ? slide.theme.majorFont : this.majorFont)
        ?? classifyCjkFont(slide.theme ? slide.theme.minorFont : this.minorFont) ?? this.fallback ?? null)) {
      this.scriptNames.add(name);
    }
    const body = (textBody: TextBody | null | undefined) => {
      for (const family of textBodyFontFamilies(textBody)) this.families.add(family);
      if (this.demand) this.proveSubstituteScripts(textBody);
    };
    for (const el of fontElements(slide.elements)) {
      if (el.type === 'shape') {
        body(el.textBody);
      } else if (el.type === 'table') {
        for (const row of el.rows) {
          for (const cell of row.cells) body(cell.textBody);
        }
      }
    }
    if (this.demand) {
      const major = slide.theme ? slide.theme.majorFont : this.majorFont;
      const minor = slide.theme ? slide.theme.minorFont : this.minorFont;
      for (const request of pptxSlideOfficeFontRequests(slide, major, minor)) {
        // Every face a slide's text can paint with is demanded, not only the
        // run families above: bullet faces (a:buFont), line-break and
        // end-of-paragraph properties, and theme tokens resolved to the
        // slide's theme fonts.
        this.families.add(request.family);
      }
      for (const el of slide.elements) {
        if (el.type === 'chart') for (const family of chartFontFamilies(el.chart)) this.families.add(family);
      }
    }
  }

  /** Record, per family a run names, the scoped-substitute scripts the run's
   * text proves (core substitute-script.ts). A CSS stack cannot split a run, so
   * every slot family of the run is credited with its text, exactly as the
   * renderer's `cssFontStack` admits a scoped substitute for the whole run. */
  private proveSubstituteScripts(textBody: TextBody | null | undefined): void {
    const scripts = (this.demand as DemandFactState).scripts;
    const prove = (text: string, families: readonly (string | null | undefined)[]) => {
      if (!text) return;
      for (const script of FONT_SUBSTITUTE_SCRIPTS) {
        if (!substituteEntryCoversText({ script }, text, 'any')) continue;
        for (const family of families) {
          // The renderer's family resolution: a ",<generic>" hint is dropped
          // (theme tokens resolve to theme fonts, which admit every script).
          const name = family?.split(',')[0]?.trim();
          if (!name) continue;
          const key = fontFamilyKey(name);
          scripts.set(key, (scripts.get(key) ?? new Set()).add(script));
        }
      }
    };
    for (const paragraph of textBody?.paragraphs ?? []) {
      for (const run of paragraph.runs) {
        if (run.type !== 'text') continue;
        prove(run.text, [run.fontFamily ?? paragraph.defFontFamily, run.fontFamilyEa, run.fontFamilyCs, run.fontFamilySym]);
      }
      // Synthesized marker text paints with the marker's resolved family
      // (renderer paragraphBullet, §21.1.2.4): a character bullet with its
      // a:buFont (a theme font without one), unless its symbol code was mapped
      // to a Unicode symbol painted in the generic sans-serif; an auto-number
      // label with its a:buFont, else the first run's family. Field text is
      // the run text above; line breaks and paragraph marks paint no glyph.
      const bullet = paragraph.bullet;
      if (bullet.type === 'char' && bullet.char
        && symbolFontToUnicode(bullet.char, bullet.fontFamily ?? null) === bullet.char) {
        prove(bullet.char, [bullet.fontFamily]);
      } else if (bullet.type === 'autoNum') {
        const firstRunFamily = paragraph.runs.find((run) => run.type === 'text' && !!run.fontFamily);
        // Each scheme labels every counter in one script.
        prove(formatAutoNum(bullet.startAt ?? 1, bullet.numType), [
          bullet.fontFamily ?? (firstRunFamily?.type === 'text' ? firstRunFamily.fontFamily : null) ?? paragraph.defFontFamily,
        ]);
      }
    }
  }

  names(): (string | null)[] {
    return [...new Set([...this.families, ...this.scripts.names(), ...this.scriptNames])];
  }

  /** {@link PptxFontDemandFacts} for {@link names}, or undefined when this
   * accumulator does not collect them. */
  demandFacts(): PptxFontDemandFacts | undefined {
    if (!this.demand) return undefined;
    // Null prototype: keys are document font names (e.g. `__proto__`).
    const substituteScripts = Object.create(null) as Record<string, readonly FontSubstituteScript[]>;
    for (const [key, scripts] of this.demand.scripts) {
      if (scripts.size) substituteScripts[key] = Object.freeze([...scripts].sort());
    }
    return Object.freeze({
      // names() lists the distinct families first, so they are its prefix.
      familyCount: this.families.size,
      substituteScripts: Object.freeze(substituteScripts),
    });
  }

  withSlide(slide: Slide): PptxFontPreloadAccumulator {
    const candidate = new PptxFontPreloadAccumulator(
      this.majorFont,
      this.minorFont,
      this.scripts.clone(),
      new Set(this.families),
      this.fallback,
      new Set(this.scriptNames),
      this.demand ? cloneDemandFacts(this.demand) : false,
    );
    candidate.addSlide(slide);
    return candidate;
  }
}

const CSS_GENERIC_FAMILY_NAMES = new Set(['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui']);

/** Authored face tuples actually painted by this slide. DrawingML run properties
 * inherit through paragraph/body defaults (§21.1.2.3.9); theme tokens resolve
 * against the same major/minor values used by the renderer. Loading is delayed
 * until the slide is requested, so unused slides and styles fetch no assets. */
export function pptxSlideOfficeFontRequests(
  slide: Slide,
  majorFont: string | null,
  minorFont: string | null,
): OfficeFontFallbackRequest[] {
  if (slide.theme) { majorFont = slide.theme.majorFont; minorFont = slide.theme.minorFont; }
  const requests = new Map<string, OfficeFontFallbackRequest>();
  const resolved = (family: string | null | undefined): string | null => {
    // A missing family has no authored or inherited font resource. The theme
    // minor face is the renderer's CSS fallback, but it must not be treated as
    // a document request for an exact Office face. Explicit +mn-* still is.
    if (!family) return null;
    if (family.startsWith('+mn-')) return minorFont;
    if (family.startsWith('+mj-')) return majorFont;
    return family.split(',')[0]?.trim() ?? null;
  };
  const add = (family: string | null | undefined, bold: boolean, italic: boolean) => {
    // Every authored face is probed, so layout can admit catalog geometry
    // only for faces proven present (no family-name gate).
    const name = resolved(family);
    if (!name || CSS_GENERIC_FAMILY_NAMES.has(name.toLowerCase())) return;
    const weight = officeFaceWeight(name, bold, italic);
    const style = italic ? 'italic' : 'normal';
    requests.set(`${name.toLowerCase()}:${weight}:${style}`, { family: name, weight, style });
    // Family presence is probed independently of the styled tuples a slide
    // uses, so an installed regular-only family stays present for bold runs.
    requests.set(`presence:${name.toLowerCase()}`, { family: name, presenceOnly: true });
  };
  const body = (textBody: TextBody | null | undefined) => {
    for (const paragraph of textBody?.paragraphs ?? []) {
      for (const run of paragraph.runs) {
        if (run.type === 'break') {
          // a:br carries its own character properties (CT_TextLineBreak); the
          // renderer sizes the break's line with that face.
          add(run.fontFamily ?? paragraph.defFontFamily,
            run.bold ?? paragraph.defBold ?? textBody?.defaultBold ?? false,
            run.italic ?? paragraph.defItalic ?? textBody?.defaultItalic ?? false);
          continue;
        }
        if (run.type !== 'text') continue;
        const bold = run.bold ?? paragraph.defBold ?? textBody?.defaultBold ?? false;
        const italic = run.italic ?? paragraph.defItalic ?? textBody?.defaultItalic ?? false;
        add(run.fontFamily ?? paragraph.defFontFamily, bold, italic);
        if (run.fontFamilyEa) add(run.fontFamilyEa, bold, italic);
        if (run.fontFamilyCs) add(run.fontFamilyCs, bold, italic);
        if (run.fontFamilySym) add(run.fontFamilySym, bold, italic);
      }
      // a:endParaRPr: the end mark contributes its face to the last line.
      const end = paragraph.endRunProperties;
      if (end) {
        add(end.fontFamily ?? paragraph.defFontFamily,
          end.bold ?? paragraph.defBold ?? textBody?.defaultBold ?? false,
          end.italic ?? paragraph.defItalic ?? textBody?.defaultItalic ?? false);
      }
      if (paragraph.bullet.type === 'char' || paragraph.bullet.type === 'autoNum') {
        // The marker renderer uses normal weight/style for both bullet kinds.
        // autoNum inherits the first text run's family when buFont is absent.
        const firstRun = paragraph.runs.find((run) => run.type === 'text' && !!run.fontFamily);
        const firstRunFamily = firstRun?.type === 'text' ? firstRun.fontFamily : null;
        add(paragraph.bullet.fontFamily ?? (paragraph.bullet.type === 'autoNum'
          ? firstRunFamily ?? paragraph.defFontFamily
          : undefined), false, false);
      }
    }
  };
  for (const element of fontElements(slide.elements)) {
    if (element.type === 'shape') body(element.textBody);
    if (element.type === 'table') {
      for (const row of element.rows) for (const cell of row.cells) body(cell.textBody);
    }
  }
  return [...requests.values()];
}


export function pptxSlideCjkFallback(slide: Slide, major: string | null, minor: string | null, fallback: CjkLang): CjkLang {
  return cjkFallbackForText(pptxSlideTextRuns(slide),
    classifyCjkFont(major) ?? classifyCjkFont(minor) ?? fallback);
}

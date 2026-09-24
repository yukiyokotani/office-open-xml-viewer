import type { CjkLang } from '@silurus/ooxml-core';
import {
  classifyCjkFont,
  cjkFallbackForText,
  scriptPreloadNamesForText,
  GOOGLE_FONT_SUBSTITUTES,
  SCRIPT_GOOGLE_FONTS,
  type FontPreloadEntry,
  type TextBody,
  type OfficeFontFallbackRequest,
} from '@silurus/ooxml-core';
import { ScriptPreloadAccumulator } from '@silurus/ooxml-core/internal/script-preload-accumulator';
import type { Presentation, Slide, SlideElement } from './types';

/** Theme-referenced typefaces commonly used by PPTX templates. Keys are
 *  lower-cased family names.
 *
 *  {@link GOOGLE_FONT_SUBSTITUTES} supplies base text-face substitutes
 *  (Calibri → Carlito, Cambria → Caladea), the popular free
 *  web fonts and the Arabic Noto fallbacks — shared with docx/xlsx; the
 *  renderer puts each substitute into the canvas font stack so a missing Office
 *  font degrades to an advance-width alternative instead of a generic system serif/sans.
 *  {@link SCRIPT_GOOGLE_FONTS} adds the CJK / Cyrillic / Thai / Devanagari /
 *  Hebrew Noto faces (CJK ordered by document language). Both load only when
 *  `useGoogleFonts` is on. The exact local Calibri route is applied only to
 *  named font choices; the CSS alias for an unavailable face is likewise
 *  enabled only by that explicit opt-in. PPTX has no format-specific additions. */
export const PPTX_GOOGLE_FONTS: Record<string, FontPreloadEntry> = {
  ...GOOGLE_FONT_SUBSTITUTES,
  ...SCRIPT_GOOGLE_FONTS,
};

/** Yield every painted text string in a text body (paragraph runs). */
function* textBodyRuns(body: TextBody | null | undefined): Generator<string> {
  for (const p of body?.paragraphs ?? []) {
    for (const r of p.runs) {
      if (r.type === 'text') yield r.text;
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
      if (run.fontFamilySym) yield run.fontFamilySym;
    }
  }
}

/** Yield every rendered text string in the presentation: shape text, table
 *  cell text and chart labels across all slides. Speaker notes and comments are
 *  not painted on the slide, so they are excluded (the renderer ignores them). */
export function* pptxSlideTextRuns(slide: Slide): Generator<string> {
  for (const el of slide.elements as SlideElement[]) {
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

/** Incremental PPTX adapter over core's single script-classification source. */
export class PptxFontPreloadAccumulator {
  private readonly scripts: ScriptPreloadAccumulator;
  private readonly families: Set<string>;

  constructor(
    private readonly majorFont: string | null,
    private readonly minorFont: string | null,
    scripts?: ScriptPreloadAccumulator,
    families?: Set<string>,
    private readonly fallback?: CjkLang,
    private readonly scriptNames = new Set<string>(),
  ) {
    const cjkLang = classifyCjkFont(majorFont) ?? classifyCjkFont(minorFont) ?? fallback ?? null;
    this.scripts = scripts ?? new ScriptPreloadAccumulator(cjkLang);
    this.families = families ?? new Set();
    if (majorFont) this.families.add(majorFont);
    if (minorFont) this.families.add(minorFont);
  }

  addSlide(slide: Slide): void {
    this.scripts.addText(pptxSlideTextRuns(slide));
    // Keep the union of per-slide choices: later kana must not remove a Han-only
    // slide's SC preload after that slide has already been published.
    for (const name of scriptPreloadNamesForText(pptxSlideTextRuns(slide),
      classifyCjkFont(this.majorFont) ?? classifyCjkFont(this.minorFont) ?? this.fallback ?? null)) {
      this.scriptNames.add(name);
    }
    for (const el of slide.elements as SlideElement[]) {
      if (el.type === 'shape') {
        for (const family of textBodyFontFamilies(el.textBody)) this.families.add(family);
      } else if (el.type === 'table') {
        for (const row of el.rows) {
          for (const cell of row.cells) {
            for (const family of textBodyFontFamilies(cell.textBody)) this.families.add(family);
          }
        }
      }
    }
  }

  names(): (string | null)[] {
    return [...new Set([...this.families, ...this.scripts.names(), ...this.scriptNames])];
  }

  withSlide(slide: Slide): PptxFontPreloadAccumulator {
    const candidate = new PptxFontPreloadAccumulator(
      this.majorFont,
      this.minorFont,
      this.scripts.clone(),
      new Set(this.families),
      this.fallback,
      new Set(this.scriptNames),
    );
    candidate.addSlide(slide);
    return candidate;
  }
}

/**
 * The font-family names to preload for a presentation: the theme major/minor
 * fonts, plus only the script-fallback Noto faces whose script the slide TEXT
 * actually contains ({@link PptxFontPreloadAccumulator}). The renderer's canvas
 * font stack still ends with the full Noto set, but eagerly fetching the
 * multi-MB CJK families for a deck with no CJK glyphs would block first paint
 * for nothing; an un-preloaded face loads lazily if it ever proves needed.
 *
 * Single source of truth shared by the main-thread `load()` and the render
 * worker. Both derive the set from the SAME parsed {@link Presentation}, so both
 * modes preload an identical set — worker/main rendering must stay
 * pixel-equivalent.
 */
export function pptxFontPreloadNames(
  pres: Presentation,
  fallback?: CjkLang,
): (string | null | undefined)[] {
  const accumulator = new PptxFontPreloadAccumulator(
    pres.majorFont,
    pres.minorFont,
    undefined, undefined, fallback,
  );
  for (const slide of pres.slides) accumulator.addSlide(slide);
  return accumulator.names();
}

/** Calibri tuples actually painted by this slide. DrawingML run properties
 * inherit through paragraph/body defaults (§21.1.2.3.9); theme tokens resolve
 * against the same major/minor values used by the renderer. Loading is delayed
 * until the slide is requested, so unused slides and styles fetch no assets. */
export function pptxSlideOfficeFontRequests(
  slide: Slide,
  majorFont: string | null,
  minorFont: string | null,
): OfficeFontFallbackRequest[] {
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
    if (resolved(family)?.toLowerCase() !== 'calibri') return;
    const weight = bold ? 700 : 400;
    const style = italic ? 'italic' : 'normal';
    requests.set(`${weight}:${style}`, { family: 'Calibri', weight, style });
  };
  const body = (textBody: TextBody | null | undefined) => {
    for (const paragraph of textBody?.paragraphs ?? []) {
      for (const run of paragraph.runs) {
        if (run.type !== 'text') continue;
        const bold = run.bold ?? paragraph.defBold ?? textBody?.defaultBold ?? false;
        const italic = run.italic ?? paragraph.defItalic ?? textBody?.defaultItalic ?? false;
        add(run.fontFamily ?? paragraph.defFontFamily, bold, italic);
        if (run.fontFamilyEa) add(run.fontFamilyEa, bold, italic);
        if (run.fontFamilySym) add(run.fontFamilySym, bold, italic);
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
  for (const element of slide.elements) {
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

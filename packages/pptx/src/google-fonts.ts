import {
  classifyCjkFont,
  chartFontFamilies,
  GOOGLE_FONT_SUBSTITUTES,
  SCRIPT_GOOGLE_FONTS,
  type FontPreloadEntry,
  type TextBody,
} from '@silurus/ooxml-core';
import { ScriptPreloadAccumulator } from '@silurus/ooxml-core/internal/script-preload-accumulator';
import type { Presentation, Slide, SlideElement } from './types';

/** Theme-referenced typefaces commonly used by PPTX templates. Keys are
 *  lower-cased family names.
 *
 *  {@link GOOGLE_FONT_SUBSTITUTES} supplies the Office substitutes (Calibri /
 *  Calibri Light → Carlito, Cambria / Cambria Math → Caladea), the popular free
 *  web fonts and the Arabic Noto fallbacks — shared with docx/xlsx; the
 *  renderer puts each substitute into the canvas font stack so a missing Office
 *  font degrades to a same-width webfont instead of a wider system serif/sans.
 *  {@link SCRIPT_GOOGLE_FONTS} adds the CJK / Cyrillic / Thai / Devanagari /
 *  Hebrew Noto faces (CJK ordered by document language). Both load only when
 *  `useGoogleFonts` is on — no binaries ship in the bundle. PPTX currently has
 *  no format-specific additions. */
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
  ) {
    const cjkLang = classifyCjkFont(majorFont) ?? classifyCjkFont(minorFont) ?? null;
    this.scripts = scripts ?? new ScriptPreloadAccumulator(cjkLang);
    this.families = families ?? new Set();
    if (majorFont) this.families.add(majorFont);
    if (minorFont) this.families.add(minorFont);
  }

  addSlide(slide: Slide): void {
    this.scripts.addText(pptxSlideTextRuns(slide));
    for (const el of slide.elements as SlideElement[]) {
      if (el.type === 'shape') {
        for (const family of textBodyFontFamilies(el.textBody)) this.families.add(family);
      } else if (el.type === 'table') {
        for (const row of el.rows) {
          for (const cell of row.cells) {
            for (const family of textBodyFontFamilies(cell.textBody)) this.families.add(family);
          }
        }
      } else if (el.type === 'chart') {
        for (const family of chartFontFamilies(el.chart)) this.families.add(family);
      }
    }
  }

  names(): (string | null)[] {
    return [...this.families, ...this.scripts.names()];
  }

  providerNames(): string[] {
    return [...this.families];
  }

  withSlide(slide: Slide): PptxFontPreloadAccumulator {
    const candidate = new PptxFontPreloadAccumulator(
      this.majorFont,
      this.minorFont,
      this.scripts.clone(),
      new Set(this.families),
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
): (string | null | undefined)[] {
  const accumulator = new PptxFontPreloadAccumulator(
    pres.majorFont,
    pres.minorFont,
  );
  for (const slide of pres.slides) accumulator.addSlide(slide);
  return accumulator.names();
}

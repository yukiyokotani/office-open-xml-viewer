import { ScriptPreloadAccumulator } from '@silurus/ooxml-core/internal/script-preload-accumulator';
import type { CjkLang, OfficeFontFallbackRequest } from '@silurus/ooxml-core';
import {
  classifyCjkFont,
  cjkLangFromLanguage,
  scriptPreloadNamesForText,
  GOOGLE_FONT_SUBSTITUTES,
  SCRIPT_GOOGLE_FONTS,
  type FontPreloadEntry,
} from '@silurus/ooxml-core';
import type {
  DocxDocumentModel,
} from './types.js';
import { docxRenderedTextUsages } from './document-content.js';

/** Theme-referenced typefaces commonly used by DOCX templates.
 *
 *  {@link GOOGLE_FONT_SUBSTITUTES} supplies advance-width substitutes for the
 *  base Office text faces (Calibri → Carlito, Cambria → Caladea), popular free
 *  web fonts and the Arabic Noto
 *  fallbacks — shared with pptx/xlsx. {@link SCRIPT_GOOGLE_FONTS} adds the
 *  CJK (KR/SC/TC/JP, plus HK sans) / Cyrillic / Thai / Devanagari / Hebrew
 *  Noto faces the renderer appends to the font chain. CJK fallbacks are ordered
 *  by document language. Both load only when `useGoogleFonts` is on — no binaries
 *  ship in the bundle. DOCX
 *  currently has no format-specific additions. */
export const DOCX_GOOGLE_FONTS: Record<string, FontPreloadEntry> = {
  ...GOOGLE_FONT_SUBSTITUTES,
  ...SCRIPT_GOOGLE_FONTS,
};

function* docxTextRuns(doc: DocxDocumentModel): Generator<string> {
  for (const usage of docxRenderedTextUsages(doc)) yield usage.text;
}

/**
 * The font-family names to preload for a document: the theme major/minor fonts,
 * plus only the script-fallback Noto faces whose script the document's TEXT
 * actually contains ({@link scriptPreloadNamesForText}). The renderer's font
 * fallback chains still END with the full Noto set, but eagerly fetching the
 * multi-MB CJK families for a document that has no CJK glyphs would block first
 * paint for nothing; an un-preloaded face loads lazily if it ever proves needed.
 *
 * Single source of truth shared by the main-thread `load()` and the render
 * worker. Both derive the set from the SAME parsed {@link DocxDocumentModel}, so
 * they preload an identical set — worker/main rendering must stay
 * pixel-equivalent. (Fonts must also be loaded before pagination, which measures
 * text; both callers await this before paginating.)
 */
export function docxFontPreloadNames(
  doc: DocxDocumentModel,
  fallback?: CjkLang,
): (string | null | undefined)[] {
  const cjkLang =
    classifyCjkFont(doc.majorFont) ?? classifyCjkFont(doc.minorFont) ?? fallback ?? null;
  const scripts = new ScriptPreloadAccumulator(cjkLang);
  const languageNames = new Set<string>();
  for (const usage of docxRenderedTextUsages(doc)) {
    const region = cjkLangFromLanguage(usage.eastAsiaLanguage);
    if (region) {
      for (const name of scriptPreloadNamesForText([usage.text], region, true)) languageNames.add(name);
    } else {
      scripts.addText([usage.text]);
    }
  }
  return [doc.majorFont, doc.minorFont, ...new Set([...scripts.names(), ...languageNames])];
}

/** Probe exact local style tuples used by rendered text. The shared loader
 * declines uncatalogued names, and the document font table alone never queues
 * a face that rendered content does not use. No font bytes are packaged. */
export function docxOfficeFontFallbackRequests(
  doc: DocxDocumentModel,
): OfficeFontFallbackRequest[] {
  const tuples = new Map<string, OfficeFontFallbackRequest>();
  const add = (family: string | null | undefined, bold = false, italic = false) => {
    const name = family?.trim();
    if (!name) return;
    const weight = bold ? 700 : 400;
    const style = italic ? 'italic' : 'normal';
    tuples.set(`${name.toLocaleLowerCase('en-US')}:${weight}:${style}`, { family: name, weight, style });
  };
  // Theme faces can govern paragraph marks and inherited runs even when a
  // particular authored run does not repeat its font name.
  add(doc.majorFont);
  add(doc.minorFont);
  for (const usage of docxRenderedTextUsages(doc)) {
    for (const family of usage.fontFamilies) add(family, usage.bold, usage.italic);
    if (usage.text && (usage.latinFontFamily === null ||
      (usage.latinFontFamily === undefined && !usage.fontFamilies.some(Boolean)))) {
      // A Latin slot can inherit the theme even when another script slot has an
      // authored face. The inherited bold/italic axes need their own resource.
      add(doc.minorFont, usage.bold, usage.italic);
    }
  }
  return [...tuples.values()];
}


export function docxScriptCjkLanguage(doc: DocxDocumentModel): CjkLang | null {
  const scripts = new ScriptPreloadAccumulator(null);
  scripts.addText(docxTextRuns(doc));
  return scripts.scriptCjkLanguage();
}

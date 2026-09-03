import {
  classifyCjkFont,
  chartFontFamilies,
  scriptPreloadNamesForText,
} from '@silurus/ooxml-core';
import type {
  DocxDocumentModel,
} from './types.js';
import { docxRenderedCharts, docxRenderedTextUsages } from './document-content.js';

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
): (string | null | undefined)[] {
  const cjkLang =
    classifyCjkFont(doc.majorFont) ?? classifyCjkFont(doc.minorFont) ?? null;
  return [
    doc.majorFont,
    doc.minorFont,
    ...scriptPreloadNamesForText(docxTextRuns(doc), cjkLang),
  ];
}

/** Authored families offered to an application font provider. Google-only
 * script fallback names deliberately stay in {@link docxFontPreloadNames}. */
export function docxFontProviderNames(doc: DocxDocumentModel): string[] {
  const names = new Set<string>();
  if (doc.majorFont) names.add(doc.majorFont);
  if (doc.minorFont) names.add(doc.minorFont);
  for (const usage of docxRenderedTextUsages(doc)) {
    for (const family of usage.fontFamilies) if (family) names.add(family);
  }
  for (const chart of docxRenderedCharts(doc)) {
    for (const family of chartFontFamilies(chart)) names.add(family);
  }
  return [...names];
}

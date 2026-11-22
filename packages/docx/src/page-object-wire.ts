import type { BodyElement, DocxDocumentModel } from './types.js';

/** Private producer wire (never authored by WordprocessingML): the drawings
 * anchored to one explicit physical page (ODF `text:anchor-page-number`),
 * as a page-local story. See `pageObjectStorySource`.
 *
 * It lives beside the shared model types, not in the parser model, so that
 * the layout can read it without depending on the parser model. */
export interface PageObjectStoryWire {
  /** The physical page, 1-based. */
  readonly page: number;
  body: BodyElement[];
}

/** The ECMA-376 §17.2.1 background drawing story of a model (private
 * `__pageBackground.body`): page-anchored, page-sized drawings laid out on
 * every page beneath all other page stories. Empty without one. */
export function documentPageBackgroundStory(doc: DocxDocumentModel): BodyElement[] {
  const background = (doc as DocxDocumentModel & { __pageBackground?: { body?: BodyElement[] } })
    .__pageBackground;
  return Array.isArray(background?.body) ? background.body : [];
}

/** The page objects of a model; empty for every DOCX document. */
export function documentPageObjects(doc: DocxDocumentModel): readonly PageObjectStoryWire[] {
  return (doc as DocxDocumentModel & { __pageObjects?: PageObjectStoryWire[] }).__pageObjects ?? [];
}

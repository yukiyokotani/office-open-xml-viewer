import type { LayoutCellBlock, LayoutSourceStore, LayoutStoryBlock } from './layout/layout-source-store.js';
import { visitPageParagraphs } from './layout/page-factory.js';
import { revisionIsOmitted } from './layout/revision-visibility.js';
import { sourceKey } from './layout/source-key.js';
import type { ParagraphAcquisitionInput } from './layout/text.js';
import type { DeepReadonly, DocumentLayout, SourceRef } from './layout/types.js';
import type { DocxStorySource } from './types.js';

/** One heading of the document outline returned by `DocxDocument.getOutline()`. */
export interface DocxOutlineEntry {
  /** 1–9: ECMA-376 §17.3.1.20 `w:outlineLvl` + 1, resolved from the
   *  paragraph's own properties or its paragraph style chain (`w:basedOn`). */
  readonly level: number;
  /** Plain heading text, with whitespace collapsed (tabs and line breaks become
   *  single spaces). Deleted and moved-away tracked text is omitted, and the
   *  list number of a numbered heading is not included. */
  readonly text: string;
  /** Authored `w14:paraId`, the same value `DocxTextRunInfo.paragraphId`
   *  reports; absent when the paragraph carries none. */
  readonly paragraphId?: string;
  /** Structural source of the heading paragraph, the same identity
   *  `DocxTextRunInfo.source` reports. Always the body story. */
  readonly source: Readonly<DocxStorySource>;
  /** 0-based page on which the heading paragraph begins in the selected layout
   *  view; absent while that page is not laid out yet (progressive layout) or
   *  when the paragraph produces no rendered fragment (for example hidden text). */
  readonly pageIndex?: number;
}

/**
 * Project the heading outline of a document from its retained layout source and
 * one selected layout.
 *
 * ECMA-376 §17.3.1.20: a paragraph belongs to the outline when it has an outline
 * level 0–8; the parser resolves the level through the paragraph's direct
 * properties and its style hierarchy, which is how Word's built-in Heading 1–9
 * styles assign levels. Style names are not consulted.
 *
 * Scope (library policy) is the body story, including paragraphs in body
 * tables. Headers, footers, notes and text boxes are excluded.
 *
 * Pages come from the retained paragraph fragments of `layout`, so they follow
 * the same view (tracked-change markup, field date) as paint, and a heading that
 * spans a page break reports the page where it begins. The walk is linear in
 * body blocks plus retained page paragraphs and is performed only on request.
 */
export function buildDocxOutline(
  source: LayoutSourceStore,
  layout: DeepReadonly<DocumentLayout> | null,
): DocxOutlineEntry[] {
  const headings: { block: ParagraphAcquisitionInput; source: SourceRef }[] = [];
  const visit = (
    blocks: readonly (LayoutStoryBlock | LayoutCellBlock)[],
    prefix: readonly number[],
  ): void => {
    blocks.forEach((block, index) => {
      const path = [...prefix, index];
      if (block.type === 'paragraph') {
        const paragraph = block as ParagraphAcquisitionInput;
        if (isOutlineLevel(paragraph.outlineLevel)) {
          headings.push({
            block: paragraph,
            source: { story: 'body', storyInstance: 'body', path },
          });
        }
      } else if (block.type === 'table') {
        block.rows.forEach((row, rowIndex) => row.cells.forEach((cell, cellIndex) => {
          visit(cell.content as readonly LayoutCellBlock[], [...path, rowIndex, cellIndex]);
        }));
      }
    });
  };
  visit(source.blocks.body, []);
  if (headings.length === 0) return [];

  const pages = layout ? bodyParagraphPages(layout) : new Map<string, number>();
  return headings.map(({ block, source: headingSource }): DocxOutlineEntry => {
    const pageIndex = pages.get(sourceKey(headingSource));
    return {
      level: block.outlineLevel! + 1,
      text: headingText(block),
      ...(block.paragraphId ? { paragraphId: block.paragraphId } : {}),
      source: headingSource,
      ...(pageIndex === undefined ? {} : { pageIndex }),
    };
  });
}

function isOutlineLevel(level: number | undefined): level is number {
  return level !== undefined && Number.isInteger(level) && level >= 0 && level <= 8;
}

/** First page of each retained body-story paragraph. */
function bodyParagraphPages(layout: DeepReadonly<DocumentLayout>): Map<string, number> {
  const pages = new Map<string, number>();
  for (const page of layout.pages) {
    if (page.parityBlank) continue;
    visitPageParagraphs(page, (paragraph) => {
      if (paragraph.source.story !== 'body') return;
      const key = sourceKey(paragraph.source);
      if (!pages.has(key)) pages.set(key, page.pageIndex);
    });
  }
  return pages;
}

/** Accepted-final plain text, the same run projection `toMarkdown()` uses for a
 *  heading: text runs, cached field results, breaks and tabs as spaces. */
function headingText(paragraph: ParagraphAcquisitionInput): string {
  let text = '';
  for (const run of paragraph.runs) {
    if ('revision' in run && revisionIsOmitted(run.revision?.kind, false)) continue;
    if (run.type === 'text') text += run.text;
    else if (run.type === 'field') text += run.fallbackText;
    else if (run.type === 'break' || run.type === 'ptab') text += ' ';
  }
  return text.replace(/\s+/g, ' ').trim();
}

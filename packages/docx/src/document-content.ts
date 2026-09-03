import type {
  BodyElement,
  DocParagraph,
  DocRun,
  DocxTextRun,
  FieldRun,
  DocTable,
  DocxDocumentModel,
  HeadersFooters,
  ShapeRun,
  ShapeText,
} from './types.js';
import type { ChartModel } from '@silurus/ooxml-core';

type InternalRenderedFontAxes = Readonly<{
  fontFamilyHighAnsi?: string | null;
  fontFamilyEastAsia?: string | null;
  fontFamilyCs?: string | null;
  boldCs?: boolean;
  italicCs?: boolean;
}>;

/** One rendered string and every authored font family that can supply it.
 * Empty text records are intentional: paragraph marks and drawing anchors can
 * affect line metrics even when they paint no glyphs. */
export interface DocxRenderedTextUsage {
  text: string;
  fontFamilies: readonly (string | null | undefined)[];
  bold?: boolean;
  italic?: boolean;
}

function* shapeTextUsages(shape: ShapeRun): Generator<DocxRenderedTextUsage> {
  if (shape.textPath) {
    yield { text: shape.textPath.string, fontFamilies: [shape.textPath.fontFamily], bold: shape.textPath.bold, italic: shape.textPath.italic };
  }
  for (const block of shape.textBlocks ?? []) {
    yield* shapeBlockUsages(block);
  }
}

function* shapeBlockUsages(block: ShapeText): Generator<DocxRenderedTextUsage> {
  if (block.numbering) {
    yield {
      text: block.numbering.text,
      fontFamilies: [block.numbering.fontFamily, block.numbering.fontFamilyEastAsia],
      bold: false,
      italic: false,
    };
  }
  if (block.runs?.length) {
    for (const run of block.runs) {
      yield {
        text: run.text,
        // A run without an explicit axis inherits the block-level face.
        fontFamilies: [
          run.fontFamily,
          run.fontFamilyEastAsia,
          block.fontFamily,
        ],
        bold: run.bold ?? block.bold,
        italic: run.italic ?? block.italic,
      };
    }
  } else {
    yield { text: block.text, fontFamilies: [block.fontFamily], bold: block.bold, italic: block.italic };
  }
}

function* runUsages(run: DocRun): Generator<DocxRenderedTextUsage> {
  if (run.type === 'text') {
    const text = run as DocxTextRun & InternalRenderedFontAxes;
    yield {
      text: run.text,
      fontFamilies: [run.fontFamily, text.fontFamilyHighAnsi, run.fontFamilyEastAsia],
      bold: run.bold,
      italic: run.italic,
    };
    yield {
      text: run.text,
      fontFamilies: [run.fontFamilyCs],
      bold: run.boldCs ?? false,
      italic: run.italicCs ?? false,
    };
  } else if (run.type === 'field') {
    const field = run as FieldRun & InternalRenderedFontAxes;
    yield {
      text: field.fallbackText,
      fontFamilies: [field.fontFamily, field.fontFamilyHighAnsi, field.fontFamilyEastAsia],
      bold: field.bold,
      italic: field.italic,
    };
    yield {
      text: field.fallbackText,
      fontFamilies: [field.fontFamilyCs],
      bold: field.boldCs ?? false,
      italic: field.italicCs ?? false,
    };
  } else if (run.type === 'shape') {
    yield* shapeTextUsages(run);
  } else if (run.type === 'anchorHost') {
    yield {
      text: '',
      fontFamilies: [run.fontFamily, run.fontFamilyEastAsia],
      bold: run.bold,
      italic: run.italic,
    };
  }
}

function* paragraphUsages(paragraph: DocParagraph): Generator<DocxRenderedTextUsage> {
  // Empty paragraphs still reserve the resolved paragraph-mark line box.
  yield {
    text: '',
    fontFamilies: [paragraph.defaultFontFamily, paragraph.defaultFontFamilyEastAsia],
  };
  if (paragraph.numbering) {
    yield {
      text: paragraph.numbering.text,
      fontFamilies: [
        paragraph.numbering.fontFamily,
        paragraph.numbering.fontFamilyEastAsia,
      ],
    };
  }
  for (const run of paragraph.runs) yield* runUsages(run);
}

function* tableUsages(table: DocTable): Generator<DocxRenderedTextUsage> {
  for (const row of table.rows) {
    for (const cell of row.cells) {
      yield* bodyUsages(cell.content as BodyElement[]);
    }
  }
}

function* headerFooterUsages(
  stories: HeadersFooters | null | undefined,
): Generator<DocxRenderedTextUsage> {
  if (!stories) return;
  for (const story of [stories.default, stories.first, stories.even]) {
    if (story) yield* bodyUsages(story.body);
  }
}

function* bodyUsages(body: readonly BodyElement[]): Generator<DocxRenderedTextUsage> {
  for (const element of body) {
    if (element.type === 'paragraph') {
      yield* paragraphUsages(element);
    } else if (element.type === 'table') {
      yield* tableUsages(element);
    } else if (element.type === 'sectionBreak') {
      // Non-final sections keep their resolved header/footer stories on the
      // marker; the top-level sets represent only the final section.
      yield* headerFooterUsages(element.headers);
      yield* headerFooterUsages(element.footers);
    }
  }
}

/** Traverse every rendered DOCX story once. This is shared by script-aware web
 * font preloading and exact-local metric discovery so those paths cannot drift
 * on nested tables, section headers/footers, notes, or drawing text. Comments
 * are excluded because the page renderer does not paint comment bodies. */
export function* docxRenderedTextUsages(
  doc: DocxDocumentModel,
): Generator<DocxRenderedTextUsage> {
  yield* bodyUsages(doc.body ?? []);
  yield* headerFooterUsages(doc.headers);
  yield* headerFooterUsages(doc.footers);
  for (const note of [...(doc.footnotes ?? []), ...(doc.endnotes ?? [])]) {
    yield* bodyUsages(note.content);
  }
}

/** Unique authored families in first-rendered-use order. */
export function docxRenderedFontFamilies(doc: DocxDocumentModel): string[] {
  const families = new Set<string>();
  for (const usage of docxRenderedTextUsages(doc)) {
    for (const family of usage.fontFamilies) {
      const trimmed = family?.trim();
      if (trimmed) families.add(trimmed);
    }
  }
  return [...families];
}

function* bodyCharts(body: readonly BodyElement[]): Generator<ChartModel> {
  for (const element of body) {
    if (element.type === 'paragraph') {
      for (const run of element.runs) {
        if (run.type === 'chart') yield run.chart;
      }
    } else if (element.type === 'table') {
      for (const row of element.rows) {
        for (const cell of row.cells) yield* bodyCharts(cell.content);
      }
    } else if (element.type === 'sectionBreak') {
      for (const stories of [element.headers, element.footers]) {
        for (const story of [stories?.default, stories?.first, stories?.even]) {
          if (story) yield* bodyCharts(story.body);
        }
      }
    }
  }
}

/** Every chart painted from the document's rendered stories. */
export function* docxRenderedCharts(doc: DocxDocumentModel): Generator<ChartModel> {
  yield* bodyCharts(doc.body ?? []);
  for (const stories of [doc.headers, doc.footers]) {
    for (const story of [stories?.default, stories?.first, stories?.even]) {
      if (story) yield* bodyCharts(story.body);
    }
  }
  for (const note of [...(doc.footnotes ?? []), ...(doc.endnotes ?? [])]) {
    yield* bodyCharts(note.content);
  }
}

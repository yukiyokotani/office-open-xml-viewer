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

type InternalRenderedFontAxes = Readonly<{
  fontFamilyHighAnsi?: string | null;
  langEastAsia?: string;
  fontFamilyEastAsia?: string | null;
  fontFamilyCs?: string | null;
}>;

/** One rendered string and every authored font family that can supply it.
 * Empty text records are intentional: paragraph marks and drawing anchors can
 * affect line metrics even when they paint no glyphs. */
export interface DocxRenderedTextUsage {
  text: string;
  eastAsiaLanguage?: string;
  fontFamilies: readonly (string | null | undefined)[];
  /** Latin/highAnsi slot after paragraph inheritance; null means theme minor.
   * Undefined marks usage records that describe only another script slot. */
  latinFontFamily?: string | null;
  bold?: boolean;
  italic?: boolean;
}

function* shapeTextUsages(shape: ShapeRun): Generator<DocxRenderedTextUsage> {
  if (shape.textPath) {
    yield {
      text: shape.textPath.string,
      fontFamilies: [shape.textPath.fontFamily],
    };
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
        latinFontFamily: run.fontFamily ?? block.fontFamily ?? null,
        bold: run.bold ?? block.bold,
        italic: run.italic ?? block.italic,
      };
    }
  } else {
    yield {
      text: block.text,
      fontFamilies: [block.fontFamily],
      latinFontFamily: block.fontFamily ?? null,
      bold: block.bold,
      italic: block.italic,
    };
  }
}

function* runUsages(run: DocRun): Generator<DocxRenderedTextUsage> {
  if (run.type === 'text') {
    const text = run as DocxTextRun & InternalRenderedFontAxes;
    yield {
      text: run.text,
      eastAsiaLanguage: text.langEastAsia,
      fontFamilies: [run.fontFamily, text.fontFamilyHighAnsi, run.fontFamilyEastAsia],
      latinFontFamily: run.fontFamily ?? text.fontFamilyHighAnsi ?? null,
      bold: run.bold,
      italic: run.italic,
    };
    if (run.fontFamilyCs) yield {
      text: run.text,
      eastAsiaLanguage: text.langEastAsia,
      fontFamilies: [run.fontFamilyCs],
      bold: run.bold,
      italic: run.italic,
    };
  } else if (run.type === 'field') {
    const field = run as FieldRun & InternalRenderedFontAxes;
    yield {
      text: field.fallbackText,
      eastAsiaLanguage: field.langEastAsia,
      fontFamilies: [field.fontFamily, field.fontFamilyHighAnsi, field.fontFamilyEastAsia],
      latinFontFamily: field.fontFamily ?? field.fontFamilyHighAnsi ?? null,
      bold: field.bold,
      italic: field.italic,
    };
    if (field.fontFamilyCs) yield {
      text: field.fallbackText,
      eastAsiaLanguage: field.langEastAsia,
      fontFamilies: [field.fontFamilyCs],
      bold: field.bold,
      italic: field.italic,
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
  for (const run of paragraph.runs) {
    for (const usage of runUsages(run)) {
      const inherited = usage.fontFamilies.some(Boolean) ? usage.fontFamilies
        : [paragraph.defaultFontFamily, paragraph.defaultFontFamilyEastAsia];
      yield {
        ...usage,
        fontFamilies: inherited,
        ...(usage.latinFontFamily === null
          ? { latinFontFamily: paragraph.defaultFontFamily ?? null }
          : {}),
      };
    }
  }
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

/** Traverse every rendered DOCX story once. Script-aware web preloading and
 * resolved native-resource probing share this traversal so those paths cannot
 * drift on nested tables, section headers/footers, notes, or drawing text.
 * Comments are excluded because the page renderer does not paint them. */
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

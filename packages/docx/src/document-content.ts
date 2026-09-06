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
  /** Families eligible for ASCII/high-ANSI scalars in this rendered string. */
  latinFontFamilies?: readonly (string | null | undefined)[];
  /** Families eligible for East-Asian scalars in this rendered string. */
  eastAsianFontFamilies?: readonly (string | null | undefined)[];
  bold?: boolean;
  italic?: boolean;
}

function* shapeTextUsages(shape: ShapeRun): Generator<DocxRenderedTextUsage> {
  if (shape.textPath) {
    yield {
      text: shape.textPath.string,
      fontFamilies: [shape.textPath.fontFamily],
      latinFontFamilies: [shape.textPath.fontFamily],
      eastAsianFontFamilies: [shape.textPath.fontFamily],
      bold: shape.textPath.bold,
      italic: shape.textPath.italic,
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
      latinFontFamilies: [block.numbering.fontFamily],
      eastAsianFontFamilies: [block.numbering.fontFamilyEastAsia ?? block.numbering.fontFamily],
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
        latinFontFamilies: [run.fontFamily, block.fontFamily],
        eastAsianFontFamilies: [
          run.fontFamilyEastAsia ?? run.fontFamily ?? block.fontFamily,
        ],
        bold: run.bold ?? block.bold,
        italic: run.italic ?? block.italic,
      };
    }
  } else {
    yield {
      text: block.text,
      fontFamilies: [block.fontFamily],
      latinFontFamilies: [block.fontFamily],
      eastAsianFontFamilies: [block.fontFamily],
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
      fontFamilies: [run.fontFamily, text.fontFamilyHighAnsi, run.fontFamilyEastAsia],
      latinFontFamilies: [run.fontFamily, text.fontFamilyHighAnsi],
      eastAsianFontFamilies: [run.fontFamilyEastAsia ?? run.fontFamily],
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
      latinFontFamilies: [field.fontFamily, field.fontFamilyHighAnsi],
      eastAsianFontFamilies: [field.fontFamilyEastAsia ?? field.fontFamily],
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
      latinFontFamilies: [run.fontFamily],
      eastAsianFontFamilies: [run.fontFamilyEastAsia ?? run.fontFamily],
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
      latinFontFamilies: [paragraph.numbering.fontFamily],
      eastAsianFontFamilies: [
        paragraph.numbering.fontFamilyEastAsia ?? paragraph.numbering.fontFamily,
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

export interface DocxResolvedFontMetricCandidate {
  readonly family: string;
  readonly probeText: string;
  /** The exact regular face is selected by an ASCII/high-ANSI slot as well as
   * an East-Asian slot, so the observed single-line allocation can apply to
   * that Latin route. */
  readonly appliesToLatin: boolean;
}

const EAST_ASIAN_SCALAR = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Bopomofo}\p{Script=Yi}]/u;
// Whitespace and generic punctuation do not select the Latin font slot. Keep
// the probe tied to glyphs that can actually make the ASCII/high-ANSI route
// win; otherwise a spaced CJK run can incorrectly broaden the metric to Latin.
const LATIN_SCALAR = /[0-9\p{Script=Latin}]/u;

function firstMatchingScalar(text: string, pattern: RegExp): string | undefined {
  for (const scalar of text) if (pattern.test(scalar)) return scalar;
  return undefined;
}

function charsetProbe(charset: string | undefined): string | undefined {
  switch (charset?.trim().toLowerCase()) {
    case '80':
    case '86':
      return '国';
    case '81':
    case '82':
      return '가';
    case '88':
      return '國';
    default:
      return undefined;
  }
}

/** Identify regular rendered script routes that can legitimately consume an
 * East-Asian resource metric. Font-table declarations that never win a script
 * slot are excluded; this prevents an unused East-Asian default from changing
 * Latin-only pagination. */
export function docxResolvedFontMetricCandidates(
  doc: DocxDocumentModel,
  fontFamilyCharsets: Readonly<Record<string, string>> = {},
): DocxResolvedFontMetricCandidate[] {
  const charsets = Object.fromEntries(Object.entries(fontFamilyCharsets)
    .map(([family, charset]) => [family.trim().toLocaleLowerCase('en-US'), charset]));
  const candidates = new Map<string, DocxResolvedFontMetricCandidate>();
  const add = (
    familyValue: string | null | undefined,
    probeText: string | undefined,
    appliesToLatin: boolean,
  ): void => {
    const family = familyValue?.trim();
    if (!family || !probeText) return;
    const key = family.toLocaleLowerCase('en-US');
    const previous = candidates.get(key);
    candidates.set(key, {
      family: previous?.family ?? family,
      probeText: previous?.probeText ?? probeText,
      appliesToLatin: (previous?.appliesToLatin ?? false) || appliesToLatin,
    });
  };

  for (const usage of docxRenderedTextUsages(doc)) {
    if (usage.bold || usage.italic) continue;
    const eastAsianText = firstMatchingScalar(usage.text, EAST_ASIAN_SCALAR);
    if (eastAsianText) {
      for (const family of usage.eastAsianFontFamilies ?? []) {
        add(family, eastAsianText, false);
      }
    }
    if (firstMatchingScalar(usage.text, LATIN_SCALAR)) {
      for (const familyValue of usage.latinFontFamilies ?? []) {
        const family = familyValue?.trim();
        if (!family) continue;
        add(family, charsetProbe(charsets[family.toLocaleLowerCase('en-US')]), true);
      }
    }
  }
  return [...candidates.values()];
}

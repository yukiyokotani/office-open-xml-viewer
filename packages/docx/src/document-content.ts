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
import {
  numberingMarkerShapeInput,
  paragraphMarkShapeInput,
  inlineObjectShapeInput,
  type InternalDocxDocumentModel,
  type InternalRunSlotMetadata,
  type InternalShapeRun,
} from './parser-model.js';
import { requestedFamily, shapeRunToDocRun, type TextFontSlots } from './layout/text.js';
import type { NumberingMarkerShapeInput } from './layout/types.js';

type InternalRenderedFontAxes = Readonly<InternalRunSlotMetadata>;

// Font-source demand (LoadOptions.fonts) admits newly resolved slot families.
// Native preflight retains its existing projection so configuring an optional
// font source cannot change default local registrations or layout.
type FontCollection = 'native-preflight' | 'font-demand';

type SlotRequest = Pick<NumberingMarkerShapeInput, 'fonts' | 'themeFonts' | 'themeFontPresence'>;

/** The family the text shaper requests for each §17.3.2.26 slot. Resource
 * collection resolves slots with the shaper's own {@link requestedFamily}, so a
 * theme reference, a direct/style-resolved face and the ascii fallback reach
 * the preload set exactly as they reach Canvas font selection. */
function slotFamilies(
  request: SlotRequest,
  slots: readonly ('ascii' | 'highAnsi' | 'eastAsia' | 'complexScript')[],
): (string | null | undefined)[] {
  return slots.map((slot) => requestedFamily(request, slot));
}

/** Slot request for an ordinary text or field result. Parser output carries
 * the resolved direct/theme slots; hand-built public runs fall back to the
 * single-axis projection exactly as the segment builder does. */
function runSlotRequest(
  ascii: string | null | undefined,
  facts: InternalRenderedFontAxes,
): SlotRequest {
  const base = ascii ?? null;
  const fallback: TextFontSlots = {
    ascii: base,
    highAnsi: facts.fontFamilyHighAnsi ?? base,
    eastAsia: facts.fontFamilyEastAsia ?? base,
    complexScript: facts.fontFamilyCs ?? base,
  };
  return {
    fonts: facts.fontSlots?.direct ?? fallback,
    themeFonts: facts.fontSlots?.theme,
    themeFontPresence: facts.fontSlots?.themePresent,
  };
}

/** Whether a run can shape any character through its complex-script slot:
 * w:cs/w:rtl force it (§17.3.2.26 step 2), and an authored or themed cs face
 * governs complex-script characters in the run's text. */
function runUsesComplexScriptSlot(facts: InternalRenderedFontAxes): boolean {
  return facts.cs === true || facts.rtl === true || facts.fontFamilyCs != null
    || facts.fontSlots?.direct.complexScript != null
    || facts.fontSlots?.themePresent.complexScript === true;
}

/** Marker and paragraph-mark usage from the production shape projection. */
function shapeInputUsage(text: string, input: NumberingMarkerShapeInput): DocxRenderedTextUsage {
  return {
    text,
    eastAsiaLanguage: input.eastAsiaLanguage,
    fontFamilies: slotFamilies(input, ['ascii', 'highAnsi', 'eastAsia', 'complexScript']),
    bold: input.weight >= 700,
    italic: input.style === 'italic',
  };
}

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

function* shapeTextUsages(shape: ShapeRun, collection: FontCollection): Generator<DocxRenderedTextUsage> {
  if (shape.textPath) {
    yield {
      text: shape.textPath.string,
      fontFamilies: [shape.textPath.fontFamily],
    };
  }
  // Parser-created text boxes lay out their complete w:txbxContent story
  // (paragraphs AND tables, recursively). The legacy textBlocks projection
  // omits tables, so it is consulted only for hand-built public shapes.
  const content = (shape as InternalShapeRun).textBoxContent;
  if (collection === 'font-demand' && content !== undefined) {
    yield* bodyUsages(content as BodyElement[], collection);
    return;
  }
  for (const block of shape.textBlocks ?? []) {
    yield* shapeBlockUsages(block, collection, content === undefined);
  }
}

function* shapeBlockUsages(
  block: ShapeText,
  collection: FontCollection,
  compatibility: boolean,
): Generator<DocxRenderedTextUsage> {
  if (block.numbering) {
    yield {
      text: block.numbering.text,
      fontFamilies: [block.numbering.fontFamily, block.numbering.fontFamilyEastAsia],
    };
  }
  if (block.runs?.length) {
    for (const run of block.runs) {
      if (compatibility && (run.rtl || run.cs)) {
        // §17.3.2.7/§17.3.2.30 select CS formatting even for digits. Apply
        // block inheritance as normalizeTextBoxInput does, then reuse layout's
        // compatibility fallback and the body's independent CS usage tuple.
        // Keep unflagged runs and parser-created native projections unchanged.
        yield* runUsages(shapeRunToDocRun({
          ...run,
          fontFamily: run.fontFamily ?? block.fontFamily,
          bold: run.bold ?? block.bold,
          italic: run.italic ?? block.italic,
        }), collection);
        continue;
      }
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

function* textResultUsages(
  text: string,
  ascii: string | null | undefined,
  facts: InternalRenderedFontAxes,
  bold: boolean | undefined,
  italic: boolean | undefined,
  collection: FontCollection,
): Generator<DocxRenderedTextUsage> {
  const request = runSlotRequest(ascii, facts);
  const [asciiFamily, highAnsiFamily, eastAsiaFamily] =
    collection === 'font-demand'
      ? slotFamilies(request, ['ascii', 'highAnsi', 'eastAsia'])
      : [ascii, facts.fontFamilyHighAnsi, facts.fontFamilyEastAsia];
  yield {
    text,
    eastAsiaLanguage: facts.langEastAsia,
    fontFamilies: [asciiFamily, highAnsiFamily, eastAsiaFamily],
    latinFontFamily: asciiFamily ?? highAnsiFamily ?? null,
    bold,
    italic,
  };
  if (collection === 'font-demand' ? runUsesComplexScriptSlot(facts) : facts.fontFamilyCs != null) {
    const complexScriptFamily = collection === 'font-demand'
      ? requestedFamily(request, 'complexScript') : facts.fontFamilyCs;
    if (complexScriptFamily) yield {
      text,
      eastAsiaLanguage: facts.langEastAsia,
      fontFamilies: [complexScriptFamily],
      // ECMA-376 §17.3.2.3/§17.3.2.17: bCs/iCs are independent of b/i.
      // Probe the tuple that complex-script paint actually requests.
      bold: facts.boldCs ?? false,
      italic: facts.italicCs ?? false,
    };
  }
}

function* runUsages(run: DocRun, collection: FontCollection): Generator<DocxRenderedTextUsage> {
  const inlineFont = inlineObjectShapeInput(run);
  if (inlineFont) yield shapeInputUsage('', inlineFont);
  if (run.type === 'text') {
    const text = run as DocxTextRun & InternalRenderedFontAxes;
    yield* textResultUsages(run.text, run.fontFamily, text, run.bold, run.italic, collection);
  } else if (run.type === 'field') {
    const field = run as FieldRun & InternalRenderedFontAxes;
    yield* textResultUsages(field.fallbackText, field.fontFamily, field, field.bold, field.italic, collection);
  } else if (run.type === 'shape') {
    yield* shapeTextUsages(run, collection);
  } else if (run.type === 'anchorHost') {
    yield {
      text: '',
      fontFamilies: [run.fontFamily, run.fontFamilyEastAsia],
      bold: run.bold,
      italic: run.italic,
    };
  }
}

function* paragraphUsages(paragraph: DocParagraph, collection: FontCollection): Generator<DocxRenderedTextUsage> {
  // Empty paragraphs still reserve the resolved paragraph-mark line box.
  yield {
    text: '',
    fontFamilies: [paragraph.defaultFontFamily, paragraph.defaultFontFamilyEastAsia],
  };
  if (collection === 'font-demand') {
    const mark = paragraphMarkShapeInput(paragraph);
    if (mark) yield shapeInputUsage('', mark);
  }
  if (paragraph.numbering) {
    // Same effective numbering-level rPr projection (all four slots and their
    // theme references) that production marker shaping consumes.
    if (collection === 'font-demand') yield shapeInputUsage(
      paragraph.numbering.text,
      numberingMarkerShapeInput(paragraph.numbering, paragraph.defaultFontSize ?? 10),
    );
    else yield {
      text: paragraph.numbering.text,
      fontFamilies: [paragraph.numbering.fontFamily, paragraph.numbering.fontFamilyEastAsia],
    };
  }
  for (const run of paragraph.runs) {
    for (const usage of runUsages(run, collection)) {
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

function* tableUsages(table: DocTable, collection: FontCollection): Generator<DocxRenderedTextUsage> {
  for (const row of table.rows) {
    for (const cell of row.cells) {
      yield* bodyUsages(cell.content as BodyElement[], collection);
    }
  }
}

function* headerFooterUsages(
  stories: HeadersFooters | null | undefined,
  collection: FontCollection,
): Generator<DocxRenderedTextUsage> {
  if (!stories) return;
  for (const story of [stories.default, stories.first, stories.even]) {
    if (story) yield* bodyUsages(story.body, collection);
  }
}

function* bodyUsages(body: readonly BodyElement[], collection: FontCollection): Generator<DocxRenderedTextUsage> {
  for (const element of body) {
    if (element.type === 'paragraph') {
      yield* paragraphUsages(element, collection);
    } else if (element.type === 'table') {
      yield* tableUsages(element, collection);
    } else if (element.type === 'sectionBreak') {
      // Non-final sections keep their resolved header/footer stories on the
      // marker; the top-level sets represent only the final section.
      yield* headerFooterUsages(element.headers, collection);
      yield* headerFooterUsages(element.footers, collection);
    }
  }
}

/** Reserved separator paragraphs. Their rule control and content mark are
 * text-free runs with their own character properties; registering them
 * through the ordinary paragraph usages lets a family used only by the
 * control join the same preload/resolution ownership as body text. */
function* reservedSeparatorUsages(
  doc: DocxDocumentModel,
  collection: FontCollection,
): Generator<DocxRenderedTextUsage> {
  const reserved = (doc as InternalDocxDocumentModel).__noteLayoutSettings?.reservedSeparators;
  for (const stories of [reserved?.footnote, reserved?.endnote]) {
    if (!stories) continue;
    for (const story of [stories.separator, stories.continuationSeparator, stories.continuationNotice]) {
      const authored = story.paragraph;
      if (!authored) continue;
      const runs = [story.rule?.control.run, authored.contentMark?.run]
        .filter((run): run is NonNullable<typeof run> => run !== undefined)
        .map((run) => ({ ...run, type: 'text', text: '' }));
      yield* paragraphUsages({ ...authored.paragraph, runs } as unknown as DocParagraph, collection);
    }
  }
}

/** Ordinary DOCX formatted listed separator paragraphs (§17.11.9). Layout
 * acquires their paragraph-mark line box, so its families join the same
 * preload/resolution ownership as body text. */
function* selectedSeparatorUsages(
  doc: DocxDocumentModel,
  collection: FontCollection,
): Generator<DocxRenderedTextUsage> {
  const settings = (doc as InternalDocxDocumentModel).__noteLayoutSettings;
  for (const paragraph of [settings?.footnoteSeparatorParagraph, settings?.footnoteContinuationSeparatorParagraph,
    settings?.endnoteSeparatorParagraph, settings?.endnoteContinuationSeparatorParagraph]) {
    if (paragraph) yield* paragraphUsages({ ...paragraph, runs: [] }, collection);
  }
}

/** Traverse every rendered DOCX story once. Script-aware web preloading and
 * resolved native-resource probing share this traversal so those paths cannot
 * drift on nested tables, section headers/footers, notes, or drawing text.
 * Comments are excluded because the page renderer does not paint them. */
export function* docxRenderedTextUsages(
  doc: DocxDocumentModel,
  collection: FontCollection = 'native-preflight',
): Generator<DocxRenderedTextUsage> {
  yield* bodyUsages(doc.body ?? [], collection);
  yield* headerFooterUsages(doc.headers, collection);
  yield* headerFooterUsages(doc.footers, collection);
  for (const note of [...(doc.footnotes ?? []), ...(doc.endnotes ?? [])]) {
    yield* bodyUsages(note.content, collection);
  }
  yield* reservedSeparatorUsages(doc, collection);
  yield* selectedSeparatorUsages(doc, collection);
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

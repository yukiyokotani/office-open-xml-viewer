import { expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { docxRenderedFontFamilies } from './document-content.js';
import { layoutDocument } from './document-layout.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types.js';
import type { DocumentLayout, NoteLayout, PaintNode } from './layout/types.js';

// Public-model boundary for listed formatted footnote separator stories
// (ECMA-376 §17.11.9) whose paragraph geometry the DOCX parser retains. The
// shapes follow the public Word 16.113.3 A24/B720/B719 controls: 72pt margins,
// P01–P14 then B01/B02 in exact 10pt lines, and a four-line exact 10pt note.
// Synthetic glyph metrics are not a font oracle; every band is the story's
// own exact line, so no scalar or empirical band enters the expectations.
const context = {
  font: '10px Arial', letterSpacing: '0px', fontKerning: 'none',
  measureText: (text: string) => ({
    width: [...text].length * 6,
    fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
    actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
  }),
} as unknown as CanvasRenderingContext2D;

const label = (prefix: string, index: number) => prefix + String(index + 1).padStart(2, '0');
const lines = (prefix: string, count: number) => Array.from({ length: count }, (_, index) => label(prefix, index));

function textRun(text: string, extra: Record<string, unknown> = {}) {
  return {
    type: 'text', text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: 'Arial', fontFamilyEastAsia: '', isLink: false,
    background: null, vertAlign: null, hyperlink: null, ...extra,
  };
}

/** Exact `linePt` lines; `references[i]` lists the note ids opening line i. */
function paragraph(texts: readonly string[], references: readonly (readonly string[])[] = [], linePt = 10): DocParagraph {
  const runs = texts.flatMap((line, index) => [
    ...(references[index] ?? []).map(id => textRun(id, { noteRef: { kind: 'footnote', id } })),
    textRun(line),
    ...(index + 1 < texts.length ? [{ type: 'break', breakType: 'line' }] : []),
  ]);
  return {
    type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: { value: linePt, rule: 'exact', explicit: true },
    numbering: null, tabStops: [], runs: runs as DocParagraph['runs'],
    defaultFontSize: 10, defaultFontFamily: 'Arial', widowControl: false,
  } as unknown as DocParagraph;
}

/** The parser's run-free effective paragraph of a formatted marker story. */
const story = (linePt: number): DocParagraph => paragraph([], [], linePt);

type Stories = Readonly<{ separator?: DocParagraph; continuationSeparator?: DocParagraph }>;

function model(
  pageHeight: number,
  body: DocParagraph[],
  note: DocParagraph,
  stories: Stories,
): DocxDocumentModel {
  return {
    section: {
      pageWidth: 324, pageHeight, marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
      headerDistance: 36, footerDistance: 36, titlePage: false, evenAndOddHeaders: false, sectionStart: 'nextPage',
    } as SectionProps,
    body: body as unknown as BodyElement[],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { Arial: 'swiss' },
    footnotes: [{ id: '1', content: [note] as unknown as BodyElement[] }],
    __noteLayoutSettings: {
      footnoteSeparator: 'short',
      footnoteContinuationSeparator: 'full',
      ...(stories.separator ? { footnoteSeparatorParagraph: stories.separator } : {}),
      ...(stories.continuationSeparator ? { footnoteContinuationSeparatorParagraph: stories.continuationSeparator } : {}),
    },
  } as unknown as DocxDocumentModel;
}

const controlBody = () => [paragraph(lines('P', 14)), paragraph(['B01', 'B02'], [['1']])];

const render = (document: DocxDocumentModel, allowFootnoteContinuation = true): DocumentLayout =>
  layoutDocument(document,
    createLayoutServices(document, { measureContext: context, allowFootnoteContinuation }),
    { currentDateMs: 0 });

function text(block: PaintNode): string {
  return block.kind === 'paragraph'
    ? block.lines.flatMap(line => line.placements.flatMap(run => run.kind === 'text' ? [run.text] : [])).join('')
    : '';
}

const notesOnPage = (page: DocumentLayout['pages'][number]): NoteLayout[] =>
  page.layers.notes.filter((node): node is NoteLayout => node.kind === 'note');
const pages = (layout: DocumentLayout) => layout.pages.map(page => ({
  body: page.layers.body.map(text),
  notes: notesOnPage(page).map(note => note.story.blocks.map(text).join('')),
}));
/** Band height, rule offset in the band and rule width of a page's first note. */
const band = (layout: DocumentLayout, pageIndex: number) => {
  const note = notesOnPage(layout.pages[pageIndex]!)[0]!;
  const rule = note.separator[0]!;
  return {
    heightPt: note.flowBounds.heightPt - note.story.advancePt,
    ruleOffsetPt: rule.from.yPt - note.flowBounds.yPt,
    ruleWidthPt: rule.to.xPt - rule.from.xPt,
  };
};

it('moves the reference line and its whole note past an authored 24pt separator story (A24 shape)', () => {
  // 180pt body: P01–P14 + B01 = 150, story band 24 + one note line = 184.
  // B01 cannot keep a real note line, so it moves with B02 and the note.
  const layout = render(model(324, controlBody(), paragraph(lines('N1L', 4)), {
    separator: story(24), continuationSeparator: story(24),
  }));
  expect(pages(layout)).toEqual([
    { body: [lines('P', 14).join('')], notes: [] },
    { body: ['1B01B02'], notes: [lines('N1L', 4).join('')] },
  ]);
  // The band is the story paragraph's exact line; the Short rule is
  // min(144pt, 180pt) long and sits 10pt/8 below the band middle
  // (WORD_NOTE_SEPARATOR_GEOMETRY).
  expect(band(layout, 1)).toEqual({ heightPt: 24, ruleOffsetPt: 13.25, ruleWidthPt: 144 });
});

it.each([
  // B720/B719: an authored exact 6pt story. 150 + 6 + 20 = 176 fits two
  // head lines on a 176pt body, but only one on a 175pt body.
  { pageHeight: 320, head: 2 },
  { pageHeight: 319, head: 1 },
])('keeps an authored 6pt story band on a $pageHeight pt page', ({ pageHeight, head }) => {
  const layout = render(model(pageHeight, controlBody(), paragraph(lines('N1L', 4)), {
    separator: story(6), continuationSeparator: story(6),
  }));
  expect(pages(layout)).toEqual([
    { body: [lines('P', 14).join(''), '1B01'], notes: [lines('N1L', 4).slice(0, head).join('')] },
    { body: ['B02'], notes: [lines('N1L', 4).slice(head).join('')] },
  ]);
  expect(band(layout, 0)).toEqual({ heightPt: 6, ruleOffsetPt: 4.25, ruleWidthPt: 144 });
  expect(band(layout, 1)).toEqual({ heightPt: 6, ruleOffsetPt: 4.25, ruleWidthPt: 180 });
});

it('charges each role its own story band and keeps scalar bands for an absent role', () => {
  // B01 + 24 + fourteen head lines = 174; the continuation story's 12pt band
  // then carries the sixteen remaining lines (172).
  const separate = render(model(324, [paragraph(['B01'], [['1']])], paragraph(lines('N', 30)), {
    separator: story(24), continuationSeparator: story(12),
  }));
  expect(pages(separate)).toEqual([
    { body: ['1B01'], notes: [lines('N', 14).join('')] },
    { body: [], notes: [lines('N', 30).slice(14).join('')] },
  ]);
  expect(band(separate, 0)).toEqual({ heightPt: 24, ruleOffsetPt: 13.25, ruleWidthPt: 144 });
  expect(band(separate, 1)).toEqual({ heightPt: 12, ruleOffsetPt: 7.25, ruleWidthPt: 180 });
  // An unformatted (absent) continuation role keeps the 6pt scalar band.
  const scalar = render(model(324, [paragraph(['B01'], [['1']])], paragraph(lines('N', 30)), {
    separator: story(24),
  }));
  expect(band(scalar, 1)).toEqual({ heightPt: 6, ruleOffsetPt: 3, ruleWidthPt: 180 });
});

it('charges the story band with continuation off and keeps notes whole', () => {
  const layout = render(model(324, controlBody(), paragraph(lines('N1L', 4)), {
    separator: story(24),
  }), false);
  expect(pages(layout)).toEqual([
    { body: [lines('P', 14).join('')], notes: [] },
    { body: ['1B01B02'], notes: [lines('N1L', 4).join('')] },
  ]);
  expect(band(layout, 1)).toEqual({ heightPt: 24, ruleOffsetPt: 13.25, ruleWidthPt: 144 });
});

/** An exact-line story with authored spacing before/after its mark line. */
const spaced = (linePt: number, spaceBefore: number, spaceAfter: number): DocParagraph =>
  ({ ...story(linePt), spaceBefore, spaceAfter }) as DocParagraph;

it.each([
  // §17.3.1.33: spacing is allocated before/after the paragraph's lines, so
  // the mark's exact 10pt line box spans 0..10 or 12..22 of the band. The
  // rule stays 10pt/8 below that line box's midpoint, never in the spacing.
  { name: 'after', separator: spaced(10, 0, 40), expected: { heightPt: 50, ruleOffsetPt: 6.25, ruleWidthPt: 144 } },
  { name: 'before', separator: spaced(10, 12, 0), expected: { heightPt: 22, ruleOffsetPt: 18.25, ruleWidthPt: 144 } },
])('reserves authored $name spacing but draws the rule within the mark line box', ({ separator, expected }) => {
  const layout = render(model(324, [paragraph(['B01'], [['1']])], paragraph(['N01']), { separator }), false);
  expect(band(layout, 0)).toEqual(expected);
});

it('places each role\'s rule within its own mark line box at each page\'s band', () => {
  // B01 + 50 + twelve head lines = 180; then 22 + fifteen lines per page.
  const layout = render(model(324, [paragraph(['B01'], [['1']])], paragraph(lines('N', 30)), {
    separator: spaced(10, 0, 40), continuationSeparator: spaced(10, 12, 0),
  }));
  expect(pages(layout).map(page => page.notes)).toEqual([
    [lines('N', 12).join('')], [lines('N', 27).slice(12).join('')], [lines('N', 30).slice(27).join('')],
  ]);
  expect(band(layout, 0)).toEqual({ heightPt: 50, ruleOffsetPt: 6.25, ruleWidthPt: 144 });
  expect(band(layout, 1)).toEqual({ heightPt: 22, ruleOffsetPt: 18.25, ruleWidthPt: 180 });
  expect(band(layout, 2)).toEqual({ heightPt: 22, ruleOffsetPt: 18.25, ruleWidthPt: 180 });
});

it('loads the story paragraph mark family with the document fonts', () => {
  const separator = { ...story(24), defaultFontFamily: 'Story Mark Face' } as DocParagraph;
  expect(docxRenderedFontFamilies(model(324, controlBody(), paragraph(lines('N1L', 4)), { separator })))
    .toContain('Story Mark Face');
});

it.each([
  { name: 'a story paragraph without its Short/Full mark', settings: { footnoteSeparatorParagraph: story(24) } },
  { name: 'a story paragraph with runs', settings: {
    footnoteSeparator: 'short', footnoteSeparatorParagraph: paragraph(['x'], [], 24),
  } },
  { name: 'a story paragraph with borders', settings: {
    footnoteSeparator: 'short',
    footnoteSeparatorParagraph: { ...story(24), borders: { top: { style: 'single', width: 0.5, color: '000000', space: 0 } } },
  } },
])('rejects $name instead of approximating it', ({ settings }) => {
  const document = {
    ...model(324, controlBody(), paragraph(lines('N1L', 4)), {}),
    __noteLayoutSettings: settings,
  } as unknown as DocxDocumentModel;
  expect(() => render(document)).toThrow(/note separator story/);
});

// Parser tests establish that a bare selected story retains its inherited
// exact 12pt paragraph. Both note kinds use that retained paragraph here.
it.each(['footnote', 'endnote'] as const)('measures bare selected %s stories with and without rule ink', (kind) => {
  for (const mark of ['short', 'none'] as const) {
    const reference = paragraph(['B01'], [['1']]);
    reference.runs = reference.runs.map(run => 'noteRef' in run
      ? { ...run, noteRef: { kind, id: '1' } } : run);
    const document = {
      ...model(324, [reference], paragraph(['N01']), {}),
      [kind === 'footnote' ? 'footnotes' : 'endnotes']: [{ id: '1', content: [paragraph(['N01'])] }],
      __noteLayoutSettings: {
        [kind === 'footnote' ? 'footnoteSeparator' : 'endnoteSeparator']: mark,
        [kind === 'footnote' ? 'footnoteSeparatorParagraph' : 'endnoteSeparatorParagraph']: story(12),
      },
    } as unknown as DocxDocumentModel;
    const [note] = notesOnPage(render(document).pages[0]!);
    expect(note!.flowBounds.heightPt - note!.story.advancePt).toBe(12);
    expect(note!.separator).toHaveLength(mark === 'none' ? 0 : 1);
  }
});

it.each([false, true])('uses the continued endnote area story even at a new-note boundary (%s)', separateNotes => {
  const base = model(324, [paragraph(['B01'])], paragraph([]), {});
  const body = paragraph(['B01']);
  body.runs.push(textRun('e', { noteRef: { kind: 'endnote', id: '1' } }) as DocParagraph['runs'][number]);
  if (separateNotes) body.runs.push(textRun('e', { noteRef: { kind: 'endnote', id: '2' } }) as DocParagraph['runs'][number]);
  const document = {
    ...base, body: [body], footnotes: [],
    endnotes: separateNotes
      ? [{ id: '1', content: [paragraph(lines('E', 14))] }, { id: '2', content: [paragraph(lines('E', 30).slice(14))] }]
      : [{ id: '1', content: [paragraph(lines('E', 30))] }],
    __noteLayoutSettings: { endnoteSeparator: 'short', endnoteContinuationSeparator: 'full',
      endnoteSeparatorParagraph: story(24), endnoteContinuationSeparatorParagraph: story(12) },
  } as unknown as DocxDocumentModel;
  const layout = render(document);
  expect(pages(layout).map(page => page.notes.join(''))).toEqual([
    lines('E', 14).join(''), lines('E', 30).slice(14).join(''),
  ]);
  expect(band(layout, 0)).toEqual({ heightPt: 24, ruleOffsetPt: 13.25, ruleWidthPt: 144 });
  expect(band(layout, 1)).toEqual({ heightPt: 12, ruleOffsetPt: 7.25, ruleWidthPt: 180 });
});

it('admits a listed endnote continuation story when the ordinary separator is absent', () => {
  const base = model(324, [paragraph(['B01'])], paragraph([]), {});
  const body = paragraph(['B01']);
  body.runs.push(textRun('e', { noteRef: { kind: 'endnote', id: '1' } }) as DocParagraph['runs'][number]);
  const document = { ...base, body: [body], footnotes: [],
    endnotes: [{ id: '1', content: [paragraph(lines('E', 30))] }],
    __noteLayoutSettings: { endnoteContinuationSeparator: 'full', endnoteContinuationSeparatorParagraph: story(12) },
  } as unknown as DocxDocumentModel;
  const layout = render(document);
  expect(pages(layout).map(page => page.notes.join(''))).toEqual([
    lines('E', 16).join(''), lines('E', 30).slice(16).join(''),
  ]);
  expect(band(layout, 1)).toEqual({ heightPt: 12, ruleOffsetPt: 7.25, ruleWidthPt: 180 });
});

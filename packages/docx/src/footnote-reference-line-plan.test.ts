import { expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types.js';
import type { DocumentLayout, NoteLayout, PaintNode } from './layout/types.js';

// Public-model boundary for reference-line footnote plans (continuation
// opted in). Exact 10pt lines and the 6pt missing-story band make every
// charge independently computable; synthetic glyph metrics are not an
// Office pagination or font oracle. Page 200x200pt, 10pt margins: 180pt body.
const context = {
  font: '10px Arial', letterSpacing: '0px', fontKerning: 'none',
  measureText: (text: string) => ({
    width: [...text].length * 6,
    fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
    actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
  }),
} as unknown as CanvasRenderingContext2D;

const label = (prefix: string, index: number, width = 2) => prefix + String(index + 1).padStart(width, '0');
const lines = (prefix: string, count: number, width = 2) =>
  Array.from({ length: count }, (_, index) => label(prefix, index, width));

function textRun(text: string, extra: Record<string, unknown> = {}) {
  return {
    type: 'text', text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: 'Arial', fontFamilyEastAsia: '', isLink: false,
    background: null, vertAlign: null, hyperlink: null, ...extra,
  };
}

/** Exact 10pt lines; `references[i]` lists the note ids that open line i. */
function paragraph(
  texts: readonly string[],
  references: readonly (readonly string[])[] = [],
  properties: Readonly<{ keepLines?: boolean; widowControl?: boolean }> = {},
): DocParagraph {
  const runs = texts.flatMap((line, index) => [
    ...(references[index] ?? []).map(id => textRun(id, { noteRef: { kind: 'footnote', id } })),
    textRun(line),
    ...(index + 1 < texts.length ? [{ type: 'break', breakType: 'line' }] : []),
  ]);
  return {
    type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: { value: 10, rule: 'exact', explicit: true },
    numbering: null, tabStops: [], runs: runs as DocParagraph['runs'],
    defaultFontSize: 10, defaultFontFamily: 'Arial', widowControl: false, ...properties,
  } as unknown as DocParagraph;
}

function model(body: DocParagraph[], footnotes: readonly (readonly DocParagraph[])[]): DocxDocumentModel {
  return {
    section: {
      pageWidth: 200, pageHeight: 200, marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false, sectionStart: 'nextPage',
    } as SectionProps,
    body: body as unknown as BodyElement[],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { Arial: 'swiss' },
    footnotes: footnotes.map((content, index) => ({ id: String(index + 1), content: content as unknown as BodyElement[] })),
  } as unknown as DocxDocumentModel;
}

const render = (document: DocxDocumentModel): DocumentLayout => layoutDocument(document,
  createLayoutServices(document, { measureContext: context, allowFootnoteContinuation: true }),
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

it.each([
  // §17.3.1.14: ten kept lines (100) + band 6 + seven note lines = 176 fit, so
  // the note head is planned against the whole kept paragraph.
  { count: 10, expected: [
    { body: ['1' + lines('B', 10).join('')], notes: [lines('N', 7).join('')] },
    { body: [], notes: [lines('N', 20).slice(7).join('')] },
  ] },
  // Twenty kept lines cannot fit any fresh page, so the keep cannot be
  // honoured: the reference line is planned alone (band 6 + 16 lines) and the
  // paragraph splits under the ordinary rules instead of failing.
  { count: 20, expected: [
    { body: ['1B01'], notes: [lines('N', 16).join('')] },
    { body: [lines('B', 20).slice(1, 14).join('')], notes: [lines('N', 20).slice(16).join('')] },
    { body: [lines('B', 20).slice(14).join('')], notes: [] },
  ] },
])('plans a first-line note against a $count-line keepLines paragraph', ({ count, expected }) => {
  const layout = render(model(
    [paragraph(lines('B', count), [['1']], { keepLines: true })],
    [[paragraph(lines('N', 20))]],
  ));
  expect(pages(layout)).toEqual(expected);
});

it('keeps a widow-controlled two-line paragraph together beside a continuing note head', () => {
  // §17.3.1.44: B01 alone would orphan B02. Both lines (20) + band 6 +
  // fifteen note lines = 176 fit, so the head is planned for the pair.
  const layout = render(model(
    [paragraph(lines('B', 2), [['1']], { widowControl: true })],
    [[paragraph(lines('N', 20))]],
  ));
  expect(pages(layout)).toEqual([
    { body: ['1B01B02'], notes: [lines('N', 15).join('')] },
    { body: [], notes: [lines('N', 20).slice(15).join('')] },
  ]);
});

it.each([
  // §17.3.1.44 widow unit of two lines: B01 references a 20-line note, B02 a
  // one-line note. 20 + band 6 + fourteen head lines + 10 = 176 fit, so the
  // first plan must leave B02's note its line instead of orphaning B01.
  { name: 'widow-controlled two-line', bodyLines: 2, secondReferenceLine: 1,
    properties: { widowControl: true }, head: 14 },
  // §17.3.1.14 keepLines unit of ten lines with the second reference on the
  // final line: 100 + 6 + six head lines + 10 = 176.
  { name: 'ten-line keepLines', bodyLines: 10, secondReferenceLine: 9,
    properties: { keepLines: true }, head: 6 },
])('reserves a later reference inside the $name body unit before planning the first', ({
  bodyLines, secondReferenceLine, properties, head,
}) => {
  const references = Array.from({ length: bodyLines }, (_, line) =>
    line === 0 ? ['1'] : line === secondReferenceLine ? ['2'] : []);
  const layout = render(model(
    [paragraph(lines('B', bodyLines), references, properties)],
    [[paragraph(lines('N', 20))], [paragraph(['M01'])]],
  ));
  const bodyText = lines('B', bodyLines).map((line, index) =>
    (index === 0 ? '1' : index === secondReferenceLine ? '2' : '') + line).join('');
  expect(pages(layout)).toEqual([
    { body: [bodyText], notes: [lines('N', head).join(''), 'M01'] },
    { body: [], notes: [lines('N', 20).slice(head).join('')] },
  ]);
  expect(notesOnPage(layout.pages[0]!).map(note => note.source.storyInstance)).toEqual(['1', '2']);
  expect(notesOnPage(layout.pages[0]!).map(note => note.separator.length > 0)).toEqual([true, false]);
});

it('moves a reference line past pages filled by an earlier note\'s carried lines', () => {
  // 60pt body region. A01's 20-line note keeps four lines beside it (10 + 6 +
  // 40) and then fills each following page with its carried lines (6 + 50).
  // B01's own note cannot keep a real line on such a page, so B01 waits for a
  // page with room instead of being forced in with its whole 200pt note.
  const document = model(
    [paragraph(['A01'], [['1']]), paragraph(['B01'], [['2']])],
    [[paragraph(lines('N', 20))], [paragraph(lines('M', 20))]],
  );
  document.section.pageHeight = 80;
  const layout = render(document);
  const pageOf = (predicate: (page: DocumentLayout['pages'][number]) => boolean) =>
    layout.pages.findIndex(predicate);
  // Body order and both notes are conserved once, in source order.
  expect(layout.pages.flatMap(page => page.layers.body.map(text))).toEqual(['1A01', '2B01']);
  for (const [id, prefix] of [['1', 'N'], ['2', 'M']] as const) {
    expect(layout.pages.flatMap(page => notesOnPage(page)
      .filter(note => note.source.storyInstance === id)
      .map(note => note.story.blocks.map(text).join(''))).join('')).toBe(lines(prefix, 20).join(''));
  }
  // Each note starts on the page that paints its reference.
  expect(pageOf(page => page.layers.body.some(block => text(block) === '1A01')))
    .toBe(pageOf(page => notesOnPage(page).some(note => note.source.storyInstance === '1')));
  expect(pageOf(page => page.layers.body.some(block => text(block) === '2B01')))
    .toBe(pageOf(page => notesOnPage(page).some(note => note.source.storyInstance === '2')));
  // The page that receives B01 still begins its band with note 1's tail,
  // and every band owns exactly one separator, on its first note.
  const bPage = pageOf(page => page.layers.body.some(block => text(block) === '2B01'));
  expect(notesOnPage(layout.pages[bPage]!).map(note => note.source.storyInstance)).toEqual(['1', '2']);
  for (const page of layout.pages) {
    expect(notesOnPage(page).map(note => note.separator.length > 0))
      .toEqual(notesOnPage(page).map((_, index) => index === 0));
  }
});

it('moves a framed reference past a page whose band holds carried note lines', () => {
  // A framed (§17.3.1.11 framePr) paragraph is admitted whole with whole
  // notes: that is this library's continuation contract (unsupported framed
  // note continuation), not a framePr requirement. B01 + 6 + sixteen head lines
  // fill page 1 (176); page 2 opens with the 14 carried lines (6 + 140), so
  // only 34pt remain there, too little for the frame's whole 5-line note
  // (50). The band must stay inside the 10..190 body region, so the frame
  // and its note move to page 3 instead of pushing the band off the page.
  const framed = { ...paragraph(['F01'], [['2']]), framePr: {
    dropCap: 'none', lines: 1, wrap: 'around', hAnchor: 'text', vAnchor: 'text',
    hRule: 'auto', hSpace: 0, vSpace: 0, w: 40,
  } } as unknown as DocParagraph;
  const layout = render(model(
    [paragraph(['B01'], [['1']]), framed],
    [[paragraph(lines('N', 30))], [paragraph(lines('M', 5))]],
  ));
  expect(pages(layout)).toEqual([
    { body: ['1B01'], notes: [lines('N', 16).join('')] },
    { body: [], notes: [lines('N', 30).slice(16).join('')] },
    { body: ['2F01'], notes: [lines('M', 5).join('')] },
  ]);
  for (const page of layout.pages) {
    for (const note of notesOnPage(page)) {
      expect(note.flowBounds.yPt).toBeGreaterThanOrEqual(10);
      expect(note.flowBounds.yPt + note.flowBounds.heightPt).toBeLessThanOrEqual(190);
    }
  }
});

// §17.3.1.33 spacing after is spacing that follows the paragraph's last line,
// and the selected note band sits below the body (§17.11). Where that band
// follows the reference paragraph, the body ends at the band, not at a page
// edge, so a plan charges the authored spacing after beside it. Shapes are
// the public L01/L03/L04 controls (exact 10pt lines, the reference closing the
// paragraph's last line, exact separator stories); Word 16.113.3 allocated the
// same lines there.
// - L01: 36 lines + 6 after (366) + band 10 + 4 note lines (40) = 416 > 415,
//   so three head lines (406) stay and the fourth continues.
// - L03: the same 416 fits a 417pt region whole (the paired boundary).
// - L04: B01 + 100 after (110) + band 6 + 6 lines = 176 <= 180 < 186.
const referenceClosing = (texts: readonly string[], spaceAfter: number): DocParagraph => {
  const closing = paragraph(texts);
  return { ...closing, spaceAfter, runs: [...closing.runs,
    textRun('1', { noteRef: { kind: 'footnote', id: '1' } })] } as unknown as DocParagraph;
};
it.each([
  { shape: 'L01', regionPt: 415, bodyLines: 36, spaceAfter: 6, separatorPt: 10, noteLines: 4, head: 3 },
  { shape: 'L03', regionPt: 417, bodyLines: 36, spaceAfter: 6, separatorPt: 10, noteLines: 4, head: 4 },
  { shape: 'L04', regionPt: 180, bodyLines: 1, spaceAfter: 100, separatorPt: 6, noteLines: 8, head: 6 },
])('charges authored spacing after beside the note band ($shape)', ({
  regionPt, bodyLines, spaceAfter, separatorPt, noteLines, head,
}) => {
  const document = model(
    [referenceClosing(lines('B', bodyLines), spaceAfter)],
    [[paragraph(lines('N', noteLines))]],
  );
  document.section.pageHeight = regionPt + 20;
  (document as unknown as Record<string, unknown>).__noteLayoutSettings = {
    footnoteSeparator: 'short',
    footnoteContinuationSeparator: 'full',
    footnoteSeparatorStory: [paragraph([], [], {})],
    footnoteContinuationSeparatorStory: [paragraph([], [], {})],
  };
  for (const story of ['footnoteSeparatorStory', 'footnoteContinuationSeparatorStory']) {
    const settings = (document as unknown as { __noteLayoutSettings: Record<string, DocParagraph[]> }).__noteLayoutSettings;
    settings[story] = [{ ...settings[story]![0]!, lineSpacing: { value: separatorPt, rule: 'exact', explicit: true } } as DocParagraph];
  }
  const body = lines('B', bodyLines).join('') + '1';
  expect(pages(render(document))).toEqual(head === noteLines
    ? [{ body: [body], notes: [lines('N', noteLines).join('')] }]
    : [
        { body: [body], notes: [lines('N', head).join('')] },
        { body: [], notes: [lines('N', noteLines).slice(head).join('')] },
      ]);
});

/** Selected exact separator stories of `pt` for both footnote roles. */
function withExactSeparator(document: DocxDocumentModel, pt: number): DocxDocumentModel {
  const story = { ...paragraph([]), lineSpacing: { value: pt, rule: 'exact', explicit: true } };
  (document as unknown as Record<string, unknown>).__noteLayoutSettings = {
    footnoteSeparator: 'short', footnoteContinuationSeparator: 'full',
    footnoteSeparatorStory: [story], footnoteContinuationSeparatorStory: [story],
  };
  return document;
}

// Public Word 16.113.3 M/N controls (exact lines, widowControl off unless
// stated, exact 10pt separator story): F = 10 lines + 6 after (106); P = 25
// lines whose last line closes with reference 1, + `pAfter` after; note 1 has
// `noteLines` exact lines of `noteLinePt`. The whole note misses beside P only
// because of P's authored after (charged), fits without it, and fits a fresh
// page; nothing follows the reference. Word split such notes and kept them
// whole where the charged total fits. Arithmetic, region R:
// - M03/M01/M04/N06/N07: 106 + 256 + 50 = 412 > 410, 406 fits; head capacity
//   410 - 362 = 48 holds band 10 + 3 lines.
// - N01: 106 + 256 + 30 = 392 > 390, 386 fits; capacity 28 holds 10 + 1 line.
// - N02 (8pt note lines): 388 > 386, 382 fits; capacity 24 holds 10 + 8.
// - Whole: M02 412 <= 412; M07 (P after 0) 106 + 250 + 50 = 406 <= 406;
//   N04 392; N05 388.
// N02/N05 note text and paragraph mark are authored Arial 8pt with exact 8pt
// lines, as in the public originals; every other note is 10pt.
it.each([
  { id: 'M03 terminal', region: 410, pAfter: 6, noteLines: 4, head: 3 },
  { id: 'M01 later note', region: 410, pAfter: 6, noteLines: 4, head: 3, following: 'note' as const },
  { id: 'M04 following body', region: 410, pAfter: 6, noteLines: 4, head: 3, following: 'body' as const },
  { id: 'N06 widowControl', region: 410, pAfter: 6, noteLines: 4, head: 3, widowControl: true },
  { id: 'N07 keepLines', region: 410, pAfter: 6, noteLines: 4, head: 3, keepLines: true },
  { id: 'N01 two 10pt lines', region: 390, pAfter: 6, noteLines: 2, head: 1 },
  { id: 'N02 two 8pt lines', region: 386, pAfter: 6, noteLines: 2, noteLinePt: 8, noteFontPt: 8, head: 1 },
  { id: 'M02 charged fit', region: 412, pAfter: 6, noteLines: 4, head: 4, following: 'note' as const },
  { id: 'M07 no after', region: 406, pAfter: 0, noteLines: 4, head: 4, following: 'note' as const },
  { id: 'N04 charged fit', region: 392, pAfter: 6, noteLines: 2, head: 2 },
  { id: 'N05 charged 8pt fit', region: 388, pAfter: 6, noteLines: 2, noteLinePt: 8, noteFontPt: 8, head: 2 },
])('continues a note that fits beside its completing paragraph only without authored after ($id)', ({
  region, pAfter, noteLines, noteLinePt = 10, noteFontPt = 10, head, following, widowControl = false, keepLines = false,
}: {
  region: number; pAfter: number; noteLines: number; noteLinePt?: number; noteFontPt?: number; head: number;
  following?: 'note' | 'body'; widowControl?: boolean; keepLines?: boolean;
}) => {
  const f = { ...paragraph(lines('F', 10), [], { widowControl }), spaceAfter: 6 } as DocParagraph;
  const p = { ...referenceClosing(lines('P', 25), pAfter), widowControl, keepLines } as DocParagraph;
  const q = following === 'note' ? [paragraph(lines('Q', 10), [['2']])]
    : following === 'body' ? [paragraph(lines('Q', 10))] : [];
  const authored = paragraph(lines('N', noteLines));
  const note = { ...authored, lineSpacing: { value: noteLinePt, rule: 'exact', explicit: true },
    defaultFontSize: noteFontPt,
    runs: authored.runs.map(run => run.type === 'text' ? { ...run, fontSize: noteFontPt } : run),
  } as DocParagraph;
  const document = withExactSeparator(model([f, p, ...q],
    following === 'note' ? [[note], [paragraph(lines('M', 1))]] : [[note]]), 10);
  document.section.pageHeight = region + 20;
  const layout = render(document);
  if (head < noteLines) {
    // A real head keeps P's retained after clear of the band: the plan was
    // sized beside the charged body, not the discounted one.
    const reference = layout.pages[0]!.layers.body[1]!;
    if (reference.kind !== 'paragraph') throw new Error('Expected the reference paragraph');
    const band = notesOnPage(layout.pages[0]!)[0]!;
    expect(reference.flowBounds.yPt + reference.flowBounds.heightPt + reference.spacing.afterPt)
      .toBeLessThanOrEqual(band.flowBounds.yPt);
    expect(band.flowBounds.heightPt).toBe(10 + head * noteLinePt);
  }
  const qText = following === 'note' ? '2' + lines('Q', 10).join('') : lines('Q', 10).join('');
  const tail = head < noteLines ? [lines('N', noteLines).slice(head).join('')] : [];
  expect(pages(layout)).toEqual([
    { body: [lines('F', 10).join(''), lines('P', 25).join('') + '1'], notes: [lines('N', head).join('')] },
    ...(q.length > 0 || tail.length > 0
      ? [{ body: q.length > 0 ? [qText] : [], notes: [...tail, ...(following === 'note' ? ['M01'] : [])] }]
      : []),
  ]);
});

/** A paragraph whose last line closes with a reference to note `id`. */
const closingReference = (texts: readonly string[], spaceAfter: number, id: string): DocParagraph => {
  const closing = paragraph(texts);
  return { ...closing, spaceAfter, runs: [...closing.runs,
    textRun(id, { noteRef: { kind: 'footnote', id } })] } as unknown as DocParagraph;
};

// Public Word 16.113.3 carry controls (exact 10pt lines, 410pt region, exact
// 10pt separator and continuation separator stories). Page 1 is the M03
// after-only shape, so note 1 continues 3 + 1. Page 2 opens with that incoming
// tail (continuation separator 10 + one line = 20), then Q and B; B's last line
// closes with reference 2 and nothing follows it. Whole note 2 misses beside B
// even without B's after, and fits a fresh page:
// - C01: Q 10 + after 6, B 25 + after 6 (362): 362 + 20 + 40 = 422 > 410, 416
//   without B's after; head capacity 410 - 20 - 362 = 28 holds two lines.
// - C04: Q 11 + after 12, B 25 + after 6 (378): 418 > 410, 412 without after;
//   capacity 12 holds one line of a two-line note.
// - C05: Q 10, B 26, no after (360): 420 > 410; capacity 30 holds three lines.
// Word kept every body line on page 2 and split note 2 at those heads; where the
// charged total fits it kept note 2 whole (C02: 350 + 20 + 40 = 410; C03: 362 +
// 20 + 20 = 402).
it.each([
  { id: 'C01', qLines: 10, qAfter: 6, bLines: 25, bAfter: 6, note2: 4, head: 2 },
  { id: 'C04', qLines: 11, qAfter: 12, bLines: 25, bAfter: 6, note2: 2, head: 1 },
  { id: 'C05', qLines: 10, qAfter: 0, bLines: 26, bAfter: 0, note2: 4, head: 3 },
  { id: 'C02', qLines: 10, qAfter: 0, bLines: 25, bAfter: 0, note2: 4, head: 4 },
  { id: 'C03', qLines: 10, qAfter: 6, bLines: 25, bAfter: 6, note2: 2, head: 2 },
])('continues a closing reference on a page that opens with an incoming note tail ($id)', ({
  qLines, qAfter, bLines, bAfter, note2, head,
}) => {
  const f = { ...paragraph(lines('F', 10)), spaceAfter: 6 } as DocParagraph;
  const q = { ...paragraph(lines('Q', qLines)), spaceAfter: qAfter } as DocParagraph;
  const document = withExactSeparator(model(
    [f, closingReference(lines('P', 25), 6, '1'), q, closingReference(lines('B', bLines), bAfter, '2')],
    [[paragraph(lines('N', 4))], [paragraph(lines('M', note2))]],
  ), 10);
  document.section.pageHeight = 430;
  const layout = render(document);
  const second = layout.pages[1]!;
  const band = notesOnPage(second);
  // The incoming tail's reserve is charged exactly once beside the new head.
  expect(band.reduce((sum, note) => sum + note.flowBounds.heightPt, 0)).toBe(20 + head * 10);
  const closing = second.layers.body[1]!;
  if (closing.kind !== 'paragraph') throw new Error('Expected the closing paragraph');
  expect(closing.flowBounds.yPt + closing.flowBounds.heightPt + closing.spacing.afterPt)
    .toBeLessThanOrEqual(band[0]!.flowBounds.yPt);
  expect(pages(layout)).toEqual([
    { body: [lines('F', 10).join(''), lines('P', 25).join('') + '1'], notes: ['N01N02N03'] },
    { body: [lines('Q', qLines).join(''), lines('B', bLines).join('') + '2'],
      notes: ['N04', lines('M', head).join('')] },
    ...(head < note2 ? [{ body: [], notes: [lines('M', note2).slice(head).join('')] }] : []),
  ]);
});

it('keeps the existing whole-note exclusion on a page without an incoming tail', () => {
  // Counter-shape, no incoming tail: F + P with reference 1 closing P, 405pt
  // region. 106 + 256 + 50 = 412 and 406 without P's after both miss, the
  // whole note fits a fresh page, and no text or break follows: no
  // continuation permission applies, so the library's existing exclusion
  // keeps note 1 whole and moves the reference line. No Word observation of
  // this shape is claimed; it bounds the incoming-tail permission.
  const f = { ...paragraph(lines('F', 10)), spaceAfter: 6 } as DocParagraph;
  const document = withExactSeparator(model([f, closingReference(lines('P', 25), 6, '1')],
    [[paragraph(lines('N', 4))]]), 10);
  document.section.pageHeight = 425;
  expect(pages(render(document))).toEqual([
    { body: [lines('F', 10).join(''), lines('P', 25).slice(0, 24).join('')], notes: [] },
    { body: ['P251'], notes: ['N01N02N03N04'] },
  ]);
});

it('does not treat a carried line-free paragraph as an incoming note tail', () => {
  // Note 1 is four real lines followed by an empty (run-free) paragraph that
  // acquires no lines (its mark-only advance is 10). 416pt region, exact 10pt
  // separator stories. Page 1 (after-only class): 106 + 256 + whole 60 = 422 >
  // 416, 416 without P's after; head capacity 416 - 362 = 54 holds separator
  // 10 + the four lines, so only the empty paragraph continues. Page 2 opens
  // with that empty fragment (continuation separator 10 + 10 = 20), then Q
  // (10 lines) and B (26 lines, no after) closing with reference 2: 360 + 40
  // = 400 > 396, fits a fresh page, no text, break or after. A line-free
  // fragment is no retained tail line, so the existing exclusion keeps note 2
  // whole and moves B26 with it; the empty fragment itself stays retained.
  const f = { ...paragraph(lines('F', 10)), spaceAfter: 6 } as DocParagraph;
  const document = withExactSeparator(model(
    [f, closingReference(lines('P', 25), 6, '1'), paragraph(lines('Q', 10)), closingReference(lines('B', 26), 0, '2')],
    [[paragraph(lines('N', 4)), paragraph([])], [paragraph(lines('M', 4))]],
  ), 10);
  document.section.pageHeight = 436;
  const layout = render(document);
  expect(pages(layout)).toEqual([
    { body: [lines('F', 10).join(''), lines('P', 25).join('') + '1'], notes: ['N01N02N03N04'] },
    { body: [lines('Q', 10).join(''), lines('B', 26).slice(0, 25).join('')], notes: [''] },
    { body: ['B262'], notes: ['M01M02M03M04'] },
  ]);
  // The carried fragment is note 1's line-free paragraph, retained once.
  const carried = notesOnPage(layout.pages[1]!);
  expect(carried.map(note => [note.source.storyInstance, note.flowBounds.heightPt])).toEqual([['1', 20]]);
  expect(carried[0]!.story.blocks.map(block => block.kind === 'paragraph' ? block.lines.length : -1)).toEqual([0]);
});

it('leaves a one-line after-only note on the unchanged no-head fallback (N03)', () => {
  // N03: 106 + 256 + band 10 + one note line = 382 > 380, 376 fits, but the
  // charged head capacity (18) cannot hold band + one real line (20), so no
  // head exists to continue. That fallback is unchanged and pending an owner
  // decision; Word 16.113.3 moved P25 and the note to page 2 instead. The
  // case pins only that no separator-only or empty head is committed.
  const f = { ...paragraph(lines('F', 10)), spaceAfter: 6 } as DocParagraph;
  const document = withExactSeparator(model([f, referenceClosing(lines('P', 25), 6)],
    [[paragraph(lines('N', 1))]]), 10);
  document.section.pageHeight = 400;
  const layout = render(document);
  expect(pages(layout)).toEqual([
    { body: [lines('F', 10).join(''), lines('P', 25).join('') + '1'], notes: ['N01'] },
  ]);
  expect(notesOnPage(layout.pages[0]!)[0]!.flowBounds.heightPt).toBe(20);
});

it('continues a later completing reference of a widow unit under one reserved notice', () => {
  // Reserved separator wire with exact 10pt stories: separator,
  // continuation separator and a 10pt continuation notice. 144pt region:
  // F (60) + the two-line widow unit P (20 + 6 after). B01 opens note A (one
  // line) and B02 closes with note B (four lines). The unit's later minimum
  // for B stays the whole note (40): its only continuation permission is the
  // after-only fit, which depends on capacity the minimum cannot know. A is
  // charged whole (10 + 10). At B02: 26 + 20 + 40 = 86 > 84, and 80 fits
  // without the after, so B continues with the largest head beside the charged
  // body: 84 - 26 - 20 = 38 holds two lines + the one notice (30).
  const control = textRun('');
  const storyParagraph = { ...paragraph([]), paragraphMarkFontFacts: { fontFamily: 'Arial', fontSize: 10 } };
  const rule = (mark: 'short' | 'full') => ({
    rule: { mark, control: { run: control } },
    paragraph: { paragraph: storyParagraph, contentMark: { run: control } },
  });
  const p = referenceClosing(['B01', 'B02'], 6);
  const unit = { ...p, widowControl: true,
    runs: [textRun('1', { noteRef: { kind: 'footnote', id: '1' } }), ...p.runs.map(run =>
      (run as { noteRef?: { id: string } }).noteRef ? textRun('2', { noteRef: { kind: 'footnote', id: '2' } }) : run)],
  } as unknown as DocParagraph;
  const document = model([paragraph(lines('F', 6)), unit], [[paragraph(['A01'])], [paragraph(lines('N', 4))]]);
  document.section.pageHeight = 164;
  (document as unknown as Record<string, unknown>).__noteLayoutSettings = {
    footnoteSeparator: 'short', footnoteContinuationSeparator: 'full',
    reservedSeparators: { footnote: {
      separator: rule('short'), continuationSeparator: rule('full'),
      continuationNotice: { paragraph: { paragraph: storyParagraph, contentMark: { run: control } } },
    } },
  };
  const layout = render(document);
  expect(pages(layout)).toEqual([
    { body: [lines('F', 6).join(''), '1B01B022'], notes: ['A01', 'N01N02'] },
    { body: [], notes: ['N03N04'] },
  ]);
  const [a, b] = notesOnPage(layout.pages[0]!);
  // A keeps its whole plan; B's head carries the page's only notice.
  expect([a!.advancePt, b!.advancePt]).toEqual([20, 30]);
  expect([a!.trailing, b!.trailing?.role]).toEqual([undefined, 'continuationNotice']);
  expect(notesOnPage(layout.pages[1]!).map(note => [note.leading?.role, note.trailing])).toEqual([
    ['continuationSeparator', undefined],
  ]);
  const reference = layout.pages[0]!.layers.body[1]!;
  if (reference.kind !== 'paragraph') throw new Error('Expected the widow unit');
  expect(reference.lines).toHaveLength(2);
  expect(reference.flowBounds.yPt + reference.flowBounds.heightPt + reference.spacing.afterPt)
    .toBeLessThanOrEqual(a!.flowBounds.yPt);
});

it('plans a custom-mark reference carried by a zero-width anchor host', () => {
  // 14 lines (140) leave 40pt; the suppressed mark still anchors note 1 to
  // B01, so the head is band 6 + two lines and B02 follows the tail.
  const body = paragraph(lines('B', 2), [['1']]);
  const host = body.runs[0] as unknown as { text: string; noteRef: { customMarkFollows: boolean } };
  host.text = '';
  host.noteRef.customMarkFollows = true;
  body.runs.splice(1, 0, textRun('*') as unknown as DocParagraph['runs'][number]);
  const layout = render(model([paragraph(lines('P', 14)), body], [[paragraph(lines('N', 4))]]));
  expect(pages(layout)).toEqual([
    { body: [lines('P', 14).join(''), '*B01'], notes: ['N01N02'] },
    { body: ['B02'], notes: ['N03N04'] },
  ]);
  const hosts = layout.pages[0]!.layers.body.flatMap(block => block.kind === 'paragraph'
    ? block.lines.flatMap(line => line.placements.flatMap(placement =>
      placement.kind === 'anchor-host' && placement.noteReference ? [placement.noteReference.id] : []))
    : []);
  expect(hosts).toEqual(['1']);
});

it('charges page-dependent note acquisitions only up to each page\'s rejection frontier', () => {
  // Each of 120 lines references its own one-line note whose PAGE field makes
  // the acquisition destination-dependent, so every acquisition is charged to
  // the cumulative footnote work budget. A page holds eight lines and notes
  // (80 + 6 + 80 = 166); the ninth reference line is that page's frontier.
  // Planning references beyond the frontier on every page attempt would
  // charge the remaining paragraph again per page (quadratic in the number
  // of references) and exhaust the budget; page-local planning stays linear.
  const count = 120;
  const ids = Array.from({ length: count }, (_, index) => String(index + 1));
  const notes = ids.map((_, index) => {
    const note = paragraph([label('Q', index, 3)]);
    note.runs.push({ ...textRun(''), type: 'field', fieldType: 'page',
      instruction: 'PAGE' + ' '.repeat(50_000), fallbackText: '1' } as unknown as DocParagraph['runs'][number]);
    return [note];
  });
  const layout = render(model([paragraph(lines('B', count, 3), ids.map(id => [id]))], notes));
  expect(layout.pages).toHaveLength(15);
  expect(notesOnPage(layout.pages[0]!).map(note => note.source.storyInstance)).toEqual(ids.slice(0, 8));
  expect(layout.pages.flatMap(page => notesOnPage(page).map(note => note.source.storyInstance))).toEqual(ids);
  expect(layout.pages.flatMap(page => page.layers.body.map(text)).join(''))
    .toBe(ids.map((id, index) => id + label('B', index, 3)).join(''));
});

import { readFile } from 'node:fs/promises';
import { beforeAll, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import type { DocxDocumentModel } from './types.js';
import type { DocumentLayout, NoteLayout, PaintNode } from './layout/types.js';

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

// Exact authored line heights isolate admission and source ownership. These
// synthetic glyph metrics are not an Office pagination or font-fidelity oracle.
// Listed exact 6pt separator stories keep the band at 6pt (Word's default
// separator would be one Normal line; WORD_NOTE_SEPARATOR_GEOMETRY).
const context = {
  font: '10px Arial', letterSpacing: '0px', fontKerning: 'none',
  measureText: (text: string) => ({
    width: [...text].length * 6,
    fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
    actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
  }),
} as unknown as CanvasRenderingContext2D;

/** `referenceLines[i]` is the body line (0-based) that references note i+1;
 * every reference opens the first line unless stated. */
function control(
  bodyLines = 2,
  noteLines = [4],
  precedingLines = 14,
  remainingTwips = 800,
  referenceLines: readonly number[] = noteLines.map(() => 0),
): DocxDocumentModel {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const pPr = '<w:pPr><w:spacing w:before="0" w:after="0" w:line="200" w:lineRule="exact"/><w:widowControl w:val="0"/></w:pPr>';
  const rPr = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr>';
  const paragraph = (prefix: string, count: number, reference = false) => '<w:p>' + pPr
    + Array.from({ length: count }, (_, index) => (
      (reference ? noteLines.flatMap((_, note) => referenceLines[note] === index
        ? [`<w:r>${rPr}<w:footnoteReference w:id="${note + 1}"/></w:r>`] : []).join('') : '')
      + `<w:r>${rPr}<w:t>${prefix}${String(index + 1).padStart(2, '0')}</w:t></w:r>`
      + (index + 1 < count ? '<w:r><w:br/></w:r>' : '')
    )).join('') + '</w:p>';
  const separatorPPr = '<w:pPr><w:spacing w:before="0" w:after="0" w:line="120" w:lineRule="exact"/></w:pPr>';
  const separatorStories = `<w:footnote w:type="separator" w:id="-1"><w:p>${separatorPPr}<w:r><w:separator/></w:r></w:p></w:footnote>`
    + `<w:footnote w:type="continuationSeparator" w:id="0"><w:p>${separatorPPr}<w:r><w:continuationSeparator/></w:r></w:p></w:footnote>`;
  const files = new Map([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'],
    ['_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="notes" Type="${R}/footnotes" Target="footnotes.xml"/><Relationship Id="settings" Type="${R}/settings" Target="settings.xml"/></Relationships>`],
    ['word/settings.xml', `<w:settings xmlns:w="${W}"><w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr></w:settings>`],
    ['word/document.xml', `<w:document xmlns:w="${W}"><w:body>${paragraph('P', precedingLines)}${paragraph('B', bodyLines, true)}<w:sectPr><w:pgSz w:w="4000" w:h="${precedingLines * 200 + remainingTwips + 400}"/><w:pgMar w:left="200" w:right="200" w:top="200" w:bottom="200"/></w:sectPr></w:body></w:document>`],
    ['word/footnotes.xml', `<w:footnotes xmlns:w="${W}">${separatorStories}${noteLines.map((count, index) => `<w:footnote w:id="${index + 1}">${paragraph(`N${index + 1}L`, count)}</w:footnote>`).join('')}</w:footnotes>`],
  ]);
  const archive = new DocxArchive(storeZip(new Map([...files].map(([name, xml]) => [name, new TextEncoder().encode(xml)]))));
  try {
    return normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally {
    archive.free();
  }
}

function text(block: PaintNode): string {
  return block.kind === 'paragraph'
    ? block.lines.flatMap(line => line.placements.filter(run => run.kind === 'text').map(run => run.text)).join('')
    : '';
}

function notesByPage(layout: DocumentLayout): string[][] {
  return layout.pages.map(page => notesOnPage(page).map(note => note.story.blocks.map(text).join('')));
}

function bodyByPage(layout: DocumentLayout): string[][] {
  return layout.pages.map(page => page.layers.body.map(text));
}

function notesOnPage(page: DocumentLayout['pages'][number]): NoteLayout[] {
  return page.layers.notes.filter((node): node is NoteLayout => node.kind === 'note');
}

function render(model: DocxDocumentModel, continuation = true): DocumentLayout {
  return layoutDocument(model, createLayoutServices(model, {
    measureContext: context, allowFootnoteContinuation: continuation,
  }), { currentDateMs: 0 });
}

const sequence = (prefix: string, count: number) => Array.from({ length: count }, (_, index) =>
  prefix + String(index + 1).padStart(2, '0')).join('');

function expectConserved(layout: DocumentLayout, bodyText: string, noteLines: number[]): void {
  expect(layout.pages.flatMap(page => page.layers.body.map(text)).join('')).toBe(bodyText);
  for (const [index, count] of noteLines.entries()) {
    expect(layout.pages.flatMap(page => notesOnPage(page)
      .filter(note => note.source.storyInstance === String(index + 1))
      .flatMap(note => note.story.blocks.map(text))).join('')).toBe(sequence(`N${index + 1}L`, count));
  }
  for (const page of layout.pages) {
    // One separator per page band, owned by the band's first note.
    expect(notesOnPage(page).map(note => note.separator.length > 0))
      .toEqual(notesOnPage(page).map((_, index) => index === 0));
    for (const note of notesOnPage(page)) {
      expect(note.flowDomainId).toBe(`notes:page:${page.pageIndex}`);
      for (const block of note.story.blocks) {
        expect(block.flowDomainId).toBe(`notes:page:${page.pageIndex}:footnote:${note.source.storyInstance}`);
      }
    }
  }
}

it.each([
  // Reference line 10 + 6pt band + two note lines = 36pt fits exactly; B02
  // (46pt) cannot displace the committed head. One twip less keeps one line.
  // Same allocation shape as the Word 16.113.3 B720/B719 source-open controls.
  { remainingTwips: 720, heads: ['N1L01N1L02', 'N1L03N1L04'] },
  { remainingTwips: 719, heads: ['N1L01', 'N1L02N1L03N1L04'] },
])('commits the reference-line note head at the $remainingTwips twip boundary', ({ remainingTwips, heads }) => {
  const layout = render(control(2, [4], 14, remainingTwips));
  expect(bodyByPage(layout)).toEqual([[sequence('P', 14), '1B01'], ['B02']]);
  expect(notesByPage(layout)).toEqual([[heads[0]], [heads[1]]]);
  expectConserved(layout, sequence('P', 14) + '1B01B02', [4]);
});

it('charges the preceding new note whole before partitioning the final referenced note', () => {
  const layout = render(control(2, [1, 4], 12, 1200));
  // 60pt remains: B01 10 + band 6 + N1 10 + three N2 lines 30 = 56; B02 does
  // not fit beside that plan.
  expect(bodyByPage(layout)).toEqual([[sequence('P', 12), '12B01'], ['B02']]);
  expect(notesByPage(layout)).toEqual([['N1L01', 'N2L01N2L02N2L03'], ['N2L04']]);
  expectConserved(layout, sequence('P', 12) + '12B01B02', [1, 4]);
});

it('keeps continuation disabled when the whole note cannot fit with the reference line', () => {
  const layout = render(control(), false);
  expect(layout.pages[0]?.layers.notes).toEqual([]);
  expect(layout.pages[1]?.layers.body.map(text)).toEqual(['1B01B02']);
  expect(notesByPage(layout)).toEqual([[], ['N1L01N1L02N1L03N1L04']]);
  expectConserved(layout, sequence('P', 14) + '1B01B02', [4]);
});

it('partitions a note at its reference line while the paragraph continues', () => {
  const layout = render(control(14));
  // Allocation matches the Word 16.113.3 PUBLIC4 source-open control (two
  // pages, B01 with the note head, B02-B14 with the tail); the two-line head
  // follows from the library's 6pt missing-story band, not from Word.
  expect(bodyByPage(layout)).toEqual([[sequence('P', 14), '1B01'], [sequence('B', 14).slice(3)]]);
  expect(notesByPage(layout)).toEqual([['N1L01N1L02'], ['N1L03N1L04']]);
  expectConserved(layout, sequence('P', 14) + '1' + sequence('B', 14), [4]);
});

it('keeps an exactly fitting full note whole without a pending tail', () => {
  const layout = render(control(2, [2], 14, 920));
  // 46pt is exactly the20pt complete body plus26pt full note.
  expect(layout.pages).toHaveLength(1);
  expect(notesByPage(layout)).toEqual([['N1L01N1L02']]);
  expectConserved(layout, sequence('P', 14) + '1B01B02', [2]);
});

it.each([
  // B01 commits note 1 whole (10 + 6 + 20); B02 fits beside it. At B03 the
  // second reference needs one real line: 30 + 26 + 10 = 66pt.
  {
    remainingTwips: 1320,
    body: [[sequence('P', 12), '1B01B022B03'], []],
    notes: [['N1L01N1L02', 'N2L01'], ['N2L02N2L03N2L04']],
  },
  // One twip less: the B03 plan cannot place a real line, so B03 and its
  // reference move together; the committed note 1 plan is retained intact.
  {
    remainingTwips: 1319,
    body: [[sequence('P', 12), '1B01B02'], ['2B03']],
    notes: [['N1L01N1L02'], ['N2L01N2L02N2L03N2L04']],
  },
])('admits a later reference line only against the earlier committed plan at $remainingTwips twips', ({
  remainingTwips, body, notes,
}) => {
  const layout = render(control(3, [2, 4], 12, remainingTwips, [0, 2]));
  expect(bodyByPage(layout)).toEqual(body);
  expect(notesByPage(layout)).toEqual(notes);
  expectConserved(layout, sequence('P', 12) + '1B01B022B03', [2, 4]);
});

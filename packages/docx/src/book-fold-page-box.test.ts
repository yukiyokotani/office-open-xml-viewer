import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import type { DocumentLayout } from './layout/types.js';
import { makeRecordingCanvas } from './test-support/recording-canvas.test-support.js';
import type { DocxDocumentModel } from './types.js';

// ECMA-376 Part 1 §17.15.1.11 bookFoldPrinting / §17.15.1.13
// bookFoldRevPrinting through the production parser and layout. Word saves the
// folded editing page in pgSz (§17.6.13), so the authored box is the page
// (`word-book-fold-authored-page-box`); the gutter stays at the bisector edge
// (`word-book-fold-gutter-right-edge`).

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function layout(settings: string, gutterTwips: number): DocumentLayout {
  const words = Array.from({ length: 400 }, (_, i) => `word${i}`).join(' ');
  // A half-Legal booklet page as Word saves it: 7in x 8.5in, orient naming
  // the landscape sheet, 0.5in margins.
  const document = `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body><w:p><w:r><w:t>${words}</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="10080" w:h="12240" w:orient="landscape" w:code="5"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="${gutterTwips}"/></w:sectPr></w:body></w:document>`;
  const files = new Map<string, string>([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'],
    ['_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rSettings" Type="${R}/settings" Target="settings.xml"/></Relationships>`],
    ['word/document.xml', document],
    ['word/settings.xml', `<w:settings xmlns:w="${W}">${settings}</w:settings>`],
  ]);
  const archive = new DocxArchive(storeZip(new Map([...files].map(([name, data]) => [
    name, new TextEncoder().encode(data),
  ]))));
  let model: DocxDocumentModel;
  try { model = JSON.parse(new TextDecoder().decode(archive.parse())) as DocxDocumentModel; }
  finally { archive.free(); }
  const ctx = makeRecordingCanvas().canvas.getContext('2d') as CanvasRenderingContext2D;
  return layoutDocument(model, createLayoutServices(model, { measureContext: ctx }), { currentDateMs: 0 });
}

describe('book-fold page box', () => {
  it.each([
    ['bookFoldPrinting', '<w:bookFoldPrinting/>'],
    ['bookFoldRevPrinting', '<w:bookFoldPrinting/><w:bookFoldRevPrinting/>'],
  ])('lays out a book-fold section on its authored pgSz page box (%s)', (_label, settings) => {
    const plain = layout('', 0);
    const folded = layout(settings, 0);
    const page = folded.pages[0];

    expect(page.geometry).toMatchObject({ widthPt: 504, heightPt: 612 });
    expect(page.section.columns).toEqual([{ xPt: 36, wPt: 432 }]);
    // Without a gutter the fold changes nothing about the page or its flow.
    expect(folded.pages.length).toBe(plain.pages.length);
  });

  it('adds the authored gutter at the bisector (right) edge of the unhalved page', () => {
    const page = layout('<w:bookFoldPrinting/>', 360).pages[0];

    expect(page.geometry).toMatchObject({ widthPt: 504, heightPt: 612 });
    expect(page.section.columns).toEqual([{ xPt: 36, wPt: 414 }]);
  });
});

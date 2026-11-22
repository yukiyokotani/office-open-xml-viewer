import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import type { DocumentLayout } from './layout/types.js';
import { makeRecordingCanvas } from './test-support/recording-canvas.test-support.js';
import type { DocxDocumentModel } from './types.js';

// ECMA-376 Part 1 §17.15.1.11/.13 (book fold) and §17.15.1.64 (two pages per
// sheet) through the production parser and layout. Word keeps the authored
// pgSz as each editing page and reads pgSz/@orient (§17.6.13) as the sheet
// arrangement (`word-two-page-sheet-imposition`). Margins are left 90, right
// 27, top 45, bottom 63 with a 21.6pt gutter, as in the Word controls.

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function layout(settings: string, pgSz: string, sectionExtra = ''): DocumentLayout {
  const paragraph = (text: string, pageBreak = false) =>
    `<w:p>${pageBreak ? '<w:r><w:br w:type="page"/></w:r>' : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`;
  const document = `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${paragraph('first')}${paragraph('second', true)}<w:sectPr><w:headerReference w:type="default" r:id="rHdr"/>${pgSz}<w:pgMar w:top="900" w:right="540" w:bottom="1260" w:left="1800" w:header="360" w:footer="360" w:gutter="432"/>${sectionExtra}</w:sectPr></w:body></w:document>`;
  const files = new Map<string, string>([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'],
    ['_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rSettings" Type="${R}/settings" Target="settings.xml"/><Relationship Id="rHdr" Type="${R}/header" Target="header1.xml"/></Relationships>`],
    ['word/document.xml', document],
    ['word/settings.xml', `<w:settings xmlns:w="${W}">${settings}</w:settings>`],
    ['word/header1.xml', `<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>`],
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

/** Physical [body left, body right, body top, body bottom, header left,
 * header right] of a page. */
function pageMargins(result: DocumentLayout, pageIndex: number): number[] {
  const page = result.pages[pageIndex]!;
  const { widthPt, heightPt } = page.geometry;
  const body = page.flowDomains.find((domain) => domain.kind === 'body')!.physicalBounds;
  const header = page.flowDomains.find((domain) => domain.id === `story:header:page:${pageIndex}`)!;
  return [
    body.xPt,
    widthPt - body.xPt - body.widthPt,
    body.yPt,
    heightPt - body.yPt - body.heightPt,
    header.physicalBounds.xPt,
    widthPt - header.physicalBounds.xPt - header.physicalBounds.widthPt,
  ].map((value) => Math.round(value * 100) / 100);
}

const LANDSCAPE_SHEET = '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>';
const PORTRAIT_SHEET = '<w:pgSz w:w="12240" w:h="15840"/>';
const TBRL = '<w:textDirection w:val="tbRl"/>';
const TBLRV = '<w:textDirection w:val="tbLrV"/>';

describe('two-page sheet imposition', () => {
  it.each([
    ['book fold on a landscape sheet', '<w:bookFoldPrinting/>', LANDSCAPE_SHEET, [792, 612],
      [48.6, 90, 45, 63, 111.6, 27], [90, 48.6, 45, 63, 111.6, 27]],
    ['book fold on a half-sheet page saved with orient="landscape"', '<w:bookFoldPrinting/>',
      '<w:pgSz w:w="10080" w:h="12240" w:orient="landscape" w:code="5"/>', [504, 612],
      [48.6, 90, 45, 63, 111.6, 27], [90, 48.6, 45, 63, 111.6, 27]],
    ['book fold on a portrait sheet', '<w:bookFoldPrinting/>', PORTRAIT_SHEET, [612, 792],
      [90, 27, 45, 63, 111.6, 27], [90, 27, 45, 63, 111.6, 27]],
    ['two-up on a landscape sheet', '<w:printTwoOnOne/>', LANDSCAPE_SHEET, [792, 612],
      [48.6, 90, 45, 63, 48.6, 90], [90, 48.6, 45, 63, 90, 48.6]],
    ['two-up on a portrait sheet', '<w:printTwoOnOne/>', PORTRAIT_SHEET, [612, 792],
      [90, 27, 84.6, 45, 90, 27], [90, 27, 45, 84.6, 90, 27]],
    // A tbRl section binds on the right; its physical header and footer take
    // the same imposed margins as a horizontal section's.
    ['a tbRl section under book fold', '<w:bookFoldPrinting/>', LANDSCAPE_SHEET, [792, 612],
      [90, 48.6, 45, 63, 90, 48.6], [90, 48.6, 45, 63, 90, 48.6], TBRL],
    ['a tbRl section two-up on a landscape sheet', '<w:printTwoOnOne/>', LANDSCAPE_SHEET, [792, 612],
      [90, 48.6, 45, 63, 90, 48.6], [48.6, 90, 45, 63, 48.6, 90], TBRL],
    // Body and header margins differ here (the header ignores gutterAtTop), so
    // this pins the vertical page-story path to the header/footer margins.
    ['a tbRl section two-up with gutterAtTop', '<w:gutterAtTop/><w:printTwoOnOne/>', LANDSCAPE_SHEET,
      [792, 612], [90, 27, 66.6, 63, 90, 48.6], [48.6, 68.4, 66.6, 63, 48.6, 90], TBRL],
    // A tbLrV section binds on the right like tbRl, although its lines start
    // at the left margin.
    ['a tbLrV section on single pages', '', PORTRAIT_SHEET, [612, 792],
      [90, 48.6, 45, 63, 90, 48.6], [90, 48.6, 45, 63, 90, 48.6], TBLRV],
    ['a tbLrV section with mirrored margins', '<w:mirrorMargins/>', PORTRAIT_SHEET, [612, 792],
      [90, 48.6, 45, 63, 90, 48.6], [48.6, 90, 45, 63, 48.6, 90], TBLRV],
    ['a tbLrV section with gutterAtTop', '<w:gutterAtTop/>', PORTRAIT_SHEET, [612, 792],
      [90, 27, 66.6, 63, 90, 27], [90, 27, 66.6, 63, 90, 27], TBLRV],
    ['a tbLrV section under book fold', '<w:bookFoldPrinting/>', LANDSCAPE_SHEET, [792, 612],
      [90, 48.6, 45, 63, 90, 48.6], [90, 48.6, 45, 63, 90, 48.6], TBLRV],
    ['a tbLrV section two-up on a landscape sheet', '<w:printTwoOnOne/>', LANDSCAPE_SHEET, [792, 612],
      [90, 48.6, 45, 63, 90, 48.6], [48.6, 90, 45, 63, 48.6, 90], TBLRV],
  ])('lays out %s on the authored page box with Word imposition margins', (
    _label, settings, pgSz, size, first, second, sectionExtra = '',
  ) => {
    const result = layout(settings, pgSz, sectionExtra);

    expect(result.pages).toHaveLength(2);
    for (const page of result.pages) {
      expect([page.geometry.widthPt, page.geometry.heightPt]).toEqual(size);
    }
    expect(pageMargins(result, 0)).toEqual(first);
    expect(pageMargins(result, 1)).toEqual(second);
  });
});

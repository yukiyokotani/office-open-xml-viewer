/// <reference types="node" />

import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { textRunGeometryForPage } from './layout/text-index.js';
import { storeZip } from './conformance/generate.js';
import type { DocumentLayout } from './layout/types.js';
import type { DocxDocumentModel } from './types.js';

// ECMA-376 CT_Tbl allows a table with no w:tr (EG_ContentRowContent is
// minOccurs="0"); template engines emit one when a repeated row is bound to
// an empty list, and rows wrapped in an empty w:sdt flatten to none.
// §17.4.37 (tbl) defines a table as content in rows, so such a table has
// nothing to lay out. Through the real parser, in every story that takes a
// table, the document must lay out in finite time and exactly as without the
// table: same pages, every run in place, no table-boundary paragraph spacing,
// no float exclusion, no logical-table sequence barrier, no Cover Pages tail
// and no REF bookmark barrier. Synthetic linear metrics are a test
// measurement service only.

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const PACKAGE_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const EMU = 12700;
const TW = 20;

const encoder = new TextEncoder();
const xml = (body: string) => encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${body}`);

const exact24 = `<w:spacing w:before="0" w:after="0" w:line="${24 * TW}" w:lineRule="exact"/>`;
const para = (text: string) => `<w:p><w:pPr>${exact24}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;

/** Fixed-layout table over `widthsPt`, one row of `cells`, top/bottom
 * margins `marginPt`, left/right 0, optional 1pt single rules. */
function table(cells: readonly string[], widthsPt: readonly number[], marginPt: number, ruled: boolean): string {
  const rule = (side: string) => `<w:${side} w:val="single" w:sz="8" w:space="0" w:color="000000"/>`;
  const borders = ruled
    ? `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideV'].map(rule).join('')}</w:tblBorders>`
    : '';
  const total = widthsPt.reduce((sum, width) => sum + width, 0);
  return `<w:tbl><w:tblPr><w:tblW w:w="${total * TW}" w:type="dxa"/>${borders}<w:tblLayout w:type="fixed"/>`
    + `<w:tblCellMar><w:top w:w="${marginPt * TW}" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>`
    + `<w:bottom w:w="${marginPt * TW}" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>`
    + `<w:tblGrid>${widthsPt.map((width) => `<w:gridCol w:w="${width * TW}"/>`).join('')}</w:tblGrid><w:tr>`
    + cells.map((content, index) =>
      `<w:tc><w:tcPr><w:tcW w:w="${widthsPt[index]! * TW}" w:type="dxa"/></w:tcPr>${content}</w:tc>`).join('')
    + '</w:tr></w:tbl>';
}

/** A 320pt borderless margin-free outer cell holding PARENT, `inner`, and
 * PARENT-AFTER (the r2 containment topology). */
const nestedIn = (inner: string) => table([para('PARENT') + inner + para('PARENT-AFTER')], [320], 0, false);

type Stories = Readonly<{
  body: string;
  header?: string;
  footer?: string;
  footnote?: string;
  endnote?: string;
  /** §17.6.20 section `<w:textDirection>`; absent is horizontal. */
  textDirection?: 'tbRl';
}>;

function docx(stories: Stories): Uint8Array {
  const parts = [
    ['header', 'header1.xml', 'hdr', stories.header],
    ['footer', 'footer1.xml', 'ftr', stories.footer],
    ['footnotes', 'footnotes.xml', 'footnotes', stories.footnote],
    ['endnotes', 'endnotes.xml', 'endnotes', stories.endnote],
  ] as const;
  const present = parts.filter(([, , , content]) => content !== undefined);
  const references = (stories.header !== undefined ? '<w:headerReference w:type="default" r:id="rIdheader"/>' : '')
    + (stories.footer !== undefined ? '<w:footerReference w:type="default" r:id="rIdfooter"/>' : '');
  const document = `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:wps="${WPS}">`
    + `<w:body>${stories.body}<w:sectPr>${references}<w:pgSz w:w="${400 * TW}" w:h="${400 * TW}"/>`
    + `<w:pgMar w:top="${40 * TW}" w:right="${40 * TW}" w:bottom="${40 * TW}" w:left="${40 * TW}" `
    + `w:header="${20 * TW}" w:footer="${20 * TW}" w:gutter="0"/>`
    + (stories.textDirection ? `<w:textDirection w:val="${stories.textDirection}"/>` : '')
    + '</w:sectPr></w:body></w:document>';
  const styles = `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>`
    + '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults>'
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    // A default table style gives adjacent tables one §17.4.37 identity.
    + '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style></w:styles>';
  const files = new Map<string, Uint8Array>([
    ['[Content_Types].xml', xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
      + `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + `<Override PartName="/word/document.xml" ContentType="${CONTENT_TYPE}.document.main+xml"/>`
      + `<Override PartName="/word/styles.xml" ContentType="${CONTENT_TYPE}.styles+xml"/>`
      + present.map(([type, file]) =>
        `<Override PartName="/word/${file}" ContentType="${CONTENT_TYPE}.${type}+xml"/>`).join('')
      + '</Types>')],
    ['_rels/.rels', xml(`<Relationships xmlns="${PACKAGE_RELS}">`
      + `<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/_rels/document.xml.rels', xml(`<Relationships xmlns="${PACKAGE_RELS}">`
      + `<Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/>`
      + present.map(([type, file]) => `<Relationship Id="rId${type}" Type="${R}/${type}" Target="${file}"/>`).join('')
      + '</Relationships>')],
    ['word/document.xml', xml(document)],
    ['word/styles.xml', xml(styles)],
  ]);
  for (const [type, file, root, content] of present) {
    const inner = type === 'footnotes' ? `<w:footnote w:id="1">${content}</w:footnote>`
      : type === 'endnotes' ? `<w:endnote w:id="1">${content}</w:endnote>`
        : content;
    const element = type === 'footnotes' || type === 'endnotes' ? `w:${type}` : `w:${root}`;
    files.set(`word/${file}`, xml(`<${element} xmlns:w="${W}">${inner}</${element}>`));
  }
  return storeZip(files);
}

const noteReference = (kind: 'footnote' | 'endnote') =>
  `<w:p><w:pPr>${exact24}</w:pPr><w:r><w:t>REF</w:t></w:r><w:r><w:${kind}Reference w:id="1"/></w:r></w:p>`;

/** A body paragraph anchoring a page-positioned (50, 100) 300×250 wrap-none
 * text box, turned `degrees`, whose story (`vert`) holds `content`. */
function textBox(degrees: number, vert: 'horz' | 'wordArtVert', content: string): string {
  const cx = 300 * EMU;
  const cy = 250 * EMU;
  return '<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
    + 'relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
    + `<wp:positionH relativeFrom="page"><wp:posOffset>${50 * EMU}</wp:posOffset></wp:positionH>`
    + `<wp:positionV relativeFrom="page"><wp:posOffset>${100 * EMU}</wp:posOffset></wp:positionV>`
    + `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>`
    + `<wp:docPr id="1" name="Box"/><a:graphic><a:graphicData uri="${WPS}"><wps:wsp><wps:cNvSpPr txBox="1"/>`
    + `<wps:spPr><a:xfrm rot="${degrees * 60000}"><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`
    + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>'
    + `<wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx>`
    + `<wps:bodyPr rot="0" vert="${vert}" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t">`
    + '<a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>';
}

function measureContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText: (text: string) => ({
      width: [...text].length * px() * 0.5,
      actualBoundingBoxAscent: px() * 0.8,
      actualBoundingBoxDescent: px() * 0.2,
      fontBoundingBoxAscent: px() * 0.8,
      fontBoundingBoxDescent: px() * 0.2,
    } as TextMetrics),
  } as unknown as CanvasRenderingContext2D;
}

function layout(stories: Stories): DocumentLayout {
  const archive = new DocxArchive(docx(stories));
  let model: DocxDocumentModel;
  try {
    model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally {
    archive.free();
  }
  return layoutDocument(model, createLayoutServices(model, { measureContext: measureContext() }), {
    currentDateMs: 0,
  });
}

function observed(result: DocumentLayout) {
  return {
    pageCount: result.pages.length,
    runs: result.pages.flatMap((_, pageIndex) => textRunGeometryForPage(result, pageIndex).map((run) => ({
      pageIndex,
      text: run.placement.text,
      bounds: run.placement.bounds,
      pointToPage: run.pointToPage,
    }))),
    diagnostics: result.diagnostics,
  };
}

const GRID = '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>';
const WIDTH = '<w:tblW w:w="4000" w:type="dxa"/>';
const ZERO_ROW_TABLES = [
  ['in-flow', `<w:tbl><w:tblPr>${WIDTH}</w:tblPr>${GRID}</w:tbl>`],
  ['text-anchored floating', `<w:tbl><w:tblPr><w:tblpPr w:vertAnchor="text" w:horzAnchor="margin" w:tblpY="20"/>${WIDTH}</w:tblPr>${GRID}</w:tbl>`],
  ['page-anchored floating', `<w:tbl><w:tblPr><w:tblpPr w:vertAnchor="page" w:horzAnchor="page" w:tblpX="1000" w:tblpY="1000"/>${WIDTH}</w:tblPr>${GRID}</w:tbl>`],
  ['grid-free', '<w:tbl><w:tblPr/><w:tblGrid/></w:tbl>'],
  ['text-anchored floating, wide text distances', `<w:tbl><w:tblPr><w:tblpPr w:leftFromText="0" w:rightFromText="0" w:topFromText="0" w:bottomFromText="800" w:vertAnchor="text" w:horzAnchor="text" w:tblpY="0"/>${WIDTH}</w:tblPr>${GRID}</w:tbl>`],
  ['empty-sdt', `<w:tbl><w:tblPr/>${GRID}<w:sdt><w:sdtContent/></w:sdt><w:bookmarkStart w:id="0" w:name="b"/><w:bookmarkEnd w:id="0"/></w:tbl>`],
] as const;

const P = (text: string) => para(text);
const SPACED = (text: string) =>
  `<w:p><w:pPr><w:spacing w:before="240" w:after="360"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
const ONE_ROW = table([P('CELL')], [200], 0, true);
const ONE_ROW_FLOATING = ONE_ROW.replace('<w:tblPr>', '<w:tblPr><w:tblpPr w:vertAnchor="text" w:horzAnchor="margin" w:tblpY="0"/>');

const cover = (content: string) => '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Cover Pages"/>'
  + `</w:docPartObj></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
/** A bookmark from a leading-page-break paragraph past `t` to a plain one,
 * and a REF field to it, whose result takes the bookmarked leading break. */
const refBookmark = (t: string) =>
  `<w:p><w:pPr>${exact24}</w:pPr><w:bookmarkStart w:id="7" w:name="Anchor"/><w:r><w:br w:type="page"/></w:r>`
  + `<w:r><w:t>X</w:t></w:r></w:p>${t}<w:p><w:pPr>${exact24}</w:pPr><w:r><w:t>Y</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>`
  + `<w:p><w:pPr>${exact24}</w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF Anchor </w:instrText></w:r>`
  + '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>cached</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';

/** Each story builder places the zero-row table `t`; `''` is the control. */
const STORIES: ReadonlyArray<readonly [string, (t: string) => Stories]> = [
  ['body, between paragraphs', (t) => ({ body: P('A') + t + P('B') })],
  ['body, between spaced paragraphs', (t) => ({ body: SPACED('A') + t + SPACED('B') })],
  ['body, before a page-break-before paragraph', (t) => ({
    body: t + `<w:p><w:pPr><w:pageBreakBefore/>${exact24}</w:pPr><w:r><w:t>B</w:t></w:r></w:p>`,
  })],
  ['body, last block', (t) => ({ body: P('A') + t })],
  ['body, only block', (t) => ({ body: t })],
  ['body, two in a row', (t) => ({ body: P('A') + t + t + P('B') })],
  ['body, between tables', (t) => ({ body: P('A') + ONE_ROW + t + ONE_ROW + P('B') })],
  ['body, between floating tables', (t) => ({ body: P('A') + ONE_ROW_FLOATING + t + ONE_ROW_FLOATING + P('B') })],
  ['body, vertical (tbRl) section', (t) => ({ body: P('A') + t + P('B'), textDirection: 'tbRl' })],
  ['body, as a whole Cover Pages block', (t) => ({ body: P('A') + cover(t) + P('AFTER') })],
  ['body, ending a Cover Pages block', (t) => ({ body: cover(P('COVER') + t) + P('AFTER') })],
  ['body, inside a REF bookmark', (t) => ({ body: refBookmark(t) })],
  ['nested in a body cell', (t) => ({ body: P('A') + nestedIn(t) + P('B') })],
  ['nested in a body cell, before the end-of-cell mark', (t) => ({
    body: P('A') + table([P('C') + t + '<w:p/>'], [320], 0, false) + P('B'),
  })],
  ['nested in a body cell spanning pages', (t) => ({
    body: P('A') + table([P('C') + t + Array.from({ length: 30 }, (_, i) => P(`L${i}`)).join('')], [320], 0, false) + P('B'),
  })],
  ['header', (t) => ({ body: P('A'), header: P('H') + t + P('H2') })],
  ['footer', (t) => ({ body: P('A'), footer: P('F') + t + P('F2') })],
  ['nested in a header cell', (t) => ({ body: P('A'), header: nestedIn(t) })],
  ['footnote', (t) => ({ body: noteReference('footnote') + P('A'), footnote: P('N') + t + P('N2') })],
  ['endnote', (t) => ({ body: noteReference('endnote') + P('A'), endnote: P('E') + t + P('E2') })],
  ['nested in a footnote cell', (t) => ({ body: noteReference('footnote') + P('A'), footnote: nestedIn(t) })],
  ['text box', (t) => ({ body: textBox(0, 'horz', P('T') + t + P('T2')) + P('A') })],
  ['nested in a text box cell', (t) => ({ body: textBox(0, 'horz', nestedIn(t)) + P('A') })],
];

describe.each(ZERO_ROW_TABLES)('a zero-row %s table', (_kind, zeroRow) => {
  it.each(STORIES)('%s lays out exactly as without it', (_story, build) => {
    expect(observed(layout(build(zeroRow)))).toEqual(observed(layout(build(''))));
  });
});

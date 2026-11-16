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

// A header, footer, footnote or endnote story is laid out in a fixed page
// region, so a page or column break (§17.18.4), a paragraph sectPr (§17.6.18)
// or a Cover Pages building block (§17.5.2) in it can start no page or column.
// Through the real parser, such a story must lay out, on one page, exactly as
// the same story with the construct reduced to an ignored lastRenderedPageBreak
// hint (§17.3.3.13) or removed: the library policy (StoryFlow in the parser),
// matching how a page or column break in a table cell or text box is treated.
// Synthetic linear metrics are a test measurement service only.

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const TW = 20;

const encoder = new TextEncoder();
const xml = (body: string) => encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${body}`);

const exact24 = `<w:spacing w:before="0" w:after="0" w:line="${24 * TW}" w:lineRule="exact"/>`;
const para = (text: string) => `<w:p><w:pPr>${exact24}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;

type Stories = Readonly<{
  body: string;
  header?: string;
  footer?: string;
  footnote?: string;
  endnote?: string;
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
  const document = `<w:document xmlns:w="${W}" xmlns:r="${R}">`
    + `<w:body>${stories.body}<w:sectPr>${references}<w:pgSz w:w="${400 * TW}" w:h="${400 * TW}"/>`
    + `<w:pgMar w:top="${40 * TW}" w:right="${40 * TW}" w:bottom="${40 * TW}" w:left="${40 * TW}" `
    + `w:header="${20 * TW}" w:footer="${20 * TW}" w:gutter="0"/>`
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

const p = (runs: string, props = '') => `<w:p><w:pPr>${exact24}${props}</w:pPr><w:r>${runs}</w:r></w:p>`;
const HINT = '<w:lastRenderedPageBreak/>';
const cover = (content: string) => '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Cover Pages"/>'
  + `</w:docPartObj></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;

/** [construct, its block, the control block without a page-flow effect]. */
const CONSTRUCTS: ReadonlyArray<readonly [string, string, string]> = [
  ...(['page', 'column'] as const).flatMap((type) => {
    const br = `<w:br w:type="${type}"/>`;
    return [
      [`${type} break inside a paragraph`, p(`<w:t>X1</w:t>${br}<w:t>X2</w:t>`), p(`<w:t>X1</w:t>${HINT}<w:t>X2</w:t>`)],
      [`${type} break as the only run`, p(br), p(HINT)],
    ] as const;
  }),
  ['page break opening a paragraph', p('<w:br w:type="page"/><w:t>X2</w:t>'), p(`${HINT}<w:t>X2</w:t>`)],
  ['page break ending a paragraph', p('<w:t>X1</w:t><w:br w:type="page"/>'), p(`<w:t>X1</w:t>${HINT}`)],
  ['paragraph sectPr', p('<w:t>X1</w:t>', '<w:sectPr><w:type w:val="continuous"/></w:sectPr>'), p('<w:t>X1</w:t>')],
  ['Cover Pages block', cover(p('<w:t>X1</w:t>')), p('<w:t>X1</w:t>')],
];

const STORIES: ReadonlyArray<readonly [string, (block: string) => Stories]> = [
  ['header', (block) => ({ body: para('A'), header: para('H') + block + para('H2') })],
  ['footer', (block) => ({ body: para('A'), footer: para('F') + block + para('F2') })],
  ['footnote', (block) => ({ body: noteReference('footnote') + para('A'), footnote: para('N') + block + para('N2') })],
  ['endnote', (block) => ({ body: noteReference('endnote') + para('A'), endnote: para('E') + block + para('E2') })],
];

describe.each(STORIES)('a %s story', (_story, build) => {
  it.each(CONSTRUCTS)('lays out a %s as if it started nothing', (_construct, block, control) => {
    const result = observed(layout(build(block)));
    expect(result.pageCount).toBe(1);
    expect(result).toEqual(observed(layout(build(control))));
  });
});

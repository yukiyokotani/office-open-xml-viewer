/// <reference types="node" />

import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it, vi } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { attachDocumentLayoutVariants } from './layout/document-layout-variants.js';
import {
  attachDocumentLayoutRuntime,
  documentLayoutRuntimeOf,
  layoutSourceStoreOf,
} from './layout/runtime-state.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import { retainRenderWorkerDocumentLayout } from './render-worker-layout.js';
import { renderWorkerOutline } from './render-worker-metadata.js';
import { storeZip } from './conformance/generate.js';
import { DocxDocument } from './document.js';
import { buildDocxOutline } from './outline.js';
import type { DocxDocumentModel } from './types.js';
import type { RenderWorkerRequest } from './worker-protocol.js';

// The outline is projected from real parser output: levels come from
// ECMA-376 §17.3.1.20 `w:outlineLvl`, direct or through the §17.7.4.3
// `w:basedOn` paragraph-style chain, never from style names.

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const TW = 20;

const encoder = new TextEncoder();
const xml = (body: string) => encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${body}`);
const exact24 = `<w:spacing w:before="0" w:after="0" w:line="${24 * TW}" w:lineRule="exact"/>`;

function para(content: string, pPr = '', paraId?: string): string {
  const id = paraId ? ` w14:paraId="${paraId}"` : '';
  return `<w:p${id}><w:pPr>${pPr}${exact24}</w:pPr>${content}</w:p>`;
}
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const style = (id: string) => `<w:pStyle w:val="${id}"/>`;

// Heading1 and Heading2 carry their levels as Word writes them; Chapter has no
// level of its own and inherits Heading2's through basedOn. NotAHeading is
// named like a heading but has no outline level anywhere in its chain.
const STYLES = `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>`
  + '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults>'
  + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
  + '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>'
  + '<w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Heading1"/>'
  + '<w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="Chapter"><w:name w:val="Chapter"/><w:basedOn w:val="Heading2"/></w:style>'
  + '<w:style w:type="paragraph" w:styleId="NotAHeading"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/></w:style>'
  + '</w:styles>';

const BODY = [
  para(`<w:del w:id="1" w:author="A"><w:r><w:delText>Old </w:delText></w:r></w:del>${run('Intro')}`,
    style('Heading1'), '0000000A'),
  para(run('Body text')),
  para(run('Named but not outlined'), style('NotAHeading')),
  para(run('Direct level'), '<w:outlineLvl w:val="2"/>'),
  para(`${run('Sub')}<w:r><w:tab/></w:r>${run('section')}`, style('Heading2')),
  para(run('Second page'), `${style('Chapter')}<w:pageBreakBefore/>`, '0000000B'),
  `<w:tbl><w:tblPr><w:tblW w:w="${200 * TW}" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>`
    + `<w:tblGrid><w:gridCol w:w="${200 * TW}"/></w:tblGrid><w:tr><w:tc>`
    + `<w:tcPr><w:tcW w:w="${200 * TW}" w:type="dxa"/></w:tcPr>${para(run('In table'), style('Heading1'))}`
    + '</w:tc></w:tr></w:tbl>',
  para(run('After')),
].join('');

function outlineDocx(): Uint8Array {
  const document = `<w:document xmlns:w="${W}" xmlns:w14="${W14}" xmlns:r="${R}"><w:body>${BODY}`
    + '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/>'
    + `<w:pgSz w:w="${400 * TW}" w:h="${400 * TW}"/>`
    + `<w:pgMar w:top="${40 * TW}" w:right="${40 * TW}" w:bottom="${40 * TW}" w:left="${40 * TW}" `
    + `w:header="${20 * TW}" w:footer="${20 * TW}" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  return storeZip(new Map([
    ['[Content_Types].xml', xml('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + `<Override PartName="/word/document.xml" ContentType="${CONTENT_TYPE}.document.main+xml"/>`
      + `<Override PartName="/word/styles.xml" ContentType="${CONTENT_TYPE}.styles+xml"/>`
      + `<Override PartName="/word/header1.xml" ContentType="${CONTENT_TYPE}.header+xml"/></Types>`)],
    ['_rels/.rels', xml(`<Relationships xmlns="${PACKAGE_RELS}">`
      + `<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/_rels/document.xml.rels', xml(`<Relationships xmlns="${PACKAGE_RELS}">`
      + `<Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/>`
      + `<Relationship Id="rIdHeader" Type="${R}/header" Target="header1.xml"/></Relationships>`)],
    ['word/document.xml', xml(document)],
    ['word/styles.xml', xml(STYLES)],
    // A heading-styled header paragraph is outside the body outline.
    ['word/header1.xml', xml(`<w:hdr xmlns:w="${W}">${para(run('Running head'), style('Heading1'))}</w:hdr>`)],
  ]));
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

function parse(): DocxDocumentModel {
  const archive = new DocxArchive(outlineDocx());
  try {
    return normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally {
    archive.free();
  }
}

/** A main-mode engine over the real model, without a Worker or Canvas. */
function mainDocument(model: DocxDocumentModel): DocxDocument {
  const services = createLayoutServices(model, { measureContext: measureContext() });
  attachDocumentLayoutVariants({
    source: layoutSourceStore(model),
    services,
    defaultCurrentDateMs: 0,
    buildLayout: (options) => layoutDocument(model, services, options),
  });
  const document = Object.create(DocxDocument.prototype) as DocxDocument;
  attachDocumentLayoutRuntime(document, 0);
  documentLayoutRuntimeOf(document).services = services;
  Object.assign(document, { _mode: 'main', _document: model });
  return document;
}

const body = (...path: number[]) => ({ story: 'body', storyInstance: 'body', path });

const EXPECTED = [
  { level: 1, text: 'Intro', paragraphId: '0000000A', source: body(0), pageIndex: 0 },
  { level: 3, text: 'Direct level', source: body(3), pageIndex: 0 },
  { level: 2, text: 'Sub section', source: body(4), pageIndex: 0 },
  { level: 2, text: 'Second page', paragraphId: '0000000B', source: body(5), pageIndex: 1 },
  { level: 1, text: 'In table', source: body(6, 0, 0, 0), pageIndex: 1 },
];

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

describe('DocxDocument.getOutline', () => {
  it('lists body headings with style-chain levels, text, ids and laid-out pages (main mode)', async () => {
    await expect(mainDocument(parse()).getOutline()).resolves.toEqual(EXPECTED);
  });

  it('projects the same entries from the worker-retained layout, safe to post', () => {
    const model = parse();
    const worker = retainRenderWorkerDocumentLayout(
      layoutSourceStore(model),
      createLayoutServices(model, { measureContext: measureContext() }),
      0,
    );
    const entries = renderWorkerOutline(worker, {});
    expect(entries).toEqual(EXPECTED);
    expect(structuredClone(entries)).toEqual(entries);
  });

  it('omits the page of a heading past the laid-out prefix', async () => {
    const document = mainDocument(parse());
    const services = documentLayoutRuntimeOf(document).services!;
    const layout = (document as unknown as { _getLayout(): Parameters<typeof buildDocxOutline>[1] })._getLayout()!;
    const prefix = { ...layout, pages: layout.pages.slice(0, 1) };
    const entries = buildDocxOutline(layoutSourceStoreOf(services)!, prefix);
    expect(entries.map((entry) => entry.pageIndex)).toEqual([0, 0, 0, undefined, undefined]);
  });

  it('asks the worker for the outline of the active layout view (worker mode)', async () => {
    const requests: RenderWorkerRequest[] = [];
    const document = Object.create(DocxDocument.prototype) as DocxDocument;
    Object.assign(document, {
      _mode: 'worker',
      _bridge: {
        request: vi.fn(async (factory: (id: number) => RenderWorkerRequest) => {
          const request = factory(4);
          requests.push(request);
          return { type: 'outlineBuilt', id: 4, entries: EXPECTED };
        }),
      },
    });
    attachDocumentLayoutRuntime(document, 0);
    documentLayoutRuntimeOf(document).activeLayoutOptions = Object.freeze({
      currentDateMs: 20,
      showTrackedChanges: true,
    });

    await expect(document.getOutline()).resolves.toEqual(EXPECTED);
    expect(requests).toEqual([
      { type: 'outline', id: 4, opts: { currentDate: 20, showTrackedChanges: true } },
    ]);
  });
});

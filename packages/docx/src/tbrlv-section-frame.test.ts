/// <reference types="node" />

import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import type { DocxDocumentModel } from './types.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { storeZip } from './conformance/generate.js';

// `word-section-tbrlv-tblrv-page-frame`: Word lays a section-level tbRlV out
// exactly as tbLrV. The layout of a tbRlV section must therefore be the
// layout of the same section authored as tbLrV, whatever else it carries.

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const encoder = new TextEncoder();
const xml = (body: string) => encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${body}`);

function docx(direction: string, sectionExtras: string, settings: string): Uint8Array {
  const run = (text: string, rPr = '') => `<w:r><w:rPr><w:sz w:val="21"/>${rPr}</w:rPr><w:t>${text}</w:t></w:r>`;
  const paragraphs = Array.from({ length: 60 }, (_, index) =>
    `<w:p>${run(`Ab${index}「あい」`)}${run('うえお', '<w:u w:val="single"/>')}</w:p>`).join('');
  const document = `<w:document xmlns:w="${W}"><w:body>${paragraphs}`
    + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    + '<w:pgMar w:top="1200" w:right="1800" w:bottom="1600" w:left="800" w:header="720" w:footer="720" w:gutter="720"/>'
    + `${sectionExtras}<w:textDirection w:val="${direction}"/><w:docGrid w:type="lines" w:linePitch="360"/>`
    + '</w:sectPr></w:body></w:document>';
  return storeZip(new Map([
    ['[Content_Types].xml', xml('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      + '</Types>')],
    ['_rels/.rels', xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/_rels/document.xml.rels', xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="rId1" Type="${REL}/settings" Target="settings.xml"/></Relationships>`)],
    ['word/settings.xml', xml(`<w:settings xmlns:w="${W}">${settings}</w:settings>`)],
    ['word/document.xml', xml(document)],
  ]));
}

function measureContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'normal',
    measureText: (text: string) => {
      const px = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...text].length * px,
        fontBoundingBoxAscent: px * 0.8,
        fontBoundingBoxDescent: px * 0.2,
        actualBoundingBoxAscent: px * 0.8,
        actualBoundingBoxDescent: px * 0.2,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

/** The laid-out pages with the authored token itself masked. */
function pagesOf(direction: string, sectionExtras: string, settings: string): string {
  const archive = new DocxArchive(docx(direction, sectionExtras, settings));
  let model: DocxDocumentModel;
  try {
    model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally {
    archive.free();
  }
  const layout = layoutDocument(
    model,
    createLayoutServices(model, { measureContext: measureContext() }),
    { currentDateMs: 0 },
  );
  expect(layout.pages.length).toBeGreaterThan(1);
  return JSON.stringify(layout.pages).replaceAll(`"${direction}"`, '"direction"');
}

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

describe('a tbRlV section is laid out as tbLrV', () => {
  it.each([
    ['a gutter', '', ''],
    ['mirrored margins and a gutter', '', '<w:mirrorMargins/>'],
    ['two columns', '<w:cols w:num="2" w:space="720"/>', ''],
  ])('with %s', (_name, sectionExtras, settings) => {
    expect(pagesOf('tbRlV', sectionExtras, settings)).toBe(pagesOf('tbLrV', sectionExtras, settings));
  });
});

/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { measureContext } from './test-support/tab-fitting.test-support.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const encode = (text: string) => new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${text}`);

function docxBytes(paragraphXml: string, pageWidthTwips: number): Uint8Array {
  const document = `<w:document xmlns:w="${W}"><w:body><w:p>${paragraphXml}</w:p><w:sectPr>`
    + `<w:pgSz w:w="${pageWidthTwips}" w:h="15840"/>`
    + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
    + '</w:sectPr></w:body></w:document>';
  return storeZip(new Map([
    ['[Content_Types].xml', encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '</Types>')],
    ['_rels/.rels', encode(`<Relationships xmlns="${PKG}"><Relationship Id="rId1" `
      + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/document.xml', encode(document)],
  ]));
}

function lineText(paragraphXml: string, pageWidthTwips = 12240): string {
  const archive = new DocxArchive(docxBytes(paragraphXml, pageWidthTwips));
  let model;
  try {
    model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally { archive.free(); }
  const result = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext(6) }), { currentDateMs: 0 });
  const paragraph = result.pages[0]?.layers.body.find(node => node.kind === 'paragraph');
  if (paragraph?.kind !== 'paragraph') throw new Error('Missing paragraph');
  return paragraph.lines.map(line => line.placements
    .map(node => node.kind !== 'text' ? '' : node.optionalHyphenGlyph ? '~' : node.text).join('')).join('|');
}

const r = (inner: string) => `<w:r>${inner}</w:r>`;
const t = (text: string) => `<w:t xml:space="preserve">${text}</w:t>`;
const nbh = '<w:noBreakHyphen/>';
const shy = '<w:softHyphen/>';

// Rows mirror the variant table in #2193. The first four threw "Segment
// measurement does not match its retained text range context" from 0.89.0:
// the unselected soft-hyphen marker kept its '-' glyph's shape request while
// its own retained text was empty. Each row is laid out on a wide page (one
// line, marker unselected) and on a 36 pt measure that forces the joined unit
// through the wrapping paths. A selected conditional glyph is reported as '~'.
// The narrow partitions of the last four rows are those 0.89.9 already drew.
describe('#2193 soft hyphen after a no-break-hyphen word', () => {
  it.each([
    ['issue repro', r(t('reimburse out')) + r(nbh + t('of')) + r(shy) + r(nbh + t('pocket expenses.')),
      'reimburse out-of-pocket expenses.', 'reimbu|rse |out-of-p|ocket |expens|es.'],
    ['nbh + shy + nbh', r(t('aa bb')) + r(nbh + t('cc')) + r(shy) + r(nbh + t('dd')), 'aa bb-cc-dd', 'aa |bb-cc-d|d'],
    ['nbh + shy + plain', r(t('aa bb')) + r(nbh + t('cc')) + r(shy) + r(t('dd')), 'aa bb-ccdd', 'aa |bb-cc~|dd'],
    ['single run', r(t('aa bb') + nbh + t('cc') + shy + nbh + t('dd')), 'aa bb-cc-dd', 'aa |bb-cc-d|d'],
    ['no earlier space', r(t('aabb')) + r(nbh + t('cc')) + r(shy) + r(nbh + t('dd')), 'aabb-cc-dd', 'aabb-c|c-dd'],
    ['no noBreakHyphen', r(t('aa bb')) + r(t('cc')) + r(shy) + r(nbh + t('dd')), 'aa bbcc-dd', 'aa |bbcc-d|d'],
    ['literal U+00AD', r(t('aa bb')) + r(nbh + t('cc')) + r(t('\u00ad')) + r(nbh + t('dd')),
      'aa bb-cc\u00ad-dd', 'aa |bb-cc|\u00ad-dd'],
    ['soft hyphen removed', r(t('aa bb')) + r(nbh + t('cc')) + r(nbh + t('dd')), 'aa bb-cc-dd', 'aa |bb-c|c-dd'],
  ])('%s', (_name, xml, wide, narrow) => {
    expect(lineText(xml)).toBe(wide);
    expect(lineText(xml, 2880 + 36 * 20)).toBe(narrow);
  });
});

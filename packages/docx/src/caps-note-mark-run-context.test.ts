/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const encode = (text: string) => new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${text}`);

function docxBytes(paragraphXml: string): Uint8Array {
  const document = `<w:document xmlns:w="${W}"><w:body><w:p>${paragraphXml}</w:p><w:sectPr>`
    + '<w:pgSz w:w="12240" w:h="15840"/>'
    + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
    + '</w:sectPr></w:body></w:document>';
  const notes = (kind: 'footnote' | 'endnote') => `<w:${kind}s xmlns:w="${W}">`
    + `<w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}>`
    + `<w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>`
    + `<w:${kind} w:id="1"><w:p><w:r><w:t>note</w:t></w:r></w:p></w:${kind}></w:${kind}s>`;
  return storeZip(new Map([
    ['[Content_Types].xml', encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
      + '<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>'
      + '</Types>')],
    ['_rels/.rels', encode(`<Relationships xmlns="${PKG}"><Relationship Id="rId1" `
      + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/_rels/document.xml.rels', encode(`<Relationships xmlns="${PKG}">`
      + `<Relationship Id="rId1" Type="${REL}/footnotes" Target="footnotes.xml"/>`
      + `<Relationship Id="rId2" Type="${REL}/endnotes" Target="endnotes.xml"/></Relationships>`)],
    ['word/document.xml', encode(document)],
    ['word/footnotes.xml', encode(notes('footnote'))],
    ['word/endnotes.xml', encode(notes('endnote'))],
  ]));
}

const measureHost = {
  font: '', letterSpacing: '0px', fontKerning: 'auto',
  measureText(text: string) {
    return {
      width: [...text].length * 6,
      actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
      fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
    };
  },
} as unknown as CanvasRenderingContext2D;

function bodyText(paragraphXml: string): string {
  const archive = new DocxArchive(docxBytes(paragraphXml));
  let model;
  try {
    model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally { archive.free(); }
  const result = layoutDocument(model, createLayoutServices(model, { measureContext: measureHost }), { currentDateMs: 0 });
  const paragraph = result.pages[0]?.layers.body.find(node => node.kind === 'paragraph');
  if (paragraph?.kind !== 'paragraph') throw new Error('Missing paragraph');
  return paragraph.lines.map(line => line.placements
    .map(node => node.kind !== 'text' ? '' : node.optionalHyphenGlyph ? '~' : node.text).join('')).join('|');
}

const fonts = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>';
const r = (rPr: string, inner: string) => `<w:r><w:rPr>${fonts}${rPr}</w:rPr>${inner}</w:r>`;
const t = (text: string) => `<w:t xml:space="preserve">${text}</w:t>`;

// A custom note mark (CT_FtnEdnRef/@customMarkFollows) keeps a zero-width
// metric host laid out from a Latin probe. The probe is emitted through the
// run's §17.3.2.5 caps / §17.3.2.33 small-caps transform, so its run context
// must be the same transformed text. The authored probe ('x' against a shaped
// 'X') once threw "Text shape request does not match its full run context" and
// the whole document failed to render.
describe('caps note mark run context', () => {
  it.each([
    ['all-caps custom footnote mark', r('<w:caps/>', t('ab'))
      + r('<w:caps/>', '<w:footnoteReference w:customMarkFollows="1" w:id="1"/>') + r('<w:caps/>', t('*')), 'AB*'],
    ['small-caps custom endnote mark', r('<w:smallCaps/>', t('ab'))
      + r('<w:smallCaps/>', '<w:endnoteReference w:customMarkFollows="1" w:id="1"/>'), 'AB'],
  ])('%s', (_name, xml, expected) => {
    expect(bodyText(xml)).toBe(expected);
  });
});

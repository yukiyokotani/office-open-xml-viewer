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
  const footnotes = `<w:footnotes xmlns:w="${W}">`
    + '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
    + '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
    + '<w:footnote w:id="1"><w:p><w:r><w:t>note</w:t></w:r></w:p></w:footnote></w:footnotes>';
  return storeZip(new Map([
    ['[Content_Types].xml', encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
      + '</Types>')],
    ['_rels/.rels', encode(`<Relationships xmlns="${PKG}"><Relationship Id="rId1" `
      + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/_rels/document.xml.rels', encode(`<Relationships xmlns="${PKG}"><Relationship Id="rId1" `
      + `Type="${REL}/footnotes" Target="footnotes.xml"/></Relationships>`)],
    ['word/document.xml', encode(document)],
    ['word/footnotes.xml', encode(footnotes)],
  ]));
}

// A measuring context whose Symbol face is absent: every symbol code converts
// to Unicode (core symbol-face.ts judges each code against the painted face).
const absentSymbolHost = {
  font: '', letterSpacing: '0px', fontKerning: 'auto',
  measureText(text: string) {
    const cp = text.codePointAt(0) ?? 0;
    return {
      width: text.length === 1 ? 5 + (cp % 5) : [...text].length * 8,
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
  const result = layoutDocument(model, createLayoutServices(model, { measureContext: absentSymbolHost }), { currentDateMs: 0 });
  const paragraph = result.pages[0]?.layers.body.find(node => node.kind === 'paragraph');
  if (paragraph?.kind !== 'paragraph') throw new Error('Missing paragraph');
  return paragraph.lines.map(line => line.placements
    .map(node => node.kind !== 'text' ? '' : node.optionalHyphenGlyph ? '~' : node.text).join('')).join('|');
}

const fonts = (ascii: string, other = ascii) =>
  `<w:rFonts w:ascii="${ascii}" w:hAnsi="${ascii}" w:eastAsia="${other}" w:cs="${other}"/>`;
const symbol = fonts('Symbol', 'Times New Roman');
const r = (rPr: string, inner: string) => `<w:r><w:rPr>${rPr}</w:rPr>${inner}</w:r>`;
const t = (text: string) => `<w:t xml:space="preserve">${text}</w:t>`;

// Every piece of a run is emitted through the run's case (§17.3.2.5) and
// §17.3.3.30 symbol transforms, and its shaping request is checked against the
// run's display text. Pieces that are their own run (metric probes,
// optional-hyphen glyphs) once kept the authored probe as that context: an
// all-caps mark host ('x' → 'X'), and since #2192 bare Symbol codes ('x' → ξ,
// '-' → −) converted in a face that lacks them, threw
// "Text shape request does not match its full run context" and the whole
// document failed to render.
describe('symbol-font run context', () => {
  it.each([
    ['custom footnote mark host', r(symbol, t('ab'))
      + r(symbol, '<w:footnoteReference w:customMarkFollows="1" w:id="1"/>') + r(symbol, t('*')), 'αβ∗'],
    ['all-caps custom footnote mark host', r(`${fonts('Times New Roman')}<w:caps/>`, t('ab'))
      + r(`${fonts('Times New Roman')}<w:caps/>`, '<w:footnoteReference w:customMarkFollows="1" w:id="1"/>'), 'AB'],
    ['leading soft hyphen run', r(symbol, '<w:softHyphen/>') + r(symbol, t('cd')), 'χδ'],
    ['soft hyphen inside a run', r(symbol, t('ab') + '<w:softHyphen/>' + t('cd')), 'αβχδ'],
    ['tabs, PUA and bare codes', r(symbol, t('a\uF062') + '<w:tab/>' + t('c')), 'αβχ'],
    ['complex-script slot', r(`${fonts('Times New Roman', 'Symbol')}<w:rtl/>`, t('ab')), 'αβ'],
    ['east Asian slot', r(fonts('Times New Roman', 'Symbol'), t('a\u65e5b')), 'a\u65e5b'],
    ['small caps', r(`${symbol}<w:smallCaps/>`, t('aB')), 'ΑΒ'],
  ])('%s', (_name, xml, expected) => {
    expect(bodyText(xml)).toBe(expected);
  });
});

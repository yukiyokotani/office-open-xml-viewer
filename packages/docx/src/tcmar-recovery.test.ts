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
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const encode = (text: string) => new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${text}`);

function docxBytes(bodyXml: string, sectionXml: string, footerXml?: string): Uint8Array {
  const document = `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${bodyXml}<w:sectPr>`
    + (footerXml ? '<w:footerReference w:type="default" r:id="rIdFooter"/>' : '')
    + `${sectionXml}</w:sectPr></w:body></w:document>`;
  const parts = new Map([
    ['[Content_Types].xml', encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + (footerXml ? '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' : '')
      + '</Types>')],
    ['_rels/.rels', encode(`<Relationships xmlns="${PKG}"><Relationship Id="rId1" `
      + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/_rels/document.xml.rels', encode(`<Relationships xmlns="${PKG}">`
      + (footerXml ? '<Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>' : '')
      + '</Relationships>')],
    ['word/document.xml', encode(document)],
  ]);
  if (footerXml) parts.set('word/footer1.xml', encode(footerXml));
  return storeZip(parts);
}

function layout(bytes: Uint8Array, advancePt: number) {
  const archive = new DocxArchive(bytes);
  let model;
  try {
    model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally { archive.free(); }
  return layoutDocument(model, createLayoutServices(model, { measureContext: measureContext(advancePt) }), { currentDateMs: 0 });
}


// #2125: CT_TcPrBase allows one w:tcMar. Word recovers repeated ones edge by
// edge in document order (a later edge element wins; an empty or partial one
// keeps the others), and takes start over left for the same edge. Measured on
// Windows Word controls w2125-tcmar-{plain,tblmar288}-mode{14,15}.
const cellTextX = (tcMar: string, tblPr = '') => {
  const table = `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/>${tblPr}</w:tblPr>`
    + '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/>'
    + `${tcMar}</w:tcPr><w:p><w:r><w:t>Xa</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>`;
  // (tblPr may add table-level margins or an indent.)
  const result = layout(docxBytes(table, '<w:pgSz w:w="11906" w:h="16838"/>'
    + '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/>'), 5);
  const find = (nodes: readonly unknown[]): number | undefined => {
    for (const node of nodes as Array<Record<string, unknown>>) {
      if (node.kind === 'table-cell') return (node.contentBounds as { xPt: number }).xPt;
      for (const value of Object.values(node)) {
        if (Array.isArray(value)) { const x = find(value); if (x !== undefined) return x; }
      }
    }
    return undefined;
  };
  return find(result.pages[0]!.layers.body)!;
};
const mar = (edges: Record<string, number>) => `<w:tcMar>${Object.entries(edges).map(([k, v]) => `<w:${k} w:w="${v}" w:type="dxa"/>`).join('')}</w:tcMar>`;

describe('repeated w:tcMar recovery (#2125)', () => {
  it('merges repeated cell margins edge by edge as Word does', () => {
    const single = cellTextX(mar({ left: 720 }));
    const none = cellTextX(mar({ left: 0 }));
    expect(single - none).toBeCloseTo(36, 6);
    expect(cellTextX(`<w:tcMar/>${mar({ left: 720 })}`)).toBeCloseTo(single, 6);
    expect(cellTextX(`${mar({ left: 720 })}<w:tcMar/>`)).toBeCloseTo(single, 6);
    expect(cellTextX(mar({ left: 720 }) + mar({ left: 0 }))).toBeCloseTo(none, 6);
    expect(cellTextX(mar({ left: 0 }) + mar({ left: 720 }))).toBeCloseTo(single, 6);
    expect(cellTextX(mar({ right: 720 }) + mar({ left: 720 }))).toBeCloseTo(single, 6);
    expect(cellTextX(mar({ top: 360 }) + mar({ left: 720 }))).toBeCloseTo(single, 6);
    expect(cellTextX(mar({ left: 240 }) + mar({ left: 480 }) + mar({ left: 720 }))).toBeCloseTo(single, 6);
    expect(cellTextX(mar({ start: 720 }) + mar({ left: 0 }))).toBeCloseTo(single, 6);
    expect(cellTextX(mar({ left: 0 }) + mar({ start: 720 }))).toBeCloseTo(single, 6);
  });

  it('takes the margin provenance from the winning logical edge', () => {
    // An ignored physical left beside a winning start must not switch on the
    // physical-left table placement (a mode-14 table at tblInd 0).
    const tbl = '<w:tblInd w:w="0" w:type="dxa"/>';
    expect(cellTextX(mar({ start: 720, left: 0 }), tbl)).toBeCloseTo(cellTextX(mar({ start: 720 }), tbl), 6);
  });
});

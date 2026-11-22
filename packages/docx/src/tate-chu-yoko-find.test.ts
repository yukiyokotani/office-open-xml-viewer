/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { measureContext } from './test-support/tab-fitting.test-support.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import { textRunsForPage } from './text-run-projection.js';

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

const encode = (text: string) => new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${text}`);

describe('tate-chu-yoko find geometry', () => {
  it('gives both digits of a compressed tbRl cell a box, side by side across the run', () => {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const body = '<w:p><w:r><w:t>あ</w:t></w:r><w:r><w:rPr><w:sz w:val="16"/>'
      + '<w:eastAsianLayout w:id="1" w:vert="1" w:vertCompress="1"/></w:rPr><w:t>２９</w:t></w:r></w:p>';
    const bytes = storeZip(new Map([
      ['[Content_Types].xml', encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        + '<Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" '
        + 'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
      ['_rels/.rels', encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        + 'Target="word/document.xml"/></Relationships>')],
      ['word/document.xml', encode(`<w:document ${W}><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>`
        + '<w:textDirection w:val="tbRl"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        + 'w:header="0" w:footer="0" w:gutter="0"/></w:sectPr></w:body></w:document>')],
    ]));
    const archive = new DocxArchive(bytes);
    let parsed;
    try {
      parsed = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse())));
    } finally { archive.free(); }
    const model = parsed.document;
    const layout = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext(6) }), { currentDateMs: 0 });
    const run = textRunsForPage(layout, 0, { scale: 1 }).find((item) => item.text === '２９')!;
    const [two, nine] = run.clusterBoxes!;
    // Both span the cell along the run; across it they sit side by side.
    expect(two!.width).toBeGreaterThan(0);
    expect(nine!.width).toBe(two!.width);
    expect(two!.height).toBeGreaterThan(0);
    expect(nine!.height).toBe(two!.height);
    expect(two!.y).toBeCloseTo(nine!.y! + nine!.height!, 6);
  });
});

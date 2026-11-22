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

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
  + 'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const encode = (text: string) => new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${text}`);

/** The top of the text after a vanished mark holding `content`, and the
 * drawings painted. */
function observe(content: string) {
  const body = `<w:p><w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr>${content}</w:p><w:p><w:r><w:t>After</w:t></w:r></w:p>`;
  const bytes = storeZip(new Map([
    ['[Content_Types].xml', encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', encode(`<Relationships xmlns="${PKG}"><Relationship Id="rId1" `
      + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/document.xml', encode(`<w:document ${NS}><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>`
      + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr></w:body></w:document>')],
  ]));
  const archive = new DocxArchive(bytes);
  let parsed;
  try {
    parsed = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse())));
  } finally { archive.free(); }
  const model = parsed.document;
  const layout = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext(5) }), { currentDateMs: 0 });
  return {
    top: textRunsForPage(layout, 0, { scale: 1 }).find((run) => run.text === 'After')!.y,
    drawings: layout.pages.flatMap((page) => page.layers.paintOrder.filter((entry) => entry.kind === 'drawing')).length,
  };
}

describe('a vanished paragraph mark that anchors only a floating VML shape', () => {
  it('joins the next paragraph as the empty vanished mark does, and still paints the shape', () => {
    const shape = '<w:r><w:pict><v:rect style="position:absolute;margin-left:100pt;margin-top:300pt;width:20pt;height:20pt;'
      + 'mso-position-horizontal-relative:page;mso-position-vertical-relative:page" fillcolor="red" stroked="f"/></w:pict></w:r>';
    const empty = observe('');
    const anchored = observe(shape);
    expect(anchored.top).toBe(empty.top);
    expect(anchored.drawings).toBe(empty.drawings + 1);
  });
});

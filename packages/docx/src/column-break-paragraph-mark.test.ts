import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import type { ParagraphLayout } from './layout/types.js';

// `word-column-break-paragraph-mark`: the mark of a column-break-only
// paragraph opens the next column as an empty line.

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const context = {
  font: '10px Arial', letterSpacing: '0px', fontKerning: 'none',
  measureText: (text: string) => ({
    width: [...text].length * 6,
    fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
    actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
  }),
} as unknown as CanvasRenderingContext2D;

/** The first text of the second column and its offset below the column top
 * (72pt), in 18pt grid lines. */
function secondColumnFirstText(breakParagraph: string): Readonly<{ text: string; gridLine: number }> {
  const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const document = `<w:document xmlns:w="${W}"><w:body>${p('A')}${p('B')}${breakParagraph}${p('C')}${p('D')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:cols w:num="2" w:space="720"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:body></w:document>`;
  const files = new Map([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'],
    ['_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['word/document.xml', document],
  ]);
  const archive = new DocxArchive(storeZip(new Map([...files].map(([name, xml]) => [name, new TextEncoder().encode(xml)]))));
  let model;
  try {
    model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally {
    archive.free();
  }
  const page = layoutDocument(model, createLayoutServices(model, { measureContext: context }), { currentDateMs: 0 }).pages[0]!;
  const lines = page.layers.body
    .filter((node): node is ParagraphLayout => node.kind === 'paragraph')
    .flatMap((node) => node.lines)
    .filter((line) => line.bounds.xPt > 300)
    .sort((a, b) => a.bounds.yPt - b.bounds.yPt);
  const line = lines.find((entry) => entry.placements.some((placement) => placement.kind === 'text'))!;
  const placement = line.placements.find((entry) => entry.kind === 'text')!;
  return { text: placement.kind === 'text' ? placement.text : '', gridLine: (line.bounds.yPt - 72) / 18 };
}

describe('column-break paragraphs (§17.3.3.1)', () => {
  it('opens the next column with the mark of a column-break-only paragraph', () => {
    expect(secondColumnFirstText('<w:p><w:r><w:br w:type="column"/></w:r></w:p>'))
      .toEqual({ text: 'C', gridLine: 1 });
  });

  it('starts the next column with the text after an inline column break', () => {
    expect(secondColumnFirstText('<w:p><w:r><w:t>X</w:t><w:br w:type="column"/><w:t>Y</w:t></w:r></w:p>'))
      .toEqual({ text: 'Y', gridLine: 0 });
  });
});

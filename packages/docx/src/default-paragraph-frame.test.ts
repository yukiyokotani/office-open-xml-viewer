import { readFile } from 'node:fs/promises';
import { beforeAll, expect, it } from 'vitest';
import { storeZip } from './conformance/generate.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { measureContext } from './test-support/tab-fitting.test-support.js';
import { textRunsForPage } from './text-run-projection.js';
import init, { DocxArchive } from './wasm/docx_parser.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const P = 'http://schemas.openxmlformats.org/package/2006/relationships';
const neutralFrame = '<w:framePr w:w="0" w:h="0" w:hRule="exact" w:hSpace="0" w:vSpace="0" w:vAnchor="margin" w:xAlign="left" w:y="0"/>';

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

function loadModel(defaultFrame: string, directFrame = '') {
  const paragraphs = Array.from({ length: 12 }, (_, i) =>
    `<w:p>${i === 0 && directFrame ? `<w:pPr>${directFrame}</w:pPr>` : ''}<w:r><w:t>Paragraph ${i + 1} stays in the body flow.</w:t></w:r></w:p>`).join('');
  const parts = [
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'],
    ['_rels/.rels', `<Relationships xmlns="${P}"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['word/_rels/document.xml.rels', `<Relationships xmlns="${P}"><Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/></Relationships>`],
    ['word/styles.xml', `<w:styles xmlns:w="${W}"><w:docDefaults><w:pPrDefault><w:pPr>${defaultFrame}</w:pPr></w:pPrDefault></w:docDefaults></w:styles>`],
    ['word/document.xml', `<w:document xmlns:w="${W}"><w:body>${paragraphs}<w:sectPr><w:pgSz w:w="6000" w:h="3000"/><w:pgMar w:top="400" w:right="400" w:bottom="400" w:left="400" w:header="200" w:footer="200" w:gutter="0"/></w:sectPr></w:body></w:document>`],
  ] as const;
  const encoder = new TextEncoder();
  const archive = new DocxArchive(storeZip(new Map(parts.map(([name, value]) => [name, encoder.encode(value)]))));
  try {
    return normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally {
    archive.free();
  }
}

it('paginates the neutral zero-width document default like ordinary body text', () => {
  const pageText = (defaultFrame: string) => {
    const model = loadModel(defaultFrame);
    const services = createLayoutServices(model, { measureContext: measureContext() });
    const result = layoutDocument(model, services, { currentDateMs: 0 });
    return result.pages.map((_page, index) => textRunsForPage(result, index, { scale: 1 })
      .map(({ text, x, y }) => ({ text, x, y })));
  };
  const ordinary = pageText('');
  expect(ordinary.length).toBeGreaterThan(1);
  const actual = pageText(neutralFrame);
  expect(actual).toHaveLength(ordinary.length);
  expect(actual).toEqual(ordinary);
});

it('retains defaults that size or position a frame, and lets direct frame properties inherit', () => {
  const firstFrame = (defaultFrame: string, directFrame?: string) => {
    const first = loadModel(defaultFrame, directFrame).body[0];
    if (first?.type !== 'paragraph') throw new Error('Expected a paragraph');
    return first.framePr;
  };
  expect(firstFrame(neutralFrame.replace('w:w="0"', 'w:w="720"'))?.w).toBe(36);
  expect(firstFrame(neutralFrame.replace('w:h="0"', 'w:h="720"'))?.h).toBe(36);
  expect(firstFrame(neutralFrame.replace('w:xAlign="left"', 'w:xAlign="right"'))?.xAlign).toBe('right');
  expect(firstFrame('', neutralFrame)?.w).toBe(0);
  expect(firstFrame(neutralFrame, '<w:framePr w:w="720"/>')).toMatchObject({
    w: 36, vAnchor: 'margin', xAlign: 'left', y: 0,
  });
});

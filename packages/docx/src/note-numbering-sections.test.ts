import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { formatCustomNoteNumber } from './line-breaker/text-runs.js';
import type { DocxDocumentModel } from './types.js';
import type { DocumentLayout, PaintNode } from './layout/types.js';

// ECMA-376 §17.11.4/.5/.11/.12 note numbering through a parsed package:
// section-wide overrides (§17.11.5/.11), numRestart (§17.11.19), custom
// numFmt/@format patterns and Word's unsignedShort numStart
// (WORD_NOTE_NUMBER_FORMAT_DEFAULTS).
beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

const context = {
  font: '10px Arial', letterSpacing: '0px', fontKerning: 'none',
  measureText: (text: string) => ({
    width: [...text].length * 6,
    fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
    actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
  }),
} as unknown as CanvasRenderingContext2D;

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const encode = (xml: string) => new TextEncoder().encode(xml);

interface Control {
  readonly kind: 'footnote' | 'endnote';
  /** Document-wide footnotePr/endnotePr children. */
  readonly documentPr?: string;
  /** Section-wide overrides of the first and second section. */
  readonly sectionPr?: readonly [string, string];
}

/** Two sections, each with two references; notes 1-4. */
function control({ kind, documentPr = '', sectionPr = ['', ''] }: Control): DocxDocumentModel {
  const ref = (id: number) => `<w:r><w:${kind}Reference w:id="${id}"/></w:r>`;
  const pr = (children: string) => children ? `<w:${kind}Pr>${children}</w:${kind}Pr>` : '';
  const sect = (children: string, type = '') => `<w:sectPr>${pr(children)}${type}<w:pgSz w:w="12240" w:h="15840"/>`
    + '<w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr>';
  const body = `<w:p><w:r><w:t>A</w:t></w:r>${ref(1)}${ref(2)}</w:p>`
    + `<w:p><w:pPr>${sect(sectionPr[0], '<w:type w:val="continuous"/>')}</w:pPr></w:p>`
    + `<w:p><w:r><w:t>B</w:t></w:r>${ref(3)}${ref(4)}</w:p>${sect(sectionPr[1], '<w:type w:val="continuous"/>')}`;
  const notes = [1, 2, 3, 4].map((id) => `<w:${kind} w:id="${id}"><w:p><w:r><w:t>N${id}</w:t></w:r></w:p></w:${kind}>`).join('');
  const parts = new Map<string, string>([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + `<Override PartName="/word/${kind}s.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}s+xml"/>`
      + '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'],
    ['_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="d" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="n" Type="${R}/${kind}s" Target="${kind}s.xml"/><Relationship Id="s" Type="${R}/settings" Target="settings.xml"/></Relationships>`],
    ['word/settings.xml', `<w:settings xmlns:w="${W}" xmlns:mc="${MC}" xmlns:w14="${W14}">${pr(documentPr)}</w:settings>`],
    ['word/document.xml', `<w:document xmlns:w="${W}" xmlns:mc="${MC}" xmlns:w14="${W14}"><w:body>${body}</w:body></w:document>`],
    [`word/${kind}s.xml`, `<w:${kind}s xmlns:w="${W}">${notes}</w:${kind}s>`],
  ]);
  const archive = new DocxArchive(storeZip(new Map([...parts].map(([name, xml]) => [name, encode(xml)]))));
  try {
    return normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally {
    archive.free();
  }
}

/** Body reference labels in page order. */
function referenceLabels(layout: DocumentLayout): string[] {
  const labels: string[] = [];
  const visit = (node: PaintNode) => {
    if (node.kind !== 'paragraph') return;
    for (const line of node.lines) {
      for (const placement of line.placements) {
        if (placement.kind === 'text' && placement.noteReference && placement.text.length > 0) labels.push(placement.text);
      }
    }
  };
  for (const page of layout.pages) page.layers.body.forEach(visit);
  return labels;
}

function render(spec: Control): string[] {
  const model = control(spec);
  return referenceLabels(layoutDocument(model, createLayoutServices(model, { measureContext: context }), { currentDateMs: 0 }));
}

describe('section-aware note numbering', () => {
  it('applies a section-wide format and start, continuing into the next section', () => {
    expect(render({ kind: 'endnote', sectionPr: ['<w:numFmt w:val="upperLetter"/><w:numStart w:val="3"/>', ''] }))
      .toEqual(['C', 'D', 'v', 'vi']);
  });

  it('restarts each section with that section numbering', () => {
    expect(render({ kind: 'footnote',
      sectionPr: ['<w:numRestart w:val="eachSect"/>', '<w:numFmt w:val="lowerLetter"/><w:numRestart w:val="eachSect"/>'] }))
      .toEqual(['1', '2', 'a', 'b']);
  });

  it('ignores document-wide numbering in settings.xml (word-section-note-numbering)', () => {
    expect(render({ kind: 'footnote', documentPr: '<w:numFmt w:val="upperLetter"/><w:numStart w:val="3"/>' }))
      .toEqual(['1', '2', '3', '4']);
    expect(render({ kind: 'endnote', documentPr: '<w:numFmt w:val="decimal"/>' })).toEqual(['i', 'ii', 'iii', 'iv']);
  });

  it('formats a custom numFmt pattern', () => {
    const custom = '<w:numFmt w:val="custom" w:format="001, 002, 003, ..."/>';
    expect(render({ kind: 'footnote', sectionPr: [custom, custom] })).toEqual(['001', '002', '003', '004']);
    expect(formatCustomNoteNumber(7, '０１, ０２, ...')).toBe('０７');
    expect(formatCustomNoteNumber(4, 'A, B, C, ...')).toBe('D');
    expect(formatCustomNoteNumber(12, 'x')).toBe('12');
  });

  it('ignores a numStart beyond the Word unsignedShort range instead of failing', () => {
    const huge = '<w:numStart w:val="5000000"/>';
    expect(render({ kind: 'endnote', sectionPr: [huge, huge] })).toEqual(['i', 'ii', 'iii', 'iv']);
    // The largest Word start still formats within the number-format budget.
    const largest = '<w:numStart w:val="65535"/>';
    expect(render({ kind: 'endnote', sectionPr: [largest, largest] }).join('').startsWith('m'.repeat(65)))
      .toBe(true);
  });
});

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
    ['word/document.xml', `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`],
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
    expect(render({ kind: 'footnote', documentPr: '<w:numRestart w:val="eachSect"/>',
      sectionPr: ['', '<w:numFmt w:val="lowerLetter"/>'] })).toEqual(['1', '2', 'a', 'b']);
  });

  it('formats a custom numFmt pattern, also inside mc:AlternateContent', () => {
    expect(render({ kind: 'footnote', documentPr: '<mc:AlternateContent><mc:Choice Requires="w14">'
      + '<w:numFmt w:val="custom" w:format="001, 002, 003, ..."/></mc:Choice>'
      + '<mc:Fallback><w:numFmt w:val="decimal"/></mc:Fallback></mc:AlternateContent>' }))
      .toEqual(['001', '002', '003', '004']);
    expect(formatCustomNoteNumber(7, '０１, ０２, ...')).toBe('０７');
    expect(formatCustomNoteNumber(4, 'A, B, C, ...')).toBe('D');
    expect(formatCustomNoteNumber(12, 'x')).toBe('12');
  });

  it('ignores a numStart beyond the Word unsignedShort range instead of failing', () => {
    expect(render({ kind: 'endnote', documentPr: '<w:numStart w:val="5000000"/>' })).toEqual(['i', 'ii', 'iii', 'iv']);
    // The largest Word start still formats within the number-format budget.
    expect(render({ kind: 'endnote', documentPr: '<w:numStart w:val="65535"/>' }).join('').startsWith('m'.repeat(65)))
      .toBe(true);
  });
});

/** Six footnote references, two per page (hard page breaks), optional
 * document-wide footnotePr children. */
function pagedControl(documentPr: string, sectionPr = ''): DocxDocumentModel {
  const ref = (id: number) => `<w:r><w:footnoteReference w:id="${id}"/></w:r>`;
  const paragraphs = [1, 2, 3, 4, 5, 6].map((id) => `<w:p><w:r><w:t>P${id}</w:t></w:r>${ref(id)}`
    + `${id % 2 === 0 && id < 6 ? '<w:r><w:br w:type="page"/></w:r>' : ''}</w:p>`).join('');
  const notes = [1, 2, 3, 4, 5, 6].map((id) => `<w:footnote w:id="${id}"><w:p><w:r><w:footnoteRef/></w:r>`
    + `<w:r><w:t>N${id}</w:t></w:r></w:p></w:footnote>`).join('');
  const parts = new Map<string, string>([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
      + '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'],
    ['_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="d" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="n" Type="${R}/footnotes" Target="footnotes.xml"/><Relationship Id="s" Type="${R}/settings" Target="settings.xml"/></Relationships>`],
    ['word/settings.xml', `<w:settings xmlns:w="${W}"><w:footnotePr>${documentPr}</w:footnotePr></w:settings>`],
    ['word/document.xml', `<w:document xmlns:w="${W}"><w:body>${paragraphs}<w:sectPr>`
      + `${sectionPr ? `<w:footnotePr>${sectionPr}</w:footnotePr>` : ''}<w:pgSz w:w="12240" w:h="15840"/>`
      + '<w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr></w:body></w:document>'],
    ['word/footnotes.xml', `<w:footnotes xmlns:w="${W}">${notes}</w:footnotes>`],
  ]);
  const archive = new DocxArchive(storeZip(new Map([...parts].map(([name, xml]) => [name, encode(xml)]))));
  try {
    return normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally {
    archive.free();
  }
}

/** Body reference labels and in-note number labels, per page. */
function pagedLabels(model: DocxDocumentModel): Array<{ body: string[]; notes: string[] }> {
  const layout = layoutDocument(model, createLayoutServices(model, { measureContext: context }), { currentDateMs: 0 });
  return layout.pages.map((page) => {
    const collect = (nodes: readonly PaintNode[], out: string[]) => {
      for (const node of nodes) {
        if (node.kind === 'paragraph') {
          for (const line of node.lines) for (const placement of line.placements) {
            if (placement.kind === 'text' && placement.noteReference && placement.text.length > 0) out.push(placement.text);
          }
        } else if (node.kind === 'note') {
          collect(node.story.blocks, out);
        }
      }
      return out;
    };
    return { body: collect(page.layers.body, []), notes: collect(page.layers.notes as readonly PaintNode[], []) };
  });
}

describe('numRestart eachPage (pagination fixed point)', () => {
  it('restarts footnote numbering on every page, in references and notes', () => {
    expect(pagedLabels(pagedControl('<w:numRestart w:val="eachPage"/>'))).toEqual([
      { body: ['1', '2'], notes: ['1', '2'] },
      { body: ['1', '2'], notes: ['1', '2'] },
      { body: ['1', '2'], notes: ['1', '2'] },
    ]);
  });

  it('restarts at numStart in the page format, also from a section override', () => {
    expect(pagedLabels(pagedControl('<w:numFmt w:val="lowerRoman"/><w:numStart w:val="3"/><w:numRestart w:val="eachPage"/>'))
      .map((page) => page.body)).toEqual([['iii', 'iv'], ['iii', 'iv'], ['iii', 'iv']]);
    expect(pagedLabels(pagedControl('', '<w:numRestart w:val="eachPage"/>')).map((page) => page.body))
      .toEqual([['1', '2'], ['1', '2'], ['1', '2']]);
  });

  it('keeps continuous numbering without eachPage', () => {
    expect(pagedLabels(pagedControl('')).map((page) => page.body)).toEqual([['1', '2'], ['3', '4'], ['5', '6']]);
  });
});


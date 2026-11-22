import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import type { DocxDocumentModel } from './types.js';
import type { DocumentLayout, NoteLayout } from './layout/types.js';
import {
  wordNoteSeparatorRuleDropPt, wordNoteSeparatorRuleWidthPt, wordShortNoteSeparatorLengthPt,
} from './layout/body-pagination-compatibility.js';

// WORD_NOTE_SEPARATOR_GEOMETRY: synthetic reductions of the Windows Word
// separator controls. Synthetic glyph metrics are no font oracle; only the
// relations Word showed are asserted.
beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

const context = {
  font: '10px Arial', letterSpacing: '0px', fontKerning: 'none',
  measureText: (text: string) => {
    const px = parseFloat(/(\d+(?:\.\d+)?)px/.exec(context.font)?.[1] ?? '10');
    return {
      width: [...text].length * px * 0.6,
      fontBoundingBoxAscent: px * 0.8, fontBoundingBoxDescent: px * 0.2,
      actualBoundingBoxAscent: px * 0.8, actualBoundingBoxDescent: px * 0.2,
    };
  },
} as unknown as CanvasRenderingContext2D;

type Kind = 'footnote' | 'endnote';
interface Control {
  readonly kind: Kind;
  readonly listed?: boolean;
  readonly markHalfPoints?: number;
  readonly noSeparatorStories?: boolean;
  readonly textWidthPt?: number;
  readonly noteLines?: number;
}

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const encode = (xml: string) => new TextEncoder().encode(xml);

function control({ kind, listed = false, markHalfPoints, noSeparatorStories = false, textWidthPt = 468, noteLines = 2 }: Control): DocxDocumentModel {
  const rPr = markHalfPoints ? `<w:rPr><w:sz w:val="${markHalfPoints}"/></w:rPr>` : '';
  const special = (role: string, id: number, mark: string) => `<w:${kind} w:type="${role}" w:id="${id}">`
    + `<w:p><w:pPr>${rPr}</w:pPr><w:r>${rPr}<w:${mark}/></w:r></w:p></w:${kind}>`;
  const stories = noSeparatorStories ? ''
    : special('separator', -1, 'separator') + special('continuationSeparator', 0, 'continuationSeparator');
  const lines = Array.from({ length: noteLines }, (_, index) => `<w:p><w:r><w:t>N${index}</w:t></w:r></w:p>`).join('');
  const right = Math.round((612 - 72 - textWidthPt) * 20);
  const parts = new Map<string, string>([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + `<Override PartName="/word/${kind}s.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}s+xml"/>`
      + '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'],
    ['_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="d" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="n" Type="${R}/${kind}s" Target="${kind}s.xml"/><Relationship Id="s" Type="${R}/settings" Target="settings.xml"/>`
      + `<Relationship Id="t" Type="${R}/styles" Target="styles.xml"/></Relationships>`],
    ['word/styles.xml', `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>`
      + '<w:sz w:val="20"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
      + '</w:pPrDefault></w:docDefaults></w:styles>'],
    ['word/settings.xml', `<w:settings xmlns:w="${W}">`
      + (listed ? `<w:${kind}Pr><w:${kind} w:id="-1"/><w:${kind} w:id="0"/></w:${kind}Pr>` : '')
      + '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'],
    ['word/document.xml', `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>BODY</w:t></w:r>`
      + `<w:r><w:${kind}Reference w:id="1"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/>`
      + `<w:pgMar w:top="720" w:right="${right}" w:bottom="720" w:left="1440" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr></w:body></w:document>`],
    [`word/${kind}s.xml`, `<w:${kind}s xmlns:w="${W}">${stories}<w:${kind} w:id="1">${lines}</w:${kind}></w:${kind}s>`],
  ]);
  const archive = new DocxArchive(storeZip(new Map([...parts].map(([name, xml]) => [name, encode(xml)]))));
  try {
    return normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally {
    archive.free();
  }
}

const render = (model: DocxDocumentModel): DocumentLayout => layoutDocument(model,
  createLayoutServices(model, { measureContext: context, allowFootnoteContinuation: true }), { currentDateMs: 0 });

function bands(layout: DocumentLayout) {
  return layout.pages.flatMap((page) => page.layers.notes)
    .filter((node): node is NoteLayout => node.kind === 'note' && node.separator.length > 0)
    .map((note) => {
      const rule = note.separator[0]!;
      return {
        heightPt: note.flowBounds.heightPt - note.story.advancePt,
        ruleOffsetPt: rule.from.yPt - note.flowBounds.yPt,
        ruleLengthPt: rule.to.xPt - rule.from.xPt,
        ruleWidthPt: rule.widthPt,
      };
    });
}

describe('word-note-separator-geometry', () => {
  it('keeps the helper geometry of the Windows controls', () => {
    expect([100, 144, 200, 467.7].map(wordShortNoteSeparatorLengthPt)).toEqual([100, 144, 144, 144]);
    expect(wordNoteSeparatorRuleWidthPt(10)).toBe(0.5);
    expect(wordNoteSeparatorRuleWidthPt(20)).toBe(1);
    expect(wordNoteSeparatorRuleDropPt(10)).toBe(1.25);
  });

  it.each(['footnote', 'endnote'] as const)('lays out the default %s separator for unlisted or missing stories', (kind) => {
    const bare = bands(render(control({ kind })))[0]!;
    // One Normal-style mark line with the rule 10pt/8 below its middle.
    expect(bare.ruleOffsetPt).toBeCloseTo(bare.heightPt / 2 + 1.25, 5);
    expect(bare.ruleLengthPt).toBe(144);
    expect(bare.ruleWidthPt).toBe(0.5);
    // An unlisted 20pt story and a missing story change nothing.
    expect(bands(render(control({ kind, markHalfPoints: 40 })))[0]).toEqual(bare);
    expect(bands(render(control({ kind, noSeparatorStories: true })))[0]).toEqual(bare);
    // A listed 20pt story sizes the band, the rule drop and thickness.
    const listed = bands(render(control({ kind, listed: true, markHalfPoints: 40 })))[0]!;
    expect(listed.heightPt).toBeGreaterThan(bare.heightPt);
    expect(listed.ruleOffsetPt).toBeCloseTo(listed.heightPt / 2 + 2.5, 5);
    expect(listed.ruleWidthPt).toBe(1);
    // The short rule never exceeds the note area.
    expect(bands(render(control({ kind, textWidthPt: 100 })))[0]!.ruleLengthPt).toBe(100);
  });

  it('draws the default continuation rule short for footnotes and full for endnotes', () => {
    for (const [kind, width, expected] of [
      ['footnote', 468, 144], ['footnote', 200, 144], ['endnote', 468, 468], ['endnote', 200, 200],
    ] as const) {
      const continued = bands(render(control({ kind, textWidthPt: width, noteLines: 120 })));
      expect(continued.length).toBeGreaterThan(1);
      expect(continued[1]!.ruleLengthPt).toBe(expected);
    }
  });
});

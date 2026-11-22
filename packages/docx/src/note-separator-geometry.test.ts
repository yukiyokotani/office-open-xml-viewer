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
  noteSeparatorRuleSpan, wordNoteSeparatorRuleDropPt, wordNoteSeparatorRuleWidthPt, wordShortNoteSeparatorLengthPt,
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
  /** Size on the marker run alone (its paragraph mark stays 10pt). */
  readonly markerRunHalfPoints?: number;
  /** The listed separator story's paragraph, replacing the mark-only one. */
  readonly separatorParagraph?: string;
  /** Extra sectPr children (columns, bidi). */
  readonly sectExtra?: string;
  /** Paragraph properties of the Normal style. */
  readonly normalPPr?: string;
}

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const encode = (xml: string) => new TextEncoder().encode(xml);

function control({ kind, listed = false, markHalfPoints, noSeparatorStories = false, textWidthPt = 468, noteLines = 2, markerRunHalfPoints, separatorParagraph, sectExtra = '', normalPPr }: Control): DocxDocumentModel {
  const rPr = markHalfPoints ? `<w:rPr><w:sz w:val="${markHalfPoints}"/></w:rPr>` : '';
  const runRPr = markerRunHalfPoints ? `<w:rPr><w:sz w:val="${markerRunHalfPoints}"/></w:rPr>` : rPr;
  const special = (role: string, id: number, mark: string) => `<w:${kind} w:type="${role}" w:id="${id}">`
    + (role === 'separator' && separatorParagraph ? separatorParagraph
      : `<w:p><w:pPr>${rPr}</w:pPr><w:r>${runRPr}<w:${mark}/></w:r></w:p>`) + `</w:${kind}>`;
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
      + '</w:pPrDefault></w:docDefaults>'
      + (normalPPr ? `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr>${normalPPr}</w:pPr></w:style>` : '')
      + '</w:styles>'],
    ['word/settings.xml', `<w:settings xmlns:w="${W}">`
      + (listed ? `<w:${kind}Pr><w:${kind} w:id="-1"/><w:${kind} w:id="0"/></w:${kind}Pr>` : '')
      + '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'],
    ['word/document.xml', `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>BODY</w:t></w:r>`
      + `<w:r><w:${kind}Reference w:id="1"/></w:r></w:p><w:sectPr>${sectExtra}<w:pgSz w:w="12240" w:h="15840"/>`
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

  it('keeps the rule inside the note area for page-wide, right-to-left and other producers', () => {
    const columns = [{ xPt: 72, wPt: 120 }, { xPt: 212, wPt: 120 }];
    const span = (overrides: Partial<Parameters<typeof noteSeparatorRuleSpan>[0]>) => noteSeparatorRuleSpan({
      word: true, kind: 'footnote', mode: 'short', continuing: false, pageWide: true, rtl: false,
      container: { xPt: 72, widthPt: 120 }, columns, ...overrides,
    });
    // Mode 15: the page-wide area spans both columns (72-332).
    expect(span({})).toEqual({ fromXPt: 72, toXPt: 216 });
    // Right to left: the rule starts at the area's right edge, never past it.
    expect(span({ rtl: true, container: { xPt: 212, widthPt: 120 } })).toEqual({ fromXPt: 188, toXPt: 332 });
    // A page-wide endnote continuation spans the whole area, not one column.
    expect(span({ kind: 'endnote', mode: 'default', continuing: true })).toEqual({ fromXPt: 72, toXPt: 332 });
    // Per-column bands keep their column.
    expect(span({ pageWide: false })).toEqual({ fromXPt: 72, toXPt: 192 });
    // Another producer keeps the library policy: a third of the column, and
    // a story-less continuation spans it.
    expect(span({ word: false, container: { xPt: 72, widthPt: 300 } })).toEqual({ fromXPt: 72, toXPt: 172 });
    expect(span({ word: false, mode: 'default', continuing: true, container: { xPt: 72, widthPt: 300 } }))
      .toEqual({ fromXPt: 72, toXPt: 372 });
  });

  it.each(['footnote', 'endnote'] as const)('sizes a %s band and rule from a marker run formatted unlike its mark', (kind) => {
    const bare = bands(render(control({ kind, listed: true })))[0]!;
    const marker = bands(render(control({ kind, listed: true, markerRunHalfPoints: 40 })))[0]!;
    // The 20pt marker run shares the 10pt mark's line and sizes the rule.
    expect(marker.heightPt).toBeGreaterThan(bare.heightPt);
    expect(marker.ruleWidthPt).toBe(1);
    expect(marker.ruleOffsetPt).toBeCloseTo(marker.heightPt / 2 + 2.5, 5);
    expect(marker.ruleLengthPt).toBe(144);
  });

  it('lays a listed separator out in a right-to-left mode-15 page-wide area, rule inside it', () => {
    // Two 120pt columns and a 20pt gap (a 260pt area at 72-332pt).
    const layout = render(control({
      kind: 'footnote', listed: true, textWidthPt: 260,
      sectExtra: '<w:cols w:num="2" w:space="400"/><w:bidi/>',
      separatorParagraph: '<w:p><w:r><w:t>A</w:t></w:r><w:r><w:separator/></w:r></w:p>',
    }));
    const note = layout.pages[0]!.layers.notes.find((node): node is NoteLayout => node.kind === 'note')!;
    const rule = note.leading!.rule!.segment;
    expect(rule.from.xPt).toBeGreaterThanOrEqual(72);
    expect(rule.to.xPt).toBeLessThanOrEqual(332);
    expect(rule.to.xPt - rule.from.xPt).toBe(144);
    // The paragraph starts at the area's left edge: its text precedes the rule.
    expect(rule.from.xPt).toBeCloseTo(72 + 6, 5);
  });

  it('keeps separator text in the marker run and around line breaks', () => {
    const text = (separatorParagraph: string) => {
      const layout = render(control({ kind: 'footnote', listed: true, separatorParagraph }));
      const note = layout.pages[0]!.layers.notes.find((node): node is NoteLayout => node.kind === 'note')!;
      const paragraph = note.leading?.paragraph;
      return {
        text: paragraph?.kind === 'paragraph' ? paragraph.lines.map((line) => line.placements
          .flatMap((placement) => placement.kind === 'text' ? [placement.text.replace(/\u200b/gu, '')] : []).join('')) : [],
        rule: note.leading?.rule !== undefined,
      };
    };
    expect(text('<w:p><w:r><w:t>BEFORE</w:t><w:separator/><w:t>AFTER</w:t></w:r></w:p>'))
      .toEqual({ text: ['BEFOREAFTER'], rule: true });
    expect(text('<w:p><w:r><w:t>A</w:t><w:br/><w:t>B</w:t></w:r><w:r><w:separator/></w:r></w:p>'))
      .toEqual({ text: ['A', 'B'], rule: true });
  });

  it('keeps the default band when the Normal style has a paragraph border', () => {
    const plain = bands(render(control({ kind: 'endnote' })))[0]!;
    const bordered = bands(render(control({
      kind: 'endnote', normalPPr: '<w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="000000"/></w:pBdr>',
    })))[0]!;
    expect(bordered).toEqual(plain);
  });

  const leadingOf = (layout: DocumentLayout) => layout.pages[0]!.layers.notes
    .find((node): node is NoteLayout => node.kind === 'note')!;
  const paragraphTexts = (note: NoteLayout) => [note.leading?.paragraph, ...(note.leading?.followingParagraphs ?? [])]
    .flatMap((paragraph) => paragraph?.kind === 'paragraph' ? [paragraph.lines.map((line) => line.placements
      .flatMap((placement) => placement.kind === 'text' ? [placement.text] : []).join('')).join('|')] : []);

  it('renders a listed story without a marker, and text in other paragraphs', () => {
    const plain = leadingOf(render(control({ kind: 'footnote', listed: true,
      separatorParagraph: '<w:p><w:r><w:t>NOTES</w:t></w:r></w:p>' })));
    expect(paragraphTexts(plain)).toEqual(['NOTES']);
    expect(plain.leading?.rule).toBeUndefined();
    expect(plain.separator).toEqual([]);
    const two = leadingOf(render(control({ kind: 'footnote', listed: true,
      separatorParagraph: '<w:p><w:r><w:t>NOTES</w:t></w:r></w:p><w:p><w:r><w:t>A</w:t></w:r><w:r><w:separator/></w:r></w:p>' })));
    expect(paragraphTexts(two).map((text) => text.replace(/\u200b/gu, ''))).toEqual(['NOTES', 'A']);
    expect(two.leading?.rule?.segment.from.yPt).toBeGreaterThan(two.leading!.paragraph!.flowBounds.yPt
      + two.leading!.paragraph!.flowBounds.heightPt);
  });

  it('reads runs of a second WordprocessingML prefix and authored U+FFFC/U+200B as text', () => {
    const aliased = leadingOf(render(control({ kind: 'footnote', listed: true,
      separatorParagraph: `<w:p xmlns:a="${W}"><w:r><w:t>W1</w:t><w:separator/><w:t>W2</w:t></w:r><a:r><a:t>A1</a:t></a:r></w:p>` })));
    expect(paragraphTexts(aliased).map((text) => text.replace(/\u200b/gu, ''))).toEqual(['W1W2A1']);
    const authored = leadingOf(render(control({ kind: 'footnote', listed: true,
      separatorParagraph: '<w:p><w:r><w:t>&#xFFFC;</w:t></w:r><w:r><w:separator/></w:r><w:r><w:t>AFTER</w:t></w:r></w:p>' })));
    expect(paragraphTexts(authored)[0]).toContain('\ufffc');
    expect(paragraphTexts(authored)[0]).toContain('AFTER');
    expect(authored.leading?.rule).toBeDefined();
    // A 10pt authored U+200B beside a 20pt marker run: the marker sizes the rule.
    const zwsp = leadingOf(render(control({ kind: 'footnote', listed: true,
      separatorParagraph: '<w:p><w:r><w:t>&#x200B;</w:t></w:r><w:r><w:rPr><w:sz w:val="40"/></w:rPr><w:separator/></w:r></w:p>' })));
    expect(zwsp.leading?.rule?.segment.widthPt ?? zwsp.separator[0]?.widthPt).toBe(1);
  });

  it('keeps a short rule inside the note area under a large indent', () => {
    const note = leadingOf(render(control({ kind: 'footnote', listed: true, textWidthPt: 100,
      separatorParagraph: '<w:p><w:pPr><w:ind w:left="3000"/></w:pPr><w:r><w:rPr><w:sz w:val="40"/></w:rPr><w:separator/></w:r></w:p>' })));
    const rule = note.leading?.rule?.segment ?? note.separator[0]!;
    expect(rule.from.xPt).toBeGreaterThanOrEqual(72);
    expect(rule.from.xPt).toBeLessThanOrEqual(rule.to.xPt);
    expect(rule.to.xPt).toBeLessThanOrEqual(172);
  });

  it('keeps the default band when Normal has automatic spacing after', () => {
    const plain = bands(render(control({ kind: 'endnote' })))[0]!;
    const automatic = bands(render(control({ kind: 'endnote', normalPPr: '<w:spacing w:afterAutospacing="1"/>' })))[0]!;
    expect(automatic).toEqual(plain);
  });
});

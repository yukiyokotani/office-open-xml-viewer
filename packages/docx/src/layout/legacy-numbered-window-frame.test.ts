/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from '../wasm/docx_parser.js';
import { storeZip } from '../conformance/generate.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import { layoutDocument } from '../document-layout.js';
import { createLayoutServices } from '../layout-runtime.js';
import { measureContext } from '../test-support/tab-fitting.test-support.js';
import type { DocSettings } from '../types.js';

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('../wasm/docx_parser_bg.wasm', import.meta.url)) });
});

// Parsed DrawingML rectangles exercise canonical ownership, rather than loose
// public-model shapes, which intentionally retain the established frame.
function run({ left = 50, hanging = 50, floats = [[0, 150, 60]], stop, text = 'word', settings = { compatibilityMode: 14 }, external = false }:
  { left?: number; hanging?: number; floats?: number[][]; stop?: number; text?: string; settings?: DocSettings; external?: boolean } = {}) {
  const emu = (pt: number) => Math.round(pt * 12700);
  const shapes = floats.map(([x, width, height], i) => `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${i + 1}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>${emu(x!)}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="${emu(width!)}" cy="${emu(height!)}"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="${i + 1}" name="Rectangle"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(width!)}" cy="${emu(height!)}"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`).join('');
  const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';
  const ppr = '<w:spacing w:before="0" w:after="0"/>';
  const body = `${external ? `<w:p><w:pPr>${ppr}</w:pPr>${shapes}<w:r><w:t>host</w:t></w:r></w:p>` : ''}<w:p><w:pPr>${ppr}<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>${external ? '' : shapes}<w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:sz w:val="24"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
  const compat = Object.entries(settings).map(([key, value]) => key === 'compatibilityMode'
    ? `<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${value}"/>`
    : `<w:${key} w:val="${value ? 1 : 0}"/>`).join('');
  const parts = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
    '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    'word/_rels/document.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="n" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>',
    'word/document.xml': `<w:document ${ns}><w:body>${body}<w:sectPr><w:pgSz w:w="10438" w:h="8000"/><w:pgMar w:top="400" w:right="400" w:bottom="400" w:left="400" w:header="0" w:footer="0"/></w:sectPr></w:body></w:document>`,
    'word/numbering.xml': `<w:numbering ${ns}><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr>${stop === undefined ? '' : `<w:tabs><w:tab w:val="left" w:pos="${stop * 20}"/></w:tabs>`}<w:ind w:left="${left * 20}" w:hanging="${hanging * 20}"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`,
    'word/settings.xml': `<w:settings ${ns}><w:compat>${compat}</w:compat></w:settings>`,
  };
  const archive = new DocxArchive(storeZip(new Map(Object.entries(parts).map(([name, xml]) => [name, new TextEncoder().encode(xml)]))));
  let model;
  try { model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document; }
  finally { archive.free(); }
  const layout = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext(5) }), { currentDateMs: 0 });
  const p = layout.pages[0]!.layers.body.find(p => p.kind === 'paragraph' && p.lines.some(line => line.placements.some(p => p.kind === 'text' && p.role === 'numbering-marker')));
  if (p?.kind !== 'paragraph') throw new Error('Missing original numbered paragraph');
  const line = p.lines[0]!;
  const marker = line.placements.find(p => p.kind === 'text' && p.role === 'numbering-marker');
  const bodyText = line.placements.find(p => p.kind === 'text' && p.role !== 'numbering-marker');
  if (marker?.kind !== 'text' || bodyText?.kind !== 'text') throw new Error('Missing marker or source body');
  // The authored page margin is 400 twips = 20pt, exactly as the old fixture.
  return { marker: marker.origin.xPt - 20, body: bodyText.origin.xPt - 20, top: line.bounds.yPt - 20 };
}

describe('legacy numbered suffix window frame', () => {
  it('rebases automatic and authored stops without rebasing the implicit hanging stop', () => {
    expect(run()).toMatchObject({ marker: 150, body: 186 });
    expect(run({ settings: {} })).toMatchObject({ marker: 150, body: 186 });
    expect(run({ settings: { compatibilityMode: 14, doNotUseIndentAsNumberingTabStop: false, noTabHangInd: false } }))
      .toMatchObject({ marker: 150, body: 186 });
    expect(run({ stop: 200 })).toMatchObject({ marker: 150, body: 350 });
    expect(run({ floats: [[0, 30, 60]] })).toMatchObject({ marker: 30, body: 50 });
  });

  it('keeps a column-owned window at and before the authored first-line start', () => {
    expect(run({ left: 200, hanging: 50 })).toMatchObject({ marker: 300, body: 330 });
    expect(run({ left: 200, hanging: 50, floats: [[0, 30, 60]] })).toMatchObject({ marker: 180, body: 200 });
  });

  it('admits a float gap only after its own marker and suffix leave room for the head', () => {
    const result = run({ floats: [[100, 50, 60], [250, 50, 60], [450, 31.9, 60]], text: 'W'.repeat(24) });
    expect(result.top).toBeGreaterThanOrEqual(60);
    expect(result).toMatchObject({ marker: 0, body: 50 });
  });

  it.each(['doNotSuppressIndentation', 'ww11IndentRules', 'doNotUseIndentAsNumberingTabStop', 'noTabHangInd'] as const)('does not apply the false-flag frame when %s is authored true', flag => {
    expect(run({ stop: 200, settings: { compatibilityMode: 14, [flag]: true } }))
      .toMatchObject({ marker: 150, body: 200 });
  });

  it('preserves external-float geometry instead of borrowing self-owned authority', () => {
    const actual = run({ external: true, stop: 200 });
    const established = run({ external: true, stop: 200, settings: { compatibilityMode: 14, doNotSuppressIndentation: true } });
    expect(actual).toEqual(established);
  });

  it('preserves the established East Asian hanging frame outside the measured Latin class', () => {
    expect(run({ text: '漢。', stop: 200 })).toEqual(run({ text: '漢。', stop: 200,
      settings: { compatibilityMode: 14, doNotSuppressIndentation: true } }));
  });

  it('keeps the modern and no-float numbered frames unchanged', () => {
    expect(run({ floats: [] })).toMatchObject({ marker: 0, body: 50 });
    expect(run({ settings: { compatibilityMode: 15 }, stop: 200 })).toMatchObject({ marker: 150, body: 200 });
  });
});

/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from '../wasm/docx_parser.js';
import { measureContext } from '../test-support/tab-fitting.test-support.js';
import { storeZip } from '../conformance/generate.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import { layoutDocument } from '../document-layout.js';
import { createLayoutServices } from '../layout-runtime.js';

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('../wasm/docx_parser_bg.wasm', import.meta.url)) });
});

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';

// A 144pt square-wrapped picture at the 54pt margin, distR 9pt: wrap edge 207pt.
const float = '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="1828800" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="F"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="F"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdMissing"/></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';

function bytes(ind: string, levelTabs = '', styleOrigin?: string, variant: { empty?: boolean; rtl?: boolean; defaultsPPr?: string } = {}): Uint8Array {
  const enc = new TextEncoder();
  const p = (body: string, ppr = '') => `<w:p><w:pPr>${ppr}<w:spacing w:before="0" w:after="0"/></w:pPr>${body}</w:p>`;
  const item = p(variant.empty ? '' : '<w:r><w:t>word</w:t></w:r>', (styleOrigin === undefined
    ? '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' : '<w:pStyle w:val="ListItem"/>')
    + (variant.rtl ? '<w:bidi/>' : ''));
  const floatXml = variant.rtl
    ? float.replace('<wp:posOffset>0</wp:posOffset></wp:positionH>', '<wp:align>right</wp:align></wp:positionH>')
    : float;
  const styles = `<w:styles ${W}><w:docDefaults><w:pPrDefault><w:pPr>${variant.defaultsPPr ?? ''}</w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="ListItem"><w:name w:val="List Item"/><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>${styleOrigin ?? ''}</w:pPr></w:style></w:styles>`;
  const doc = `<w:document ${W}><w:body>${p(`${floatXml}<w:r><w:t>host</w:t></w:r>`)}${item}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  const numbering = `<w:numbering ${W}><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr>${levelTabs}${ind}</w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
  return storeZip(new Map([
    ['[Content_Types].xml', enc.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>')],
    ['_rels/.rels', enc.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/_rels/document.xml.rels', enc.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')],
    ['word/document.xml', enc.encode(doc)],
    ['word/numbering.xml', enc.encode(numbering)],
    ['word/styles.xml', enc.encode(styles)],
  ]));
}

function firstItemLine(input: Uint8Array) {
  const archive = new DocxArchive(input);
  let model;
  try {
    model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally { archive.free(); }
  const layout = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext(5) }), { currentDateMs: 0 });
  const paragraphs = layout.pages[0]!.layers.body.filter((node) => node.kind === 'paragraph');
  const line = paragraphs[1]!.kind === 'paragraph' ? paragraphs[1]!.lines[0]! : undefined;
  const texts = line!.placements.filter((node) => node.kind === 'text');
  const body = texts.find((node) => node.role !== 'numbering-marker')?.bounds;
  return {
    marker: texts.find((node) => node.role === 'numbering-marker')?.bounds.xPt,
    body: body?.xPt,
    ...(body ? { bodyEnd: body.xPt + body.widthPt } : {}),
  };
}

describe('numbered first line beside a square-wrapped float (#1883)', () => {
  // Windows Word 16.0 places the marker at max(max(indent, wrap) + firstLine,
  // wrap) and advances its suffix tab to the paragraph's absolute stops; the
  // fake measurer gives the "1." marker 10pt.
  it.each([
    // left 54 hanging 36: marker at the wrap edge, body at the 234pt default stop.
    { ind: '<w:ind w:left="1080" w:hanging="720"/>', tabs: '', marker: 207, body: 234 },
    // left 36 firstLine 18: the positive first-line indent stays an inset; the
    // marker ends at 235pt, past the 234pt stop.
    { ind: '<w:ind w:left="720" w:firstLine="360"/>', tabs: '', marker: 225, body: 270 },
    // A level w:ind without hanging/firstLine has no first-line indent.
    { ind: '<w:ind w:left="0"/>', tabs: '', marker: 207, body: 234 },
    // The level's own custom stop (216pt from the margin) wins over defaults.
    { ind: '<w:ind w:left="1080" w:hanging="720"/>', tabs: '<w:tabs><w:tab w:val="left" w:pos="4320"/></w:tabs>', marker: 207, body: 270 },
  ])('$ind $tabs', ({ ind, tabs, marker, body }) => {
    expect(firstItemLine(bytes(ind, tabs))).toMatchObject({ marker, body });
  });

  it('lets a paragraph style clear a style-origin level tab', () => {
    // §17.7.2: style-origin numbering sits beneath the paragraph style, whose
    // clear at 216pt removes the level's stop there; the suffix tab then uses
    // the default stops.
    const levelTab = '<w:tabs><w:tab w:val="left" w:pos="4320"/></w:tabs>';
    const ind = '<w:ind w:left="1080" w:hanging="720"/>';
    expect(firstItemLine(bytes(ind, levelTab, '')).body).toBe(270);
    expect(firstItemLine(bytes(ind, levelTab, '<w:tabs><w:tab w:val="clear" w:pos="4320"/></w:tabs>')).body).toBe(234);
  });

  it('places an empty item\'s marker at the moved edge', () => {
    expect(firstItemLine(bytes('<w:ind w:left="1080" w:hanging="720"/>', '', undefined, { empty: true })).marker).toBe(207);
  });

  it('mirrors the suffix tab for an RTL item beside a right float', () => {
    // The 612pt page mirrors the LTR geometry: the body ends where the LTR
    // body (234pt) starts, measured from the right page edge.
    const line = firstItemLine(bytes('<w:ind w:left="1080" w:hanging="720"/>', '', undefined, { rtl: true }));
    expect(line.marker! + 10).toBe(612 - 207);
    expect(line.bodyEnd).toBe(612 - 234);
  });

  it.each([
    // A document-default clear sits beneath style-origin numbering: the level stop survives.
    { defaults: '<w:tabs><w:tab w:val="clear" w:pos="4320"/></w:tabs>', level: '<w:tabs><w:tab w:val="left" w:pos="4320"/></w:tabs>', body: 270 },
    // A level clear removes a document-default stop beneath it.
    { defaults: '<w:tabs><w:tab w:val="left" w:pos="4320"/></w:tabs>', level: '<w:tabs><w:tab w:val="clear" w:pos="4320"/></w:tabs>', body: 234 },
  ])('orders numbering tabs above document defaults ($body)', ({ defaults, level, body }) => {
    const ind = '<w:ind w:left="1080" w:hanging="720"/>';
    expect(firstItemLine(bytes(ind, level, '', { defaultsPPr: defaults })).body).toBe(body);
  });
});

/** A square-wrapped shape anchored in the paragraph, `x`/`w`/`h` in pt from
 * the column and the paragraph top (zero wrap distances). */
function squareShape(id: number, x: number, w: number, h: number): string {
  const emu = (pt: number) => Math.round(pt * 12700);
  return `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>${emu(x)}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="${emu(w)}" cy="${emu(h)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="${id}" name="S${id}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="S${id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdMissing"/></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
}

/** One paragraph that anchors `floats` and holds `text`, numbered with the
 * literal marker `marker` (tab suffix) when `numbered`; the 504 pt column
 * starts 54 pt from the page edge. Returns its first two lines. */
function anchoredParagraphLines(floats: string, text: string | readonly string[], numbered: { marker: string; ind: string } | null,
  { compatibilityMode, advancePt = 5, anchorsLast = false }:
    { compatibilityMode?: number | undefined; advancePt?: number; anchorsLast?: boolean } = {}) {
  const enc = new TextEncoder();
  const ppr = numbered ? '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' : '';
  // An array is one run per element, alternately plain and bold.
  const runs = (typeof text === 'string' ? [text] : text)
    .map((piece, index) => `<w:r>${index % 2 === 1 ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t>${piece}</w:t></w:r>`).join('');
  const doc = `<w:document ${W}><w:body><w:p><w:pPr>${ppr}<w:spacing w:before="0" w:after="0"/></w:pPr>${anchorsLast ? `${runs}${floats}` : `${floats}${runs}`}</w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  const numbering = `<w:numbering ${W}><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="${numbered?.marker ?? '1.'}"/><w:lvlJc w:val="left"/><w:pPr>${numbered?.ind ?? ''}</w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
  const input = storeZip(new Map([
    ['[Content_Types].xml', enc.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')],
    ['_rels/.rels', enc.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/_rels/document.xml.rels', enc.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')],
    ['word/document.xml', enc.encode(doc)],
    ['word/numbering.xml', enc.encode(numbering)],
    ['word/settings.xml', enc.encode(`<w:settings ${W}>${compatibilityMode === undefined ? '' : `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${compatibilityMode}"/></w:compat>`}</w:settings>`)],
  ]));
  const archive = new DocxArchive(input);
  let model;
  try {
    model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally { archive.free(); }
  const layout = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext(advancePt) }), { currentDateMs: 0 });
  const paragraph = layout.pages[0]!.layers.body.find((node) => node.kind === 'paragraph');
  if (paragraph?.kind !== 'paragraph') throw new Error('no paragraph');
  return paragraph.lines.slice(0, 2).map((line) => {
    const texts = line.placements.filter((node) => node.kind === 'text');
    const body = texts.find((node) => node.role !== 'numbering-marker');
    return {
      topPt: line.bounds.yPt - paragraph.flowBounds.yPt,
      marker: texts.find((node) => node.role === 'numbering-marker')?.bounds.xPt,
      body: body?.bounds.xPt,
      text: texts.filter((node) => node.role !== 'numbering-marker').map((node) => node.text).join(''),
    };
  });
}

describe('numbered first line across float gaps (#2055, word-numbered-first-line-gap-fit)', () => {
  // Square floats 60pt tall leave the gaps [0,100], [150,250] and [300,450]
  // of the column (page x 54 + gap). The level indents left 50 hanging 50 and
  // a 5pt marker; the 36pt default stops are measured from the margin.
  const gaps = squareShape(11, 100, 50, 60) + squareShape(12, 250, 50, 60) + squareShape(13, 450, 54, 60);
  const numbered = { marker: '1', ind: '<w:ind w:left="1000" w:hanging="1000"/>' };

  it('takes the first gap that holds the word after the body that gap resolves', () => {
    // A 75pt word: the first gap leaves 50pt after its 50pt hanging stop and
    // the second 70pt after the 180pt stop; the third resolves the 324pt stop
    // and leaves 126pt, so the line stays at the paragraph top there.
    expect(anchoredParagraphLines(gaps, 'W'.repeat(15), numbered)[0]).toMatchObject({
      topPt: 0, marker: 54 + 300, body: 54 + 324, text: 'W'.repeat(15),
    });
  });

  it('admits a later gap by its own body, not the first gap\'s indent requirement', () => {
    // No indents and a custom stop at 200pt: the first gap [0,210] puts the
    // body at 200pt and leaves 10pt for the 40pt word. The second gap
    // [220,300] puts the 5pt marker at 220pt and its body at the 252pt
    // default stop, leaving 48pt: the line stays at the paragraph top there
    // although that gap (80pt) is narrower than the first gap's 200pt body.
    const floats = squareShape(31, 210, 10, 60) + squareShape(32, 300, 204, 60);
    const custom = { marker: '1', ind: '<w:tabs><w:tab w:val="left" w:pos="4000"/></w:tabs><w:ind w:left="0" w:hanging="0"/>' };
    expect(anchoredParagraphLines(floats, 'W'.repeat(8), custom)[0]).toMatchObject({
      topPt: 0, marker: 54 + 220, body: 54 + 252, text: 'W'.repeat(8),
    });
  });

  it('admits the first gap by its own body before any rejection', () => {
    // A 200pt custom stop with no indents, gaps [0,100] and [150,300] (the
    // band ends at 360pt): the first gap is narrower than the 200pt body the
    // indents give, the second resolves its body at the same 200pt stop and
    // leaves 100pt for the 40pt word, so the line stays at the paragraph top.
    const floats = squareShape(33, 100, 50, 60) + squareShape(34, 300, 204, 60);
    const custom = { marker: '1', ind: '<w:tabs><w:tab w:val="left" w:pos="4000"/></w:tabs><w:ind w:left="0" w:hanging="0"/>' };
    expect(anchoredParagraphLines(floats, 'W'.repeat(8), custom)[0]).toMatchObject({
      topPt: 0, marker: 54 + 150, body: 54 + 200, text: 'W'.repeat(8),
    });
  });

  it('sets a numbered word wider than every window below the floats', () => {
    // 600pt: no window (nor the 504pt band) holds it; it starts below the
    // floats after its body and is broken there.
    const lines = anchoredParagraphLines(gaps, 'W'.repeat(120), numbered);
    expect(lines[0]).toMatchObject({ topPt: 60, marker: 54, body: 54 + 50 });
  });

  it('lets break-time hanging punctuation admit a numbered first line\'s window', () => {
    // 10pt per character, a float over [0,477] and a custom stop at 489pt:
    // the marker sits at 477pt, its tab reaches 489pt and leaves 15pt, which
    // holds 漢 with the closing 。 hanging past the band edge. No window floor
    // may require the 20pt the pair takes without hanging.
    const custom = { marker: '1', ind: '<w:tabs><w:tab w:val="left" w:pos="9780"/></w:tabs><w:ind w:left="0" w:hanging="0"/>' };
    expect(anchoredParagraphLines(squareShape(35, 0, 477, 60), '\u6f22\u3002', custom, { advancePt: 10 })[0]).toMatchObject({
      topPt: 0, marker: 54 + 477, body: 54 + 489, text: '\u6f22\u3002',
    });
  });

  it('applies break-time hanging punctuation to the retry after a rejected gap', () => {
    // A float over [100,150] anchored after the text (so the text is the
    // head): the first gap puts the body at the 489pt stop past its end and
    // rejects 漢。 (20pt); the later gap, which ends at the band edge, leaves
    // 15pt after its body and hangs the 。.
    const custom = { marker: '1', ind: '<w:tabs><w:tab w:val="left" w:pos="9780"/></w:tabs><w:ind w:left="0" w:hanging="0"/>' };
    expect(anchoredParagraphLines(squareShape(36, 100, 50, 60), '\u6f22\u3002', custom,
      { advancePt: 10, anchorsLast: true })[0]).toMatchObject({
      topPt: 0, marker: 54 + 150, body: 54 + 489, text: '\u6f22\u3002',
    });
  });

  it('hangs a mark that ends a head joined across formatted runs on retry', () => {
    // 'A' and a bold '.' form one head unit (20pt): the first gap leaves 15pt
    // and rejects it against the float; the later gap ends at the band edge,
    // leaves 15pt after its body and hangs the period.
    const custom = { marker: '1', ind: '<w:tabs><w:tab w:val="left" w:pos="1700"/><w:tab w:val="left" w:pos="9780"/></w:tabs><w:ind w:left="0" w:hanging="0"/>' };
    expect(anchoredParagraphLines(squareShape(39, 100, 50, 60), ['A', '.'], custom,
      { advancePt: 10 })[0]).toMatchObject({ topPt: 0, text: 'A.' });
  });

  it('hangs a lone mark after its float anchor on retry, as break time does', () => {
    // '.' (20pt) after the anchor: the first gap rejects it; the later gap
    // leaves 15pt after its body and hangs it.
    const custom = { marker: '1', ind: '<w:tabs><w:tab w:val="left" w:pos="1700"/><w:tab w:val="left" w:pos="9780"/></w:tabs><w:ind w:left="0" w:hanging="0"/>' };
    expect(anchoredParagraphLines(squareShape(40, 100, 50, 60), '.', custom,
      { advancePt: 20 })[0]).toMatchObject({ topPt: 0, text: '.' });
  });

  it('does not hang punctuation in a retried gap that ends at an exclusion', () => {
    // A stop at 474pt and a second float over [489,504]: the later gap ends
    // at that float, leaves 15pt and owns its full ink, so 漢。 (20pt) goes
    // below the floats.
    const custom = { marker: '1', ind: '<w:tabs><w:tab w:val="left" w:pos="9480"/></w:tabs><w:ind w:left="0" w:hanging="0"/>' };
    expect(anchoredParagraphLines(squareShape(37, 100, 50, 60) + squareShape(38, 489, 15, 60), '\u6f22\u3002', custom,
      { advancePt: 10, anchorsLast: true })[0]).toMatchObject({ topPt: 60, text: '\u6f22\u3002' });
  });

  it('goes below the floats when no gap holds the word after its body', () => {
    // 130pt: wider than 126pt left in the third gap.
    expect(anchoredParagraphLines(gaps, 'W'.repeat(26), numbered)[0]).toMatchObject({
      topPt: 60, marker: 54, body: 54 + 50,
    });
  });
});

describe('over-wide first word after a float anchor (#2057, word-overlong-head-after-anchor)', () => {
  it('breaks the word from the float bottom without an empty line before it', () => {
    // A 10pt float leaves a 100pt window; a 600pt word is wider than the
    // 504pt column, so it moves below the float and is broken there.
    const lines = anchoredParagraphLines(squareShape(21, 100, 404, 10), 'W'.repeat(120), null);
    expect(lines[0]).toMatchObject({ topPt: 10, text: 'W'.repeat(100) });
    expect(lines[1]!.text).toBe('W'.repeat(20));
  });

  it.each([undefined, 15])('keeps an unspaced word across differently formatted runs off an anchor-only line (mode %s)', (mode) => {
    // The 800pt word joins two runs (plain, bold), so the joined-word
    // preflight judges it before the emergency splitter: the anchor alone
    // does not make the line nonempty there either.
    const lines = anchoredParagraphLines(squareShape(22, 100, 404, 10), ['W'.repeat(80), 'W'.repeat(80)], null, { compatibilityMode: mode });
    expect(lines[0]).toMatchObject({ topPt: 10, text: 'W'.repeat(100) });
    expect(lines[1]!.text).toBe('W'.repeat(60));
  });

  it('keeps the established empty line before a joined unit holding SEA dictionary text (unmeasured)', () => {
    // A Thai word and a bold "!" in separate runs (110pt per character, 550pt
    // in all) after the anchor of a float past the band edge: the joined-run
    // preflight still ends the anchor-only line first for SEA text.
    const thai = '\u0e20\u0e32\u0e29\u0e32';
    const lines = anchoredParagraphLines(squareShape(24, 505, 1, 10), [thai, '!'], null, { advancePt: 110 });
    expect(lines.map(({ topPt, text }) => ({ topPt, text }))).toEqual([
      { topPt: 0, text: '' }, { topPt: 10, text: `${thai}!` },
    ]);
  });

  it('keeps the established empty line before an over-wide SEA dictionary word (unmeasured)', () => {
    // WORD_OVERLONG_HEAD_AFTER_ANCHOR is measured for Latin words only. A
    // Thai dictionary word wider than the band (4 x 150pt) after the anchor
    // keeps the previous behaviour: the anchor-only line is ended first.
    const thai = '\u0e20\u0e32\u0e29\u0e32';
    const lines = anchoredParagraphLines(squareShape(23, 100, 404, 10), thai, null, { advancePt: 150 });
    expect(lines.map(({ topPt, text }) => ({ topPt, text }))).toEqual([
      { topPt: 10, text: '' }, { topPt: 20, text: thai.slice(0, 3) },
    ]);
  });
});

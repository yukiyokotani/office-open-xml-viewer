/// <reference types="node" />

import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from '../wasm/docx_parser.js';
import { storeZip } from '../conformance/generate.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import { createLayoutServices } from '../layout-runtime.js';
import { layoutDocument } from '../document-layout.js';

// WORD_TABLE_NONLEADING_INDENT_SCOPE: ECMA-376 §17.4.50 ignores tblInd for a
// center/right/end table in the structures outside the origin observation.

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const encoder = new TextEncoder();

type Story = 'body' | 'nested' | 'header' | 'footer' | 'textbox' | 'eaVertTextbox' | 'footnote';

const MARGINS = '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>';
const para = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

function tableXml(options: Readonly<{
  label: string; jc: string; indent: number | null; margins?: boolean | string; bidi?: boolean;
  tblW?: string; layout?: string; tcW?: string; spacing?: boolean; trPr?: string;
}>): string {
  const tcW = options.tcW ?? '<w:tcW w:w="2000" w:type="dxa"/>';
  const cell = (text: string) => `<w:tc><w:tcPr>${tcW}</w:tcPr>${para(text)}</w:tc>`;
  return `<w:tbl><w:tblPr>${options.bidi ? '<w:bidiVisual/>' : ''}${options.tblW ?? '<w:tblW w:w="4000" w:type="dxa"/>'}`
    + `<w:jc w:val="${options.jc}"/>${options.spacing ? '<w:tblCellSpacing w:w="60" w:type="dxa"/>' : ''}`
    + `${options.indent === null ? '' : `<w:tblInd w:w="${options.indent}" w:type="dxa"/>`}`
    + `${options.layout ?? '<w:tblLayout w:type="fixed"/>'}`
    + `${options.margins === false ? '' : typeof options.margins === 'string' ? options.margins : MARGINS}</w:tblPr>`
    + `<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>`
    + `<w:tr>${options.trPr ? `<w:trPr>${options.trPr}</w:trPr>` : ''}${cell(options.label)}${cell('b')}</w:tr></w:tbl>`;
}

function textBox(content: string, vert = 'horz'): string {
  const emu = (pt: number) => Math.round(pt * 12700);
  return '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
    + `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${emu(90)}</wp:posOffset></wp:positionH>`
    + `<wp:positionV relativeFrom="page"><wp:posOffset>${emu(140)}</wp:posOffset></wp:positionV>`
    + `<wp:extent cx="${emu(400)}" cy="${emu(200)}"/><wp:wrapNone/><wp:docPr id="1" name="Box"/>`
    + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>'
    + `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(400)}" cy="${emu(200)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>`
    + `<wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx><wps:bodyPr vert="${vert}" wrap="square"><a:noAutofit/></wps:bodyPr>`
    + '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
}

/** `mode: null` omits compatSetting compatibilityMode. */
function tableGeometry(story: Story, table: string, label: string, mode: number | null): { xPt: number; widthPt: number } {
  const outer = `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblLayout w:type="fixed"/>${MARGINS}</w:tblPr>`
    + `<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>${table}${para('end')}</w:tc></w:tr></w:tbl>`;
  const body = story === 'body' ? table + para('after')
    : story === 'nested' ? outer + para('after')
      : story === 'textbox' ? `<w:p>${textBox(table + para('end'))}</w:p>`
        : story === 'eaVertTextbox' ? `<w:p>${textBox(table + para('end'), 'eaVert')}</w:p>`
        : story === 'footnote' ? '<w:p><w:r><w:t>ref</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
          : para('body');
  const ns = `xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" `
    + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';
  const sect = '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/>'
    + '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="360" w:footer="360"/></w:sectPr>';
  const files = new Map<string, string>([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'],
    ['_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="d" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${[
      ['s', 'settings', 'settings.xml'], ['rIdH', 'header', 'header1.xml'], ['rIdF', 'footer', 'footer1.xml'], ['n', 'footnotes', 'footnotes.xml'],
    ].map(([id, type, target]) => `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`).join('')}</Relationships>`],
    ['word/document.xml', `<w:document ${ns}><w:body>${body}${sect}</w:body></w:document>`],
    ['word/settings.xml', `<w:settings xmlns:w="${W}"><w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>${mode === null ? ''
      : `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/></w:compat>`}</w:settings>`],
    ['word/header1.xml', `<w:hdr ${ns}>${story === 'header' ? table : ''}${para('header')}</w:hdr>`],
    ['word/footer1.xml', `<w:ftr ${ns}>${story === 'footer' ? table : ''}${para('footer')}</w:ftr>`],
    ['word/footnotes.xml', `<w:footnotes ${ns}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>`
      + '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
      + `<w:footnote w:id="1">${para('note')}${story === 'footnote' ? table : ''}${para('end')}</w:footnote></w:footnotes>`],
  ]);
  const archive = new DocxArchive(storeZip(new Map([...files].map(([name, xml]) => [name, encoder.encode(xml)]))));
  try {
    const model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
    const context = {
      font: '10px serif', letterSpacing: '0px', fontKerning: 'none',
      measureText: (text: string) => ({
        width: [...text].length * 6,
        fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
        actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
      }),
    } as unknown as CanvasRenderingContext2D;
    const page = layoutDocument(model, createLayoutServices(model, { measureContext: context }), { currentDateMs: 0 }).pages[0];
    // The innermost retained table whose content carries the label.
    const seen = new Set<object>();
    const found: { kind: string; flowBounds: { xPt: number; widthPt: number } }[] = [];
    const walk = (value: unknown) => {
      if (value === null || typeof value !== 'object' || seen.has(value)) return;
      seen.add(value);
      const node = value as { kind?: unknown; flowBounds?: { xPt: number; widthPt: number } };
      if (node.kind === 'table' && node.flowBounds && JSON.stringify(value).includes(label)) {
        found.push(node as { kind: string; flowBounds: { xPt: number; widthPt: number } });
      }
      for (const child of Object.values(value)) walk(child);
    };
    walk(page?.layers);
    const table = found.at(-1);
    if (!table) throw new Error(`no ${story} table ${label}`);
    return { xPt: table.flowBounds.xPt, widthPt: table.flowBounds.widthPt };
  } finally {
    archive.free();
  }
}

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('../wasm/docx_parser_bg.wasm', import.meta.url)) });
});

/** Shift and width change of a table with tblInd +720 and -360 against its
 * twin without tblInd: [dx(+720), dx(-360), dw(+720), dw(-360)]. */
function deltas(story: Story, mode: number | null, options: Omit<Parameters<typeof tableXml>[0], 'label' | 'indent'>): number[] {
  const at = (indent: number | null) => {
    const label = `Q${indent ?? 'n'}`;
    return tableGeometry(story, tableXml({ ...options, label, indent }), label, mode);
  };
  const base = at(null);
  const plus = at(720);
  const minus = at(-360);
  return [plus.xPt - base.xPt, minus.xPt - base.xPt, plus.widthPt - base.widthPt, minus.widthPt - base.widthPt];
}
/** Shift of a table with tblInd +720 and -360 against its twin without tblInd. */
function shifts(story: Story, mode: number | null, options: Omit<Parameters<typeof tableXml>[0], 'label' | 'indent'>): number[] {
  return deltas(story, mode, options).slice(0, 2);
}

describe('WORD_TABLE_NONLEADING_INDENT_SCOPE', () => {
  it.each([
    ['percentage cell widths', { tcW: '<w:tcW w:w="2500" w:type="pct"/>' }],
    ['nil cell widths', { tcW: '<w:tcW w:w="0" w:type="nil"/>' }],
    ['a percentage table width', { tblW: '<w:tblW w:w="2000" w:type="pct"/>' }],
    ['an AutoFit auto-width table with dxa cells', { tblW: '<w:tblW w:w="0" w:type="auto"/>', layout: '' }],
    ['cell spacing', { spacing: true }],
    ['omitted margins', { margins: false }],
    ['a right-only margin', { margins: '<w:tblCellMar><w:right w:w="108" w:type="dxa"/></w:tblCellMar>' }],
  ] as const)('a body table with %s ignores a nonleading tblInd in every measured mode', (_name, structure) => {
    for (const mode of [null, 11, 14, 15]) {
      for (const jc of ['center', 'right']) {
        expect(shifts('body', mode, { ...structure, jc }), `${jc} ${mode}`).toEqual([0, 0]);
      }
    }
    // A leading table keeps the indent, and bidiVisual mirrors a right table.
    expect(shifts('body', 15, { ...structure, jc: 'left' })).toEqual([36, -18]);
    expect(shifts('body', 15, { ...structure, jc: 'right', bidi: true })).toEqual([0, 0]);
  });
  it.each(['nested', 'header', 'footer', 'textbox', 'footnote'] as const)('a %s table ignores a nonleading tblInd', (story) => {
    for (const margins of [true, false]) {
      for (const jc of ['center', 'end']) {
        expect(shifts(story, 14, { jc, margins }), `${jc} ${margins}`).toEqual([0, 0]);
        expect(shifts(story, null, { jc, margins, bidi: true }), `${jc} ${margins} bidi`).toEqual([0, 0]);
      }
    }
    expect(shifts(story, 15, { jc: 'left' })[0]).toBeCloseTo(36, 8);
  });
  it('an unmeasured compatibility mode keeps the previous indent contract', () => {
    expect(shifts('body', 13, { jc: 'right', margins: false })).toEqual([36, -18]);
    expect(shifts('header', 13, { jc: 'right' })).toEqual([36, -18]);
    // An AutoFit table without any preferred width waits for its width fix.
    expect(shifts('body', 15, { jc: 'right', tblW: '<w:tblW w:w="0" w:type="auto"/>', layout: '',
      tcW: '<w:tcW w:w="0" w:type="auto"/>' })).toEqual([36, -18]);
    // Unmeasured margin spellings stay outside the scope as well.
    expect(shifts('header', 15, { jc: 'right',
      margins: '<w:tblCellMar><w:left w:w="0" w:type="nil"/></w:tblCellMar>' })).toEqual([36, -18]);
  });
  it('decides scope boundaries from effective facts, comparing position and size', () => {
    // §17.4.87: a percent measurement is a percentage whatever @type says,
    // and an omitted @type is dxa.
    expect(deltas('header', 15, { jc: 'right',
      margins: '<w:tblCellMar><w:left w:type="dxa" w:w="20%"/></w:tblCellMar>' })).toEqual([36, -18, 0, 0]);
    expect(deltas('header', 15, { jc: 'right',
      margins: '<w:tblCellMar><w:left w:w="108"/><w:right w:w="108"/></w:tblCellMar>' })).toEqual([0, 0, 0, 0]);
    // A quarter-turned (eaVert) text-box story in a horizontal section is unmeasured.
    expect(deltas('eaVertTextbox', 15, { jc: 'right' })).toEqual(deltas('eaVertTextbox', 13, { jc: 'right' }));
    expect(deltas('eaVertTextbox', 15, { jc: 'right' }).slice(0, 2)).toEqual([36, -18]);
    // tcW pct 0 is no preference: still a content-sized AutoFit table.
    const auto = { jc: 'right', tblW: '<w:tblW w:w="0" w:type="auto"/>', layout: '' };
    expect(deltas('body', 15, { ...auto, tcW: '<w:tcW w:w="0" w:type="pct"/>' }))
      .toEqual(deltas('body', 13, { ...auto, tcW: '<w:tcW w:w="0" w:type="pct"/>' }));
    expect(shifts('body', 15, { ...auto, tcW: '<w:tcW w:w="0" w:type="pct"/>' })).toEqual([36, -18]);
    // Authored skips the parser clears against the two-column grid stay out
    // (omitted margins keep the origin observation out of the comparison).
    for (const trPr of ['<w:gridAfter w:val="1"/>', '<w:gridBefore w:val="3"/>']) {
      const options = { jc: 'right', margins: false, trPr };
      expect(deltas('body', 15, options), trPr).toEqual(deltas('body', 13, options));
      expect(shifts('body', 15, options), trPr).toEqual([36, -18]);
    }
  });
});

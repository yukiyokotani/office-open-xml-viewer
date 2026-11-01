/// <reference types="node" />

import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from '../wasm/docx_parser.js';
import { storeZip } from '../conformance/generate.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import { createLayoutServices } from '../layout-runtime.js';
import { layoutDocument } from '../document-layout.js';
import type { TableLayout } from './types.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const encoder = new TextEncoder();

const LOGICAL_TABLE_STYLE = '<w:style w:type="table" w:styleId="LogicalTable"><w:name w:val="Logical table"/></w:style>';

/** `mode: null` omits compatSetting compatibilityMode. */
function layoutTable(tableXml: string, mode: number | null = 14, tableStyles = LOGICAL_TABLE_STYLE): TableLayout {
  const parts = new Map([
    ['[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
      <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
    </Types>`],
    ['_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>`],
    ['word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
      <Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
    </Relationships>`],
    ['word/document.xml', `<w:document xmlns:w="${W}"><w:body>${tableXml}
      <w:sectPr><w:pgSz w:w="12240" w:h="14400"/>
        <w:pgMar w:left="1440" w:right="1440" w:top="1440" w:bottom="1440"/>
      </w:sectPr></w:body></w:document>`],
    ['word/styles.xml', `<w:styles xmlns:w="${W}">${tableStyles}</w:styles>`],
    ['word/settings.xml', `<w:settings xmlns:w="${W}">${mode === null ? '' : `<w:compat>
      <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/>
    </w:compat>`}</w:settings>`],
  ]);
  const archive = new DocxArchive(storeZip(new Map(
    [...parts].map(([name, xml]) => [name, encoder.encode(xml)]),
  )));
  try {
    const model = normalizeInternalDocumentModel(
      JSON.parse(new TextDecoder().decode(archive.parse())),
    ).document;
    const context = {
      font: '10px serif', letterSpacing: '0px', fontKerning: 'none',
      measureText: (text: string) => ({
        width: [...text].length * 6,
        fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
        actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
      }),
    } as unknown as CanvasRenderingContext2D;
    const result = layoutDocument(model, createLayoutServices(model, { measureContext: context }), {
      currentDateMs: 0,
    });
    const table = result.pages[0]?.layers.body.find((block) => block.kind === 'table');
    if (table?.kind !== 'table') throw new Error('expected retained parsed table');
    return table;
  } finally {
    archive.free();
  }
}

const margins = (left = 360) => `<w:tblCellMar><w:left w:type="dxa" w:w="${left}"/><w:right w:type="dxa" w:w="180"/></w:tblCellMar>`;
const cell = (width: number, left?: number) => `<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="${width}"/>${left === undefined ? '' : `<w:tcMar><w:left w:type="dxa" w:w="${left}"/></w:tcMar>`}</w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>`;
function table(rows: string, options = ''): string {
  // ECMA-376 CT_TblPrBase orders style/width/placement before layout/margins.
  const propertyOrder = ['tblStyle', 'bidiVisual', 'tblW', 'jc', 'tblCellSpacing', 'tblInd'];
  const orderedOptions = [...options.matchAll(/<w:(\w+)\b[^>]*\/>/g)]
    .sort((a, b) => propertyOrder.indexOf(a[1] ?? '') - propertyOrder.indexOf(b[1] ?? ''))
    .map(([xml]) => xml).join('');
  return `<w:tbl><w:tblPr>${orderedOptions}<w:tblLayout w:type="fixed"/>${margins()}</w:tblPr><w:tblGrid><w:gridCol w:w="2160"/><w:gridCol w:w="3600"/></w:tblGrid>${rows}</w:tbl>`;
}
const indent = '<w:tblInd w:type="dxa" w:w="108"/>';
const row = `<w:tr>${cell(2160)}${cell(3600)}</w:tr>`;
beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('../wasm/docx_parser_bg.wasm', import.meta.url)) });
});

describe('Word table origin and occupied grid', () => {
  it.each([11, 12, 14, 15])('mode %s distinguishes missing indentation and authored leading indentation', (mode) => {
    const missing = layoutTable(table(row), mode);
    const authored = layoutTable(table(row, indent), mode);
    expect(missing.flowBounds.xPt).toBe(72);
    expect(authored.flowBounds.xPt).toBeCloseTo(mode === 15 ? 77.4 : 59.4, 8);
    expect(layoutTable(table(row, indent.replace('108', '0')), mode).flowBounds.xPt)
      .toBeCloseTo(mode === 15 ? 72 : 54, 8);
    const rtl = layoutTable(table(row, indent + '<w:bidiVisual/>'), mode);
    expect(rtl.flowBounds.xPt).toBeCloseTo(mode === 15 ? 246.6 : 264.6, 8);
  });
  it('an omitted compatibility mode hangs the first cell margin like mode 14', () => {
    for (const options of [indent, indent + '<w:bidiVisual/>', '<w:jc w:val="right"/>']) {
      expect(layoutTable(table(row, options), null).flowBounds.xPt)
        .toBeCloseTo(layoutTable(table(row, options), 14).flowBounds.xPt, 8);
    }
    expect(layoutTable(table(row, indent), null).flowBounds.xPt).toBeCloseTo(59.4, 8);
  });
  it('a table style tblInd is authored indentation, and unspecified horizontal cell margins are zero', () => {
    // WORD_TABLE_ORIGIN_COMPATIBILITY / WORD_TABLE_UNSPECIFIED_HORIZONTAL_CELL_MARGIN.
    const normalTable = (tblPr: string) => `<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr>${tblPr}</w:tblPr></w:style>`;
    const styleMargins = '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>';
    const indented = '<w:style w:type="table" w:styleId="Indented"><w:basedOn w:val="TableNormal"/><w:tblPr><w:tblInd w:w="500" w:type="dxa"/></w:tblPr></w:style>';
    const styled = normalTable('<w:tblInd w:w="0" w:type="dxa"/>' + styleMargins) + indented;
    const plain = (tblPr = '', tcMar = '') => `<w:tbl><w:tblPr>${tblPr}<w:tblW w:type="dxa" w:w="5000"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="5000"/>${tcMar}</w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const textStart = (layout: TableLayout) => layout.rows[0]?.cells[0]?.contentBounds.xPt;
    const tcMar = '<w:tcMar><w:left w:w="400" w:type="dxa"/></w:tcMar>';
    for (const mode of [null, 11, 12, 14]) {
      // The default style's tblInd 0 hangs the style or tcMar margin outside the band.
      expect(layoutTable(plain(), mode, styled).flowBounds.xPt).toBeCloseTo(66.6, 8);
      expect(textStart(layoutTable(plain('', tcMar), mode, styled))).toBeCloseTo(72, 8);
      // A basedOn style's tblInd 500 is the table indentation.
      expect(textStart(layoutTable(plain('<w:tblStyle w:val="Indented"/>'), mode, styled))).toBeCloseTo(97, 8);
      // Without tblInd anywhere in the hierarchy the same margin does not hang.
      expect(layoutTable(plain('', tcMar), mode, normalTable(styleMargins)).flowBounds.xPt).toBeCloseTo(72, 8);
    }
    expect(textStart(layoutTable(plain('<w:tblStyle w:val="Indented"/>'), 15, styled))).toBeCloseTo(102.4, 8);
    // No style margins at all: the text starts at the table edge on both sides.
    for (const styles of ['', normalTable('<w:tblInd w:w="0" w:type="dxa"/>')]) {
      const bare = layoutTable(plain(), null, styles);
      expect(textStart(bare)).toBeCloseTo(72, 8);
      expect(bare.rows[0]?.cells[0]?.contentBounds.widthPt).toBeCloseTo(250, 8);
    }
  });
  it('a style-supplied center or right jc ignores the style tblInd (§17.4.50), and a direct jc wins', () => {
    const styles = '<w:style w:type="table" w:styleId="Centered"><w:name w:val="Centered"/><w:tblPr><w:jc w:val="center"/><w:tblInd w:w="720" w:type="dxa"/></w:tblPr></w:style>'
      + '<w:style w:type="table" w:styleId="Righted"><w:name w:val="Righted"/><w:basedOn w:val="Centered"/><w:tblPr><w:jc w:val="right"/></w:tblPr></w:style>';
    const tbl = (tblPr: string, marginsXml = '') => `<w:tbl><w:tblPr>${tblPr}<w:tblW w:type="dxa" w:w="5000"/>${marginsXml}</w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="5000"/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    // 468pt band at x=72, 250pt table: centred at 181, right-aligned at 290.
    // The measured origin class (explicit legacy margins, measured modes);
    // unmeasured classes keep their established nonleading-indent contract.
    for (const mode of [null, 11, 12, 14, 15]) {
      for (const marginsXml of [margins()]) {
        expect(layoutTable(tbl('<w:tblStyle w:val="Centered"/>', marginsXml), mode, styles).flowBounds.xPt).toBeCloseTo(181, 8);
        const right = layoutTable(tbl('<w:tblStyle w:val="Righted"/>', marginsXml), mode, styles).flowBounds.xPt;
        expect(right).toBeCloseTo(290 + (mode !== 15 ? 18 : 0), 8);
        // A direct leading jc overrides the style jc, so the style indent applies.
        expect(layoutTable(tbl('<w:tblStyle w:val="Centered"/><w:jc w:val="left"/>', marginsXml), mode, styles).flowBounds.xPt)
          .toBeCloseTo(72 + 36 - (mode !== 15 ? 18 : 0), 8);
      }
    }
  });
  it('an inherited style tblInd never shifts a non-leading table in an unmeasured class', () => {
    // No left margin anywhere and a pct cell width: both outside the origin observation.
    const styles = '<w:style w:type="table" w:styleId="Centered"><w:name w:val="Centered"/><w:tblPr><w:jc w:val="center"/><w:tblInd w:w="720" w:type="dxa"/></w:tblPr></w:style>'
      + '<w:style w:type="table" w:styleId="Plain"><w:name w:val="Plain"/><w:tblPr><w:jc w:val="center"/></w:tblPr></w:style>';
    for (const tcW of ['<w:tcW w:type="dxa" w:w="5000"/>', '<w:tcW w:type="pct" w:w="5000"/>']) {
      const tbl = (tblPr: string) => `<w:tbl><w:tblPr>${tblPr}<w:tblW w:type="dxa" w:w="5000"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr>${tcW}</w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
      for (const mode of [null, 13, 14]) {
        const unindented = layoutTable(tbl('<w:tblStyle w:val="Plain"/>'), mode, styles).flowBounds.xPt;
        expect(layoutTable(tbl('<w:tblStyle w:val="Centered"/>'), mode, styles).flowBounds.xPt).toBeCloseTo(unindented, 8);
        // A direct leading jc keeps the inherited indent.
        expect(layoutTable(tbl('<w:tblStyle w:val="Centered"/><w:jc w:val="left"/>'), mode, styles).flowBounds.xPt).toBeCloseTo(72 + 36, 8);
      }
    }
  });
  it('an inherited style tblInd follows each row\'s resulting jc (trPr and tblPrEx overrides)', () => {
    const styles = '<w:style w:type="table" w:styleId="Indented"><w:name w:val="Indented"/><w:tblPr><w:tblInd w:w="720" w:type="dxa"/></w:tblPr></w:style>'
      + '<w:style w:type="table" w:styleId="CenteredIndented"><w:name w:val="CenteredIndented"/><w:basedOn w:val="Indented"/><w:tblPr><w:jc w:val="center"/></w:tblPr></w:style>';
    const rowJc = (how: 'trPr' | 'tblPrEx', jc: string) => how === 'trPr'
      ? `<w:tr><w:trPr><w:jc w:val="${jc}"/></w:trPr>` : `<w:tr><w:tblPrEx><w:jc w:val="${jc}"/></w:tblPrEx>`;
    const tbl = (style: string, tr: string, marginsXml: string) => `<w:tbl><w:tblPr><w:tblStyle w:val="${style}"/><w:tblW w:type="dxa" w:w="5000"/>${marginsXml}</w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>${tr}<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="5000"/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const rowX = (layout: TableLayout) => layout.rows[0]?.cells[0]?.flowBounds.xPt;
    // 468pt band at x=72, 250pt table: a centred row sits at 181.
    for (const how of ['trPr', 'tblPrEx'] as const) {
      for (const mode of [null, 13, 14, 15]) {
        // Leading table, centred row: the inherited indent is ignored (no horizontal margins: unmeasured).
        expect(rowX(layoutTable(tbl('Indented', rowJc(how, 'center'), ''), mode, styles))).toBeCloseTo(181, 8);
        // Centred table, leading row: the inherited indent applies to that row.
        expect(rowX(layoutTable(tbl('CenteredIndented', rowJc(how, 'left'), ''), mode, styles))).toBeCloseTo(108, 8);
      }
      for (const mode of [null, 14, 15]) {
        // Measured class (explicit legacy margins): the leading row keeps the indent and its margin hang.
        expect(rowX(layoutTable(tbl('CenteredIndented', rowJc(how, 'left'), margins()), mode, styles)))
          .toBeCloseTo(mode === 15 ? 108 : 90, 8);
        expect(rowX(layoutTable(tbl('Indented', rowJc(how, 'center'), margins()), mode, styles))).toBeCloseTo(181, 8);
      }
    }
  });
  it.each(['center', 'right'])('%s ignores indentation and uses the first cell margin only at the logical end', (jc) => {
    const rows = `<w:tr>${cell(2160, 540)}${cell(3600)}</w:tr><w:tr>${cell(2160, 120)}${cell(3600)}</w:tr>`;
    const options = indent + `<w:jc w:val="${jc}"/>`;
    const legacy = layoutTable(table(rows, options));
    expect(legacy.rows.map((r) => r.cells[0]?.flowBounds.xPt)).toEqual(jc === 'center' ? [162, 162] : [279, 258]);
    const rtl = layoutTable(table(rows, options + '<w:bidiVisual/>'));
    expect(rtl.rows.map((r) => r.cells[0]?.flowBounds.xPt)).toEqual(jc === 'center' ? [342, 342] : [225, 246]);
    const modern = layoutTable(table(rows, options), 15);
    expect(modern.rows.map((r) => r.cells[0]?.flowBounds.xPt)).toEqual(jc === 'center' ? [162, 162] : [252, 252]);
  });
  it.each([11, 12, 14, 15])('mode %s ignores nonleading indentation in AutoFit width acquisition', (mode) => {
    // WORD_TABLE_ORIGIN_COMPATIBILITY: the mode-11/12/14/15 center/end
    // controls vary omitted/zero/positive/negative indentation with explicit
    // margins (pages 2/3 and RTL pages 5/6: zero, +36pt, −36pt controls).
    // ECMA-376 §17.4.50 ignores it for every nonleading row.
    // A preferred width beyond the band exposes the width-ceiling leak.
    for (const jc of ['center', 'right']) {
      for (const indentTwips of [null, 0, 720, -720]) {
        const xml = table(`<w:tr>${cell(4680)}${cell(4680)}</w:tr>`, `<w:tblW w:type="dxa" w:w="10800"/><w:jc w:val="${jc}"/>${indentTwips === null ? '' : `<w:tblInd w:type="dxa" w:w="${indentTwips}"/>`}`)
          .replace('w:tblLayout w:type="fixed"', 'w:tblLayout w:type="autofit"')
          .replace('<w:gridCol w:w="2160"/><w:gridCol w:w="3600"/>', '<w:gridCol w:w="4680"/><w:gridCol w:w="4680"/>');
        expect(layoutTable(xml, mode).columnWidthsPt).toEqual([234, 234]);
        // §17.4.50 uses the resulting row justification, including tblPrEx.
        const rowAligned = xml.replace(`<w:jc w:val="${jc}"/>`, '<w:jc w:val="left"/>')
          .replace('<w:tr>', `<w:tr><w:tblPrEx><w:jc w:val="${jc}"/></w:tblPrEx>`);
        expect(layoutTable(rowAligned, mode).columnWidthsPt).toEqual([234, 234]);
      }
    }
  });
  it("keeps main's AutoFit ceiling in an unmeasured compatibility mode", () => {
    const xml = table(`<w:tr>${cell(4680)}${cell(4680)}</w:tr>`, '<w:tblW w:type="dxa" w:w="10800"/><w:jc w:val="center"/><w:tblInd w:type="dxa" w:w="-720"/>')
      .replace('w:tblLayout w:type="fixed"', 'w:tblLayout w:type="autofit"')
      .replace('<w:gridCol w:w="2160"/><w:gridCol w:w="3600"/>', '<w:gridCol w:w="4680"/><w:gridCol w:w="4680"/>');
    expect(layoutTable(xml, 13).columnWidthsPt).toEqual([270, 270]);
  });
  it('disabled seeded member-property mixes preserve main independently for split and single', async () => {
    // These synthetic inputs each contain non-dxa cell preferences, outside
    // both Office observations. Main already differs across authored seams;
    // #1664 preserves each form rather than repairing that unrelated defect.
    // The geometry oracle was captured from the recorded revision in this
    // fixture using the same parser/measurement inputs. It was re-recorded for
    // #1881, whose Word controls keep fixed tracks and absolute-width AutoFit
    // grids beyond the container; these cases all carry a 540pt dxa tblW.
    const baseline = JSON.parse(await readFile(new URL('./table-origin-disabled-main.json', import.meta.url), 'utf8')) as { revision: string; cases: { split: number[][]; single: number[][] }[] };
    expect(baseline.cases).toHaveLength(128);
    let seed = 1664;
    const pick = <T,>(values: readonly T[]): T => {
      seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
      return values[(seed >>> 8) % values.length]!;
    };
    const geometry = (layout: TableLayout) => [layout.columnWidthsPt,
      ...layout.rows.flatMap((r) => r.cells.map((c) => [
        c.flowBounds.xPt, c.flowBounds.yPt, c.flowBounds.widthPt, c.flowBounds.heightPt,
      ]))];
    for (let caseIndex = 0; caseIndex < 128; caseIndex++) {
      const mode = pick([11, 12, 14, 15]);
      const layout = pick(['fixed', 'autofit']);
      const jc = pick(['left', 'center', 'right']);
      const indentation = pick([-720, 0, 108, 720]);
      const exception = pick(['', '<w:tblPrEx><w:tblInd w:type="dxa" w:w="720"/></w:tblPrEx>',
        '<w:tblPrEx><w:tblW w:type="auto" w:w="0"/></w:tblPrEx>',
        '<w:tblPrEx><w:tblLayout w:type="fixed"/></w:tblPrEx>']);
      const makeRow = () => {
        const before = pick([0, 1]);
        const after = pick([0, 1]);
        const width = () => cell(2160).replace('w:type="dxa" w:w="2160"',
          `w:type="${pick(['dxa', 'pct', 'auto', 'nil'])}" w:w="2160"`);
        return `<w:tr><w:trPr><w:gridBefore w:val="${before}"/><w:gridAfter w:val="${after}"/></w:trPr>${width()}${width()}</w:tr>`;
      };
      const rows = [makeRow().replace('<w:tr>', `<w:tr>${exception}`), makeRow()];
      const make = (rows: string) => table(rows, `<w:tblStyle w:val="LogicalTable"/><w:tblW w:type="dxa" w:w="10800"/><w:jc w:val="${jc}"/><w:tblInd w:type="dxa" w:w="${indentation}"/>`)
        .replace('w:tblLayout w:type="fixed"', `w:tblLayout w:type="${layout}"`)
        .replace('<w:tblGrid>', '<w:tblGrid><w:gridCol w:w="720"/>');
      expect(rows.join('')).toMatch(/w:type="(?:pct|auto|nil)" w:w="2160"/);
      expect(geometry(layoutTable(rows.map(make).join(''), mode)), `split case ${caseIndex}`)
        .toEqual(baseline.cases[caseIndex]?.split);
      expect(geometry(layoutTable(make(rows.join('')), mode)), `single case ${caseIndex}`)
        .toEqual(baseline.cases[caseIndex]?.single);
    }
  }, 30_000);
  it('enabled member-property mixes preserve split/single and Word-backed nominal geometry', () => {
    // WORD_TABLE_ORIGIN_COMPATIBILITY: exported fixed, explicit-margin,
    // positive-dxa controls establish these mode/alignment/bidi boundaries.
    // ECMA-376 §17.4.50 ignores indentation for center/end. Legacy leading
    // placement hangs the first logical cell's margin; end hangs each row's.
    const rows = [`<w:tr>${cell(2160, 540)}${cell(3600)}</w:tr>`,
      `<w:tr>${cell(2160, 120)}${cell(3600)}</w:tr>`];
    for (const mode of [11, 12, 14, 15]) {
      for (const jc of ['left', 'center', 'right']) {
        for (const rtl of [false, true]) {
          const make = (rows: string) => table(rows, indent
            + `<w:tblStyle w:val="LogicalTable"/><w:jc w:val="${jc}"/>${rtl ? '<w:bidiVisual/>' : ''}`);
          const single = layoutTable(make(rows.join('')), mode);
          const split = layoutTable(rows.map(make).join(''), mode);
          expect(split.columnWidthsPt).toEqual([108, 180]);
          expect(split.columnWidthsPt).toEqual(single.columnWidthsPt);
          expect(split.rows.map((r) => r.cells.map((c) => c.flowBounds)))
            .toEqual(single.rows.map((r) => r.cells.map((c) => c.flowBounds)));
          const firstCellXs = jc === 'center' ? [rtl ? 342 : 162, rtl ? 342 : 162]
            : jc === 'right' ? (mode === 15 ? [252, 252]
              : rtl ? [225, 246] : [279, 258])
            : [rtl ? (mode === 15 ? 426.6 : 453.6) : (mode === 15 ? 77.4 : 50.4),
              rtl ? (mode === 15 ? 426.6 : 453.6) : (mode === 15 ? 77.4 : 50.4)];
          split.rows.forEach((r, i) => expect(r.cells[0]?.flowBounds.xPt).toBeCloseTo(firstCellXs[i]!, 8));
        }
      }
    }
  });
  it.each([false, true])('shares disabled width eligibility across mixed-preference members (reverse=%s)', (reverse) => {
    const dxa = `<w:tr>${cell(4680)}${cell(4680)}</w:tr>`;
    const automatic = dxa.replaceAll('w:type="dxa"', 'w:type="auto"');
    const rows = reverse ? [automatic, dxa] : [dxa, automatic];
    const make = (rows: string) => table(rows,
      '<w:tblStyle w:val="LogicalTable"/><w:tblW w:type="dxa" w:w="10800"/><w:jc w:val="center"/><w:tblInd w:type="dxa" w:w="-720"/>')
      .replace('w:tblLayout w:type="fixed"', 'w:tblLayout w:type="autofit"')
      .replace('<w:gridCol w:w="2160"/><w:gridCol w:w="3600"/>', '<w:gridCol w:w="4680"/><w:gridCol w:w="4680"/>');
    const single = layoutTable(make(rows.join('')));
    const split = layoutTable(rows.map(make).join(''));
    // Disabled whole-group eligibility: independently retain main's 540pt
    // grid and signed center translation in both forms and both row orders.
    for (const layout of [split, single]) {
      expect(layout.columnWidthsPt).toEqual([270, 270]);
      expect(layout.rows.map((r) => r.cells.map((c) => c.flowBounds))).toEqual(
        [72, 82].map((yPt) => [0, 270].map((xPt) => ({ xPt, yPt, widthPt: 270, heightPt: 10 }))),
      );
    }
  });
  it('enabled grid normalization shares first-row indentation even outside the origin observation', () => {
    // WORD_FIXED_UNUSED_LEADING_GRID is enabled by fixed positive-dxa rows;
    // missing legacy margins disable only WORD_TABLE_ORIGIN_COMPATIBILITY.
    // MS-OI29500 §2.1.156 still gives the first logical row's indent authority.
    const skipped = `<w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr>${cell(2160)}${cell(3600)}</w:tr>`;
    const first = skipped.replace('<w:tr>', '<w:tr><w:tblPrEx><w:tblInd w:type="dxa" w:w="720"/></w:tblPrEx>');
    const make = (rows: string) => table(rows, '<w:tblStyle w:val="LogicalTable"/><w:tblInd w:type="dxa" w:w="0"/>')
      .replace(margins(), '').replace('<w:tblGrid>', '<w:tblGrid><w:gridCol w:w="720"/>');
    const split = layoutTable(make(first) + make(skipped));
    const single = layoutTable(make(first + skipped));
    expect(split.columnWidthsPt).toEqual([0, 108, 180]);
    expect(split.columnWidthsPt).toEqual(single.columnWidthsPt);
    expect(split.rows.map((r) => r.cells.map((c) => c.flowBounds)))
      .toEqual(single.rows.map((r) => r.cells.map((c) => c.flowBounds)));
    expect(split.rows.map((r) => r.cells[0]?.flowBounds.xPt)).toEqual([108, 108]);
  });
  it('split centred members share the first logical row\'s authored indent over a style-inherited one', () => {
    // A style-inherited tblInd is filtered for non-leading rows only when it is
    // the selected indent; the logical first row's tblPrEx indent is authored.
    const styles = '<w:style w:type="table" w:styleId="LogicalTable"><w:name w:val="Logical table"/><w:tblPr><w:tblInd w:w="720" w:type="dxa"/></w:tblPr></w:style>';
    const skipped = `<w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr>${cell(2160)}${cell(3600)}</w:tr>`;
    const first = skipped.replace('<w:tr>', '<w:tr><w:tblPrEx><w:tblInd w:type="dxa" w:w="720"/></w:tblPrEx>');
    const make = (rows: string) => table(rows, '<w:tblStyle w:val="LogicalTable"/><w:jc w:val="center"/>')
      .replace(margins(), '').replace('<w:tblGrid>', '<w:tblGrid><w:gridCol w:w="720"/>');
    // WORD_FIXED_UNUSED_LEADING_GRID (modes 11/12/14/15) makes the members one
    // logical table; no horizontal margins keeps the origin observation off.
    for (const mode of [11, 14, 15]) {
      const split = layoutTable(make(first) + make(skipped) + make(skipped), mode, styles);
      const single = layoutTable(make(first + skipped + skipped), mode, styles);
      expect(split.rows.map((r) => r.cells[0]?.flowBounds.xPt))
        .toEqual(single.rows.map((r) => r.cells[0]?.flowBounds.xPt));
      // Centred 288pt occupied grid in the 468pt band at x=72 (162), plus the shared 36pt.
      expect(split.rows.map((r) => r.cells[0]?.flowBounds.xPt)).toEqual([198, 198, 198]);
    }
  });
  it('first-row exceptions anchor every leading row with the first cell override', () => {
    const rows = `<w:tr><w:tblPrEx><w:tblInd w:type="dxa" w:w="720"/>${margins(240)}</w:tblPrEx>${cell(2160, 540)}${cell(3600)}</w:tr>${row}`;
    expect(layoutTable(table(rows, indent)).rows.map((r) => r.cells[0]?.flowBounds.xPt)).toEqual([81, 81]);
  });
  it('preserves the previous origin for unresolved missing-margin and spaced classes', () => {
    const missing = table(row, indent + '<w:jc w:val="right"/>').replace(margins(), '');
    expect(layoutTable(missing).flowBounds.xPt).toBeCloseTo(257.4, 8);
    const logicalOverride = table(`<w:tr>${cell(2160, 540).replace('<w:left ', '<w:start ')}${cell(3600)}</w:tr>`, indent);
    expect(layoutTable(logicalOverride).flowBounds.xPt).toBeCloseTo(77.4, 8);
    const unpreferred = table(row.replaceAll('w:type="dxa"', 'w:type="auto"'), indent);
    expect(layoutTable(unpreferred).flowBounds.xPt).toBeCloseTo(77.4, 8);
    const autoFit = table(row, indent).replace('w:tblLayout w:type="fixed"', 'w:tblLayout w:type="autofit"');
    expect(layoutTable(autoFit).flowBounds.xPt).toBeCloseTo(77.4, 8);
    const cleared = table(`<w:tr><w:tblPrEx><w:tblW w:type="dxa" w:w="0"/></w:tblPrEx>${cell(2160)}${cell(3600)}</w:tr>`, indent + '<w:tblW w:type="dxa" w:w="5760"/>')
      .replace('w:tblLayout w:type="fixed"', 'w:tblLayout w:type="autofit"');
    expect(layoutTable(cleared).flowBounds.xPt).toBeCloseTo(77.4, 8);
    const spaced = table(row, indent + '<w:tblCellSpacing w:type="dxa" w:w="120"/>');
    expect(layoutTable(spaced).flowBounds.xPt).toBeCloseTo(77.4, 8);
  });
  it('preserves member indentation outside the measured logical-table scope', () => {
    const style = '<w:tblStyle w:val="LogicalTable"/>';
    const exception = row.replace('<w:tr>', '<w:tr><w:tblPrEx><w:tblInd w:type="dxa" w:w="720"/></w:tblPrEx>');
    expect(layoutTable(table(exception, style) + table(row, style), 13).rows
      .map((r) => r.cells[0]?.flowBounds.xPt)).toEqual([108, 72]);
    const unpreferred = row.replaceAll('w:type="dxa"', 'w:type="auto"');
    expect(layoutTable(table(exception, style) + table(unpreferred, style), 14).rows
      .map((r) => r.cells[0]?.flowBounds.xPt)).toEqual([108, 72]);
  });
  it('preserves the occupied-grid solver for unmeasured percentage, negative-width and nested classes', () => {
    const skipped = `<w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr>${cell(2160)}${cell(3600)}</w:tr>`;
    const gridTable = table(skipped, '<w:tblW w:type="auto" w:w="0"/>')
      .replace('<w:tblGrid>', '<w:tblGrid><w:gridCol w:w="720"/>');
    const percentage = gridTable.replace('w:type="auto" w:w="0"', 'w:type="pct" w:w="5000"');
    expect(layoutTable(percentage).columnWidthsPt).toEqual(layoutTable(percentage, 13).columnWidthsPt);
    const negative = gridTable.replace('<w:tcW w:type="dxa" w:w="2160"', '<w:tcW w:type="dxa" w:w="-2160"');
    expect(layoutTable(negative).columnWidthsPt).toEqual(layoutTable(negative, 13).columnWidthsPt);
    const outer = table(`<w:tr>${cell(9360).replace('<w:p>', gridTable + '<w:p>')}</w:tr>`)
      .replace('<w:tblGrid><w:gridCol w:w="2160"/><w:gridCol w:w="3600"/>', '<w:tblGrid><w:gridCol w:w="9360"/>');
    const nested = layoutTable(outer).rows[0]?.cells[0]?.blocks.find((block) => block.layout.kind === 'table')?.layout;
    expect(nested?.kind === 'table' ? nested.columnWidthsPt : null).toEqual([36, 108, 180]);
  });
  it.each([14, 15])('mode %s drops a universally skipped leading fixed track, while an occupied preceding row preserves it', (mode) => {
    const skipped = (before: number) => `<w:tr><w:trPr><w:gridBefore w:val="1"/><w:wBefore w:type="dxa" w:w="${before}"/></w:trPr>${cell(2160)}${cell(3600)}</w:tr>`;
    const full = `<w:tr>${cell(720)}${cell(2160)}${cell(3600)}</w:tr>`;
    const gridTable = (rows: string, preferred: boolean) => table(rows, '<w:tblInd w:type="dxa" w:w="0"/>' + (preferred ? '<w:tblW w:type="dxa" w:w="6480"/>' : '<w:tblW w:type="auto" w:w="0"/>'))
      .replace('<w:tblGrid>', '<w:tblGrid><w:gridCol w:w="720"/>');
    for (const before of [0, 360, 720]) {
      const automatic = layoutTable(gridTable(skipped(before), false), mode);
      expect(automatic.columnWidthsPt).toEqual([0, 108, 180]);
      const preferred = layoutTable(gridTable(skipped(before), true), mode);
      expect(preferred.columnWidthsPt).toEqual([0, 121.5, 202.5]);
      const occupied = layoutTable(gridTable(full + skipped(before), true), mode);
      expect(occupied.columnWidthsPt).toEqual([36, 108, 180]);
      expect(occupied.rows[1]?.cells[0]?.flowBounds.xPt).toBeCloseTo((mode === 15 ? 72 : 54) + 36, 8);
    }
  });
  it.each([11, 12, 14, 15])('mode %s lays out split logical tables like a single tbl', (mode) => {
    const skipped = `<w:tr><w:trPr><w:gridBefore w:val="1"/><w:wBefore w:type="dxa" w:w="720"/></w:trPr>${cell(2160)}${cell(3600)}</w:tr>`;
    const full = `<w:tr>${cell(720)}${cell(2160)}${cell(3600)}</w:tr>`;
    const override = full.replace(cell(720), cell(720, 540));
    const exception = full.replace('<w:tr>', '<w:tr><w:tblPrEx><w:tblInd w:type="dxa" w:w="720"/></w:tblPrEx>');
    const unpreferred = skipped.replace('<w:tcW w:type="dxa" w:w="2160"', '<w:tcW w:type="auto" w:w="0"');
    const gridTable = (rows: string, preferred: boolean) => table(rows,
      '<w:tblStyle w:val="LogicalTable"/><w:tblInd w:type="dxa" w:w="0"/>'
      + (preferred ? '<w:tblW w:type="dxa" w:w="6480"/>' : '<w:tblW w:type="auto" w:w="0"/>'))
      .replace('<w:tblGrid>', '<w:tblGrid><w:gridCol w:w="720"/>');
    const geometry = (layout: TableLayout) => ({
      columns: layout.columnWidthsPt,
      bounds: layout.flowBounds,
      rows: layout.rows.map((r) => r.cells.map((c) => c.flowBounds)),
    });
    for (const preferred of [false, true]) {
      // Occupancy on either side of an authored seam must preserve track zero;
      // all-skipped members must still normalize. Different first-cell margins
      // also exercise the logical first-row leading anchor.
      for (const rows of [[full, skipped], [skipped, full], [skipped, skipped], [override, full], [exception, full],
        [skipped.replace('<w:tr>', '<w:tr><w:tblPrEx><w:tblW w:type="dxa" w:w="6480"/></w:tblPrEx>'), skipped],
        [skipped.replace('<w:tr>', '<w:tr><w:tblPrEx><w:tblW w:type="auto" w:w="0"/></w:tblPrEx>'), skipped],
        [skipped, skipped.replace('<w:tr>', '<w:tr><w:tblPrEx><w:tblW w:type="dxa" w:w="6480"/></w:tblPrEx>')],
        ['', skipped.replace('<w:tr>', '<w:tr><w:tblPrEx><w:tblW w:type="dxa" w:w="6480"/></w:tblPrEx>')],
      ]) {
        const single = layoutTable(gridTable(rows.join(''), preferred), mode);
        const split = layoutTable(rows.map((r) => gridTable(r, preferred)).join(''), mode);
        expect(geometry(split)).toEqual(geometry(single));
        if (rows[0] === full && rows[1] === skipped) {
          expect(split.rows[1]?.cells.map((c) => c.flowBounds.xPt))
            .toEqual(mode === 15 ? [108, 216] : [90, 198]);
          const keepNext = '<w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:t>lead</w:t></w:r></w:p>';
          expect(geometry(layoutTable(keepNext + rows.map((r) => gridTable(r, preferred)).join(''), mode)))
            .toEqual(geometry(layoutTable(keepNext + gridTable(rows.join(''), preferred), mode)));
        }
      }
    }
    const expectMainSkippedRows = (layout: TableLayout) => {
      expect(layout.columnWidthsPt).toEqual([36, 108, 180]);
      expect(layout.rows.map((r) => r.cells.map((c) => c.flowBounds))).toEqual(
        [72, 82].map((yPt) => [
          { xPt: 108, yPt, widthPt: 108, heightPt: 10 },
          { xPt: 216, yPt, widthPt: 180, heightPt: 10 },
        ]),
      );
    };
    // Non-dxa preferences disable both observations. Main's skipped-track
    // geometry is the oracle for each authored representation independently.
    expectMainSkippedRows(layoutTable(gridTable(unpreferred, false) + gridTable(skipped, false), mode));
    expectMainSkippedRows(layoutTable(gridTable(unpreferred + skipped, false), mode));
    for (const layout of ['fixed', 'autofit']) {
      const first = skipped.replace('<w:tr>', `<w:tr><w:tblPrEx><w:tblLayout w:type="${layout}"/></w:tblPrEx>`);
      const authored = (rows: string) => gridTable(rows, false)
        .replace('w:tblLayout w:type="fixed"', 'w:tblLayout w:type="autofit"');
      if (layout === 'fixed') {
        // A first-row fixed exception enables the occupied-grid observation.
        expect(geometry(layoutTable(authored(first) + authored(skipped), mode)))
          .toEqual(geometry(layoutTable(authored(first + skipped), mode)));
      } else {
        expectMainSkippedRows(layoutTable(authored(first) + authored(skipped), mode));
        expectMainSkippedRows(layoutTable(authored(first + skipped), mode));
      }
      // Later layout exceptions cannot enable a disabled logical table.
      expectMainSkippedRows(layoutTable(authored(skipped) + authored(first), mode));
      expectMainSkippedRows(layoutTable(authored(skipped + first), mode));
    }
  });
});

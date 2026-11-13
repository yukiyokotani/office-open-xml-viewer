/// <reference types="node" />
import { readFileSync, readdirSync } from 'node:fs';
import { beforeAll, describe, expect, it } from 'vitest';
// @ts-ignore — shared synthetic ZIP fixture is a plain JavaScript module.
import { storedZip } from '../../../tests/fixtures/chart-ex-packages.mjs';
import { initSync, XlsxArchive } from './wasm/xlsx_parser.js';
import { XlsxFontDemand } from './font-source-load.js';
import { OFFICE_FONT_SUBSTITUTIONS } from '@silurus/ooxml-core/internal/font-sources';
import { renderViewport } from './renderer.js';
import { resolveSharedStringRows } from './shared-strings.js';
import type { ParsedWorkbook, Row, Worksheet } from './types.js';
import { SYMBOL_ENCODED, symbolRealm } from './test-support/symbol-realm.js';

// Design §6 invariant: finite XLSX scalar demand is a superset of every scalar
// the renderer measures or paints. A scalar outside the demand belongs to a
// unicode-range subset that never loads, so its glyph would silently fall back
// to a host font. Exercised over targeted synthetic workbooks (number formats,
// dates, rich and inline text, phonetic runs, errors, booleans, shapes) and
// the public demo workbook; every paint path the collector cannot prove
// declines to 'all' (charts, equations, vertical/stacked/RTL text, metafiles,
// and any scalar that is not normalization-inert, e.g. a precomposed é).

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function workbookBytes(styles: string, sheet: string, sharedStrings?: string, drawing?: string, cols = ''): Uint8Array {
  return storedZip([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      + '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
      + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
      + (sharedStrings ? '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' : '')
      + (drawing ? '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' : '')
      + '</Types>'],
    ['_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ['xl/workbook.xml', `<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`],
    ['xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>`
      + `<Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/>`
      + (sharedStrings ? `<Relationship Id="rId3" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/>` : '')
      + '</Relationships>'],
    ['xl/styles.xml', `<styleSheet xmlns="${MAIN}">${styles}</styleSheet>`],
    ...(sharedStrings ? [['xl/sharedStrings.xml', `<sst xmlns="${MAIN}">${sharedStrings}</sst>`]] : []),
    ['xl/worksheets/sheet1.xml', `<worksheet xmlns="${MAIN}" xmlns:r="${REL}">${cols}<sheetData>${sheet}</sheetData>${drawing ? '<drawing r:id="rIdD"/>' : ''}</worksheet>`],
    ...(drawing ? [
      ['xl/worksheets/_rels/sheet1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdD" Type="${REL}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`],
      ['xl/drawings/drawing1.xml', '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        + drawing + '</xdr:wsDr>'],
    ] : []),
  ]) as Uint8Array;
}

const fonts = '<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font>'
  + '<font><b/><sz val="12"/><name val="Cambria"/></font><font><sz val="11"/><name val="Meiryo"/></font></fonts>';
const base = '<fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders>';

const numberFormats = workbookBytes(
  '<numFmts count="9">'
  + '<numFmt numFmtId="164" formatCode="&quot;€&quot;#,##0.00;[Red]\\-&quot;€&quot;#,##0.00"/>'
  + '<numFmt numFmtId="165" formatCode="[$¥-411]#,##0"/>'
  + '<numFmt numFmtId="166" formatCode="dddd, mmmm d, yyyy"/>'
  + '<numFmt numFmtId="167" formatCode="h:mm:ss AM/PM"/>'
  + '<numFmt numFmtId="168" formatCode="[$-411]ggge&quot;年&quot;m&quot;月&quot;d&quot;日&quot;(aaa)"/>'
  + '<numFmt numFmtId="169" formatCode="# ?/?"/>'
  + '<numFmt numFmtId="170" formatCode="0.00E+00"/>'
  + '<numFmt numFmtId="171" formatCode="0;0;0;&quot;» &quot;@&quot; ‰&quot;"/>'
  + '<numFmt numFmtId="172" formatCode="*·0"/>'
  + '</numFmts>' + fonts + base
  + '<cellXfs count="13">'
  + ['0', '164', '165', '166', '167', '168', '169', '170', '171', '172', '14', '22', '10']
    .map((id, index) => `<xf numFmtId="${id}" fontId="${index % 3}" fillId="0" borderId="0" applyNumberFormat="1"/>`).join('')
  + '</cellXfs>',
  '<row r="1">'
    + [[1, '-1234.5'], [2, '98765'], [3, '45291.75'], [4, '0.6234'], [5, '45291'], [6, '3.375'], [7, '0.000123'], [8, '7'], [9, '42'],
      [10, '45291'], [11, '45291.5'], [12, '0.257']]
      .map(([style, value], index) => `<c r="${String.fromCharCode(65 + index)}1" s="${style}"><v>${value}</v></c>`).join('')
    + '</row><row r="2"><c r="A2" t="b"><v>1</v></c><c r="B2" t="e"><v>#DIV/0!</v></c>'
    + '<c r="C2" s="8" t="inlineStr"><is><t>inline Øresund æ ł</t></is></c><c r="D2" t="str"><f>A1&amp;"x"</f><v>formula ð</v></c>'
    + '<c r="E2"><v>123456789012345</v></c><c r="F2"><v>-0.000000123</v></c></row>',
);

const richText = workbookBytes(
  fonts + base + '<cellXfs count="3"><xf numFmtId="0" fontId="0"/><xf numFmtId="0" fontId="1"/><xf numFmtId="0" fontId="2"/></cellXfs>',
  '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s" s="2"><v>1</v></c><c r="C1" t="s" s="1"><v>2</v></c></row>'
    + '<row r="2"><c r="A2" t="inlineStr"><is><r><rPr><b/><rFont val="Roboto"/></rPr><t>Bold Ωμεγα </t></r><r><t>plain — “quoted”</t></r></is></c></row>',
  '<si><t>Straße naive cafe</t></si>'
    + '<si><t>東京都</t><rPh sb="0" eb="3"><t>トウキョウト</t></rPh><phoneticPr fontId="2"/></si>'
    + '<si><r><rPr><i/><rFont val="Lato"/></rPr><t>Łazy-free </t></r><r><rPr><rFont val="Calibri"/></rPr><t>runs ✓</t></r></si>',
);

const shapes = workbookBytes(
  fonts + base + '<cellXfs count="1"><xf numFmtId="0" fontId="0"/></cellXfs>',
  '<row r="1"><c r="A1" t="inlineStr"><is><t>cell</t></is></c></row>',
  undefined,
  '<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
    + '<xdr:to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
    + '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="Text 1"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>'
    + '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="1200000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>'
    + '<xdr:txBody><a:bodyPr wrap="square"/><a:p><a:r><a:rPr lang="en-US" sz="1400" b="1"><a:latin typeface="Nunito"/></a:rPr><a:t>Shape text: wrapping across lines – 100%</a:t></a:r></a:p>'
    + '<a:p><a:r><a:rPr lang="ru-RU" sz="1100"/><a:t>Привет\tмир</a:t></a:r><a:br><a:rPr/></a:br><a:r><a:rPr sz="1100"/><a:t>after break</a:t></a:r></a:p></xdr:txBody>'
    + '</xdr:sp><xdr:clientData/></xdr:twoCellAnchor>',
);

/** Canvas stand-in recording every string measured or painted. */
function recordingContext() {
  let font = '10px serif';
  const texts: string[] = [];
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const noop = () => {};
  const metrics = (text: string) => ({
    width: [...text].length * px() * 0.5,
    fontBoundingBoxAscent: px() * 0.8,
    fontBoundingBoxDescent: px() * 0.2,
    actualBoundingBoxAscent: px() * 0.8,
    actualBoundingBoxDescent: px() * 0.2,
    actualBoundingBoxLeft: 0,
    actualBoundingBoxRight: [...text].length * px() * 0.5,
  }) as TextMetrics;
  const target: Record<string | symbol, unknown> = {
    get font() { return font; },
    set font(value: string) { font = value; },
    canvas: { width: 1600, height: 900 },
    letterSpacing: '0px',
    globalAlpha: 1,
    measureText: (text: string) => { texts.push(text); return metrics(text); },
    fillText: (text: string) => { texts.push(text); },
    strokeText: (text: string) => { texts.push(text); },
    createLinearGradient: () => ({ addColorStop: noop }),
    createRadialGradient: () => ({ addColorStop: noop }),
    createPattern: () => null,
    getTransform: () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0, isIdentity: true }),
    getLineDash: () => [],
    getImageData: (_x: number, _y: number, w: number, h: number) => ({ data: new Uint8ClampedArray(Math.max(0, w * h * 4)) }),
  };
  const ctx = new Proxy(target, {
    get(object, key) { return key in object ? object[key] : noop; },
    set(object, key, value) { object[key] = value; return true; },
  });
  return { ctx: ctx as unknown as CanvasRenderingContext2D, texts };
}

const decode = <T>(data: Uint8Array) => JSON.parse(new TextDecoder().decode(data)) as T;

/** The parsed workbook and every worksheet, pulled as the workbook does. */
function parse(bytes: Uint8Array): { workbook: ParsedWorkbook; sheets: Worksheet[] } {
  const archive = new XlsxArchive(bytes, undefined, undefined, undefined);
  try {
    const workbook = decode<ParsedWorkbook>(archive.parse());
    const sheets: Worksheet[] = [];
    workbook.workbook.sheets.forEach((meta, index) => {
      archive.open_sheet_cursor(index, meta.name);
      const rows: Row[] = [];
      for (let pull = 0; pull < 100_000; pull++) {
        const unit = decode<{ kind: 'preview' } | { kind: 'rows'; rows: Row[] } | { kind: 'finished'; worksheet: Worksheet }>(
          archive.pull_sheet_cursor(512));
        if (unit.kind === 'rows') rows.push(...unit.rows);
        if (unit.kind === 'finished') {
          archive.acknowledge_sheet_cursor_terminal();
          resolveSharedStringRows(rows, workbook.sharedStrings);
          unit.worksheet.rows = rows;
          sheets.push(unit.worksheet);
          break;
        }
      }
    });
    return { workbook, sheets };
  } finally {
    archive.free();
  }
}

function paintedScalars(workbook: ParsedWorkbook, ws: Worksheet): Set<number> {
  const { ctx, texts } = recordingContext();
  const lastRow = Math.max(40, ...ws.rows.map((row) => row.index));
  const lastCol = Math.max(16, ...ws.rows.flatMap((row) => row.cells.map((cell) => cell.col)));
  renderViewport(ctx, ws, workbook.styles, { row: 1, col: 1, rows: lastRow, cols: lastCol },
    { dpr: 1, loadedImages: new Map() }, 'jp');
  const scalars = new Set<number>();
  for (const text of texts) for (const ch of text) scalars.add(ch.codePointAt(0) as number);
  return scalars;
}

beforeAll(() => {
  initSync({ module: new Uint8Array(readFileSync(new URL('./wasm/xlsx_parser_bg.wasm', import.meta.url))) });
});

describe('XLSX font demand covers every measured or painted scalar', () => {
  const check = (label: string, bytes: Uint8Array): Array<'all' | 'finite'> => {
    const { workbook, sheets } = parse(bytes);
    const outcomes: Array<'all' | 'finite'> = [];
    for (const ws of sheets) {
      // One realm's demand as the workbook builds it: the workbook delta,
      // then the sheet's drawings and rows.
      const demand = new XlsxFontDemand(workbook, 'jp', []);
      const demanded = new Set<number>();
      let all = false;
      const take = () => {
        const delta = demand.take();
        if (delta?.scalars === 'all') all = true;
        else for (const cp of delta?.scalars ?? []) demanded.add(cp);
      };
      take();
      demand.worksheetParts(ws);
      demand.newRowsOf(ws);
      take();
      if (all) { outcomes.push('all'); continue; }
      const missing = [...paintedScalars(workbook, ws)].filter((cp) => !demanded.has(cp))
        .map((cp) => `U+${cp.toString(16).toUpperCase().padStart(4, '0')}`);
      expect(missing, `${label} sheet ${ws.name}`).toEqual([]);
      outcomes.push('finite');
    }
    return outcomes;
  };

  it('targeted synthetic workbooks', () => {
    for (const [label, bytes] of Object.entries({ numberFormats, richText, shapes })) {
      expect(check(label, bytes), label).toEqual(['finite']);
    }
  });

  it('public demo workbook', () => {
    const demo = new URL('../public/demo/', import.meta.url);
    const outcomes: Array<'all' | 'finite'> = [];
    for (const name of readdirSync(demo).filter((file) => file.endsWith('.xlsx'))) {
      outcomes.push(...check(name, new Uint8Array(readFileSync(new URL(name, demo)))));
    }
    // The invariant is vacuous for 'all' demand; record that some sheets are finite.
    expect(outcomes).toContain('finite');
  }, 120_000);
});

// Symbol-encoded faces (symbol-paint.ts): a code paints as authored where the
// face draws it, as its Unicode conversion where it does not; the demand
// covers both outcomes without declining to 'all'. (A precomposed authored
// code such as Wingdings 0xFC is not normalization-inert, so its text still
// demands 'all': core scalar-demand.ts.)
const symbols = workbookBytes(
  '<numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;a&quot;0"/></numFmts>'
    + '<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><sz val="11"/><name val="Wingdings"/></font>'
    + '<font><sz val="11"/><name val="Symbol"/></font></fonts>' + base
    + '<cellXfs count="4"><xf numFmtId="0" fontId="0"/><xf numFmtId="0" fontId="1"/><xf numFmtId="0" fontId="2"/>'
    + '<xf numFmtId="164" fontId="2" applyNumberFormat="1"/></cellXfs>',
  '<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>JJ l</t></is></c><c r="B1" s="2" t="s"><v>0</v></c>'
    + '<c r="C1" s="1" t="s"><v>1</v></c><c r="D1" s="3"><v>7</v></c><c r="E1" s="1" t="e"><v>#N/A</v></c></row>'
    + '<row r="2"><c r="A2" t="inlineStr"><is><r><t>Done </t></r><r><rPr><rFont val="Wingdings"/></rPr><t>n</t></r></is></c>'
    + '<c r="B2" s="1"><v>12</v></c></row>',
  '<si><t>abg</t></si><si><t>\uF0FC\uF06C</t></si>',
  '<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
    + '<xdr:to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
    + '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="Text 1"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>'
    + '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="1200000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>'
    + '<xdr:txBody><a:bodyPr wrap="square"/><a:p><a:r><a:rPr sz="1400"><a:sym typeface="Wingdings"/></a:rPr><a:t>\uF0FC done</a:t></a:r>'
    + '<a:r><a:rPr sz="1400"><a:latin typeface="Webdings"/></a:rPr><a:t>a</a:t></a:r></a:p></xdr:txBody>'
    + '</xdr:sp><xdr:clientData/></xdr:twoCellAnchor>',
);

describe('XLSX font demand covers symbol-encoded text in either face outcome', () => {
  it('stays finite and covers the painted form, kept or converted', () => {
    const { workbook, sheets } = parse(symbols);
    const ws = sheets[0]!;
    const demand = new XlsxFontDemand(workbook, 'jp', []);
    const deltas = [demand.take()];
    demand.worksheetParts(ws);
    demand.newRowsOf(ws);
    deltas.push(demand.take());
    expect(deltas.some((delta) => delta?.scalars === 'all')).toBe(false);
    const demanded = new Set(deltas.flatMap((delta) => delta?.scalars === 'all' ? [] : delta?.scalars ?? []));
    expect(deltas.flatMap((delta) => delta?.families ?? [])).toEqual(expect.arrayContaining(['Wingdings', 'Symbol', 'Webdings']));
    for (const realm of [symbolRealm(), symbolRealm({ wingdings: SYMBOL_ENCODED, symbol: SYMBOL_ENCODED, webdings: SYMBOL_ENCODED })]) {
      renderViewport(realm.ctx, ws, workbook.styles, { row: 1, col: 1, rows: 40, cols: 16 }, { dpr: 1, loadedImages: new Map() }, 'jp');
      const scalars = new Set<number>();
      for (const record of [...realm.fills, ...realm.measures]) {
        for (const ch of record.text) scalars.add(ch.codePointAt(0) as number);
      }
      const missing = [...scalars].filter((cp) => !demanded.has(cp))
        .map((cp) => `U+${cp.toString(16).toUpperCase().padStart(4, '0')}`);
      expect(missing).toEqual([]);
      // The conversion did run where the faces are missing.
      if (realm.fills.some((fill) => fill.text.includes('α'))) expect(scalars.has(0x2713)).toBe(true);
    }
  });
});

describe('cell text is charged to the face the renderer paints it with', () => {
  const styles = '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><sz val="11"/><name val="Sakkal Majalla"/></font></fonts>'
    + base + '<cellXfs count="2"><xf numFmtId="0" fontId="0"/><xf numFmtId="0" fontId="1"/></cellXfs>';
  const arabic = 'مرحبا';
  const painted = (workbook: ParsedWorkbook, ws: Worksheet) => {
    const fonts: string[] = [];
    let font = '10px serif';
    const noop = () => {};
    const ctx = new Proxy({
      get font() { return font; }, set font(value: string) { font = value; },
      canvas: { width: 800, height: 400 }, globalAlpha: 1,
      measureText: (text: string) => ({ width: [...text].length * 5, actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2, fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2 }),
      fillText: (text: string) => { if (text.includes(arabic)) fonts.push(font); },
      createLinearGradient: () => ({ addColorStop: noop }), createRadialGradient: () => ({ addColorStop: noop }),
      createPattern: () => null, getLineDash: () => [],
      getTransform: () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0, isIdentity: true }),
    } as Record<string | symbol, unknown>, {
      get(target, key) { return key in target ? target[key] : noop; },
      set(target, key, value) { target[key] = value; return true; },
    }) as unknown as CanvasRenderingContext2D;
    renderViewport(ctx, ws, workbook.styles, { row: 1, col: 1, rows: 5, cols: 3 }, { dpr: 1, loadedImages: new Map() }, 'jp');
    return fonts;
  };
  const provenFor = (sheetXml: string, cols = '') => {
    const { workbook, sheets } = parse(workbookBytes(styles, sheetXml, undefined, undefined, cols));
    const ws = sheets[0]!;
    const demand = new XlsxFontDemand(workbook, 'jp', OFFICE_FONT_SUBSTITUTIONS);
    demand.take();
    demand.worksheetParts(ws);
    demand.newRowsOf(ws);
    return { scripts: demand.take()?.substituteScripts?.['sakkal majalla'] ?? [], paint: painted(workbook, ws) };
  };

  it('a cell without c@s in a column styled with the scoped face proves its script', () => {
    const { scripts, paint } = provenFor(`<row r="1"><c r="A1" t="inlineStr"><is><t>${arabic}</t></is></c></row>`,
      '<cols><col min="1" max="1" width="20" style="1"/></cols>');
    expect(paint.length).toBeGreaterThan(0);
    for (const font of paint) expect(font).toContain('"Sakkal Majalla"');
    expect(scripts).toContain('arabic');
  });

  it('a row default style (not in the worksheet model) is not charged, exactly as the renderer paints', () => {
    const { scripts, paint } = provenFor(`<row r="1" s="1" customFormat="1"><c r="A1" t="inlineStr"><is><t>${arabic}</t></is></c></row>`);
    expect(paint.length).toBeGreaterThan(0);
    for (const font of paint) expect(font).not.toContain('Sakkal Majalla');
    expect(scripts).toEqual([]);
  });
});

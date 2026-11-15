/// <reference types="node" />
import { readFileSync } from 'node:fs';
import { beforeAll, expect, it } from 'vitest';
// @ts-ignore — shared synthetic ZIP fixture is a plain JavaScript module.
import { storedZip } from '../../../tests/fixtures/chart-ex-packages.mjs';
import { initSync, XlsxArchive } from './wasm/xlsx_parser.js';
import { renderViewport } from './renderer.js';
import type { ParsedWorkbook, Row, Worksheet } from './types.js';

// ECMA-376 §18.3.1.53 `<is>` inline string with §18.4.6 `<rPh>` furigana and
// `ph="1"`. The reading must survive the parser's JSON (as `phoneticRuns` /
// `phoneticPr`, the names the TS model reads) and reach the canvas. It used to
// be serialized as `phonetic_runs`, so inline-string furigana never showed.

beforeAll(() => {
  initSync({ module: new Uint8Array(readFileSync(new URL('./wasm/xlsx_parser_bg.wasm', import.meta.url))) });
});

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function inlinePhoneticWorkbook(): Uint8Array {
  return storedZip([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      + '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
      + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
      + '</Types>'],
    ['_rels/.rels', `<Relationships xmlns="${REL}">`
      + `<Relationship Id="rId1" Type="${OFFICE_REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ['xl/workbook.xml', `<workbook xmlns="${MAIN}" xmlns:r="${OFFICE_REL}">`
      + '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'],
    ['xl/_rels/workbook.xml.rels', `<Relationships xmlns="${REL}">`
      + `<Relationship Id="rId1" Type="${OFFICE_REL}/worksheet" Target="worksheets/sheet1.xml"/>`
      + `<Relationship Id="rId2" Type="${OFFICE_REL}/styles" Target="styles.xml"/></Relationships>`],
    ['xl/styles.xml', `<styleSheet xmlns="${MAIN}">`
      + '<fonts count="2"><font><name val="Calibri"/><sz val="11"/></font><font><name val="Calibri"/><sz val="6"/></font></fonts>'
      + '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
      + '<borders count="1"><border/></borders><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>'
      + '</styleSheet>'],
    ['xl/worksheets/sheet1.xml', `<worksheet xmlns="${MAIN}"><sheetData><row r="1" ht="30" customHeight="1">`
      + '<c r="A1" t="inlineStr" ph="1"><is><t>森</t><rPh sb="0" eb="1"><t>モリ</t></rPh><phoneticPr fontId="1"/></is></c>'
      + '</row></sheetData></worksheet>'],
  ]) as Uint8Array;
}

function parse(bytes: Uint8Array): { workbook: ParsedWorkbook; worksheet: Worksheet } {
  const archive = new XlsxArchive(bytes, undefined, undefined, undefined);
  const decode = (data: Uint8Array) => JSON.parse(new TextDecoder().decode(data));
  try {
    const workbook = decode(archive.parse()) as ParsedWorkbook;
    archive.open_sheet_cursor(0, 'Sheet1');
    const rows: Row[] = [];
    for (let pull = 0; pull < 4; pull++) {
      const unit = decode(archive.pull_sheet_cursor(128)) as
        | { kind: 'preview' }
        | { kind: 'rows'; rows: Row[] }
        | { kind: 'finished'; worksheet: Worksheet };
      if (unit.kind === 'rows') rows.push(...unit.rows);
      if (unit.kind === 'finished') {
        archive.acknowledge_sheet_cursor_terminal();
        return { workbook, worksheet: { ...unit.worksheet, rows } };
      }
    }
    throw new Error('single-row worksheet did not finish');
  } finally {
    archive.free();
  }
}

function recordingCtx() {
  const texts: string[] = [];
  let font = '11px sans-serif';
  const noop = () => {};
  const ctx: Record<string, unknown> = {
    canvas: { width: 400, height: 200 },
    get font() { return font; },
    set font(v: string) { font = v; },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textBaseline: 'alphabetic', textAlign: 'left', letterSpacing: '0px', direction: 'ltr', globalAlpha: 1,
    measureText: (s: string) => ({ width: [...s].length * 8 }),
    fillText: (t: string) => { texts.push(t); },
    strokeText: noop, fillRect: noop, strokeRect: noop, clearRect: noop, beginPath: noop, closePath: noop,
    moveTo: noop, lineTo: noop, rect: noop, arc: noop, fill: noop, stroke: noop, clip: noop, save: noop,
    restore: noop, translate: noop, rotate: noop, scale: noop, setLineDash: noop, setTransform: noop,
    createLinearGradient: () => ({ addColorStop: noop }),
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, texts };
}

it('carries inline-string furigana through the parser JSON to the phonetic band', () => {
  const { workbook, worksheet } = parse(inlinePhoneticWorkbook());
  const cell = worksheet.rows[0].cells[0];
  expect(cell.showPhonetic).toBe(true);
  expect(cell.value).toMatchObject({
    type: 'text',
    text: '森',
    phoneticRuns: [{ sb: 0, eb: 1, text: 'モリ' }],
    phoneticPr: { fontId: 1 },
  });
  expect(JSON.stringify(cell.value)).not.toMatch(/phonetic_/);

  const { ctx, texts } = recordingCtx();
  renderViewport(ctx, worksheet, workbook.styles, { row: 1, col: 1, rows: 1, cols: 1 });
  expect(texts).toContain('森');
  expect(texts).toContain('モリ');
});

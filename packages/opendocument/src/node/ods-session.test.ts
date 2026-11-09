// OpenDocument spreadsheets through the Node XLSX session: the claimed
// package is projected into the XLSX model and displayed with the XLSX
// number formatter.
import { describe, expect, it } from 'vitest';
import { buildOdsFixture } from '../ods-fixture.js';
import { testOdsSource } from '../test-sources.js';
import { formatCellValue, materializeXlsxWorkbook, skia, skiaFactory } from './ods-facade.js';

type Cell = { col: number; row: number; value: { type: string; text?: string; number?: number }; styleIndex?: number };
type Worksheet = {
  name: string;
  rows: Array<{ index: number; cells: Cell[]; hidden?: boolean }>;
  mergeCells: Array<{ top: number; left: number; bottom: number; right: number }>;
  colWidthRanges?: Array<{ min: number; max: number; width: number }>;
  defaultColWidth: number;
  rowHeights: Record<string, number>;
  images: Array<{ imagePath: string; mimeType: string; fromCol: number; fromRow: number; toCol: number; toRow: number }>;
  charts: Array<{ chart: { chartType: string; series: Array<{ name: string; values: Array<number | null> }>; categories: string[] } }>;
  hyperlinks: Array<{ col: number; row: number; url?: string | null; location?: string | null }>;
  definedNames: Array<{ name: string; formula: string }>;
  conditionalFormats: unknown[];
};

const open = async (bytes: Uint8Array) => {
  const workbook = await materializeXlsxWorkbook(bytes, { ods: testOdsSource(), ...(skia ? { factory: skiaFactory() } : {}) });
  return {
    workbook,
    sheets: workbook.worksheets as unknown as Worksheet[],
    text: (sheet: number, row: number, col: number) => {
      const cell = (workbook.worksheets[sheet] as unknown as Worksheet).rows
        .find(r => r.index === row)?.cells.find(c => c.col === col);
      return cell ? formatCellValue(cell as never, workbook.workbookIndex.styles, null, workbook.workbookIndex.workbook.date1904) : '';
    },
  };
};

const NUMBER_STYLES = '<number:number-style style:name="N2"><number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/></number:number-style>'
  + '<number:percentage-style style:name="P1"><number:number number:decimal-places="1" number:min-decimal-places="1" number:min-integer-digits="1"/><number:text>%</number:text></number:percentage-style>'
  + '<number:date-style style:name="D1"><number:year number:style="long"/><number:text>/</number:text><number:month number:style="long"/><number:text>/</number:text><number:day number:style="long"/></number:date-style>'
  + '<number:time-style style:name="T1" number:truncate-on-overflow="false"><number:hours/><number:text>:</number:text><number:minutes number:style="long"/></number:time-style>'
  + '<number:currency-style style:name="C1P0"><number:currency-symbol>$</number:currency-symbol><number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/></number:currency-style>'
  + '<number:currency-style style:name="C1"><number:text>(</number:text><number:currency-symbol>$</number:currency-symbol><number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/><number:text>)</number:text><style:map style:condition="value()&gt;=0" style:apply-style-name="C1P0"/></number:currency-style>';

describe('ODS through the XLSX Node session', () => {
  it('displays every value type through its data style, with huge repeats kept as runs', async () => {
    const cellStyles = ['N2', 'P1', 'D1', 'T1', 'C1']
      .map(n => `<style:style style:name="ce${n}" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="${n}"/>`).join('');
    const { sheets, text } = await open(buildOdsFixture({
      automaticStyles: NUMBER_STYLES + cellStyles
        + '<style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="2.54cm"/></style:style>'
        + '<style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="0.6cm"/></style:style>',
      body: '<table:table table:name="Data"><table:table-column table:style-name="co1" table:number-columns-repeated="16384"/>'
        + '<table:table-row table:style-name="ro1">'
        + '<table:table-cell table:style-name="ceN2" office:value-type="float" office:value="1234.5"><text:p>1,234.50</text:p></table:table-cell>'
        + '<table:table-cell table:style-name="ceP1" office:value-type="percentage" office:value="0.256"><text:p>25.6%</text:p></table:table-cell>'
        + '<table:table-cell table:style-name="ceD1" office:value-type="date" office:date-value="2024-03-05"><text:p>2024/03/05</text:p></table:table-cell>'
        + '<table:table-cell table:style-name="ceT1" office:value-type="time" office:time-value="PT27H30M00S"><text:p>27:30</text:p></table:table-cell>'
        + '<table:table-cell table:style-name="ceC1" office:value-type="currency" office:value="-12" office:currency="USD"><text:p>($12.00)</text:p></table:table-cell>'
        + '<table:table-cell office:value-type="boolean" office:boolean-value="true"><text:p>TRUE</text:p></table:table-cell>'
        + '<table:table-cell office:value-type="string" table:number-columns-spanned="2"><text:p>a <text:s text:c="2"/>b</text:p><text:p>line</text:p></table:table-cell>'
        + '<table:covered-table-cell/>'
        + '<table:table-cell table:number-columns-repeated="16376"/></table:table-row>'
        + '<table:table-row table:style-name="ro1" table:number-rows-repeated="1048575"><table:table-cell table:number-columns-repeated="16384"/></table:table-row>'
        + '</table:table>',
    }));
    expect(sheets[0]!.name).toBe('Data');
    expect([1, 2, 3, 4, 5, 6, 7].map(col => text(0, 1, col)))
      .toEqual(['1,234.50', '25.6%', '2024/03/05', '27:30', '($12.00)', 'TRUE', 'a   b\nline']);
    expect(sheets[0]!.mergeCells).toEqual([{ top: 1, left: 7, bottom: 1, right: 8 }]);
    // One content row; the million-row filler is the default height.
    expect(sheets[0]!.rows.map(r => r.index)).toEqual([1]);
    expect(Object.keys(sheets[0]!.rowHeights)).toEqual([]);
  });

  it('maps cell formats, spans, links and named ranges', async () => {
    const { workbook, sheets } = await open(buildOdsFixture({
      automaticStyles: '<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default">'
        + '<style:table-cell-properties fo:background-color="#ffff00" fo:border="0.74pt solid #ff0000" style:text-align-source="fix" style:vertical-align="middle" fo:wrap-option="wrap" style:rotation-angle="90"/>'
        + '<style:paragraph-properties fo:text-align="center"/><style:text-properties fo:font-weight="bold" fo:font-size="14pt" fo:color="#0000ff"/></style:style>'
        + '<style:style style:name="T1" style:family="text"><style:text-properties fo:font-style="italic" fo:color="#ff0000"/></style:style>',
      body: '<table:table table:name="S"><table:table-column table:number-columns-repeated="3"/>'
        + '<table:table-row><table:table-cell table:style-name="ce1" office:value-type="string"><text:p>x<text:span text:style-name="T1">y</text:span></text:p></table:table-cell>'
        + '<table:table-cell office:value-type="string"><text:p><text:a xlink:href="https://example.test/">link</text:a></text:p></table:table-cell>'
        + '<table:table-cell table:style-name="ce1"/></table:table-row></table:table>'
        + '<table:named-expressions><table:named-range table:name="Area" table:base-cell-address="$S.$A$1" table:cell-range-address="$S.$A$1:.$B$1"/></table:named-expressions>',
    }));
    const styles = workbook.workbookIndex.styles;
    const cell = sheets[0]!.rows[0]!.cells[0]!;
    const xf = styles.cellXfs[cell.styleIndex!]!;
    expect(xf).toMatchObject({ alignH: 'center', alignV: 'center', wrapText: true, textRotation: 90 });
    expect(styles.fonts[xf.fontId]).toMatchObject({ bold: true, size: 14, color: '0000FF' });
    expect(styles.fills[xf.fillId]).toMatchObject({ patternType: 'solid', fgColor: 'FFFF00' });
    expect(styles.borders[xf.borderId]?.left).toEqual({ style: 'thin', color: 'FF0000' });
    expect(cell.value).toMatchObject({ type: 'text', text: 'xy' });
    const runs = (cell.value as { runs?: Array<{ text: string; font?: { italic: boolean; bold: boolean } }> }).runs;
    expect(runs?.map(r => [r.text, r.font?.italic, r.font?.bold])).toEqual([['x', false, true], ['y', true, true]]);
    // An empty styled cell is materialized because it paints a fill.
    expect(sheets[0]!.rows[0]!.cells.map(c => c.col)).toEqual([1, 2, 3]);
    expect(sheets[0]!.hyperlinks).toEqual([{ col: 2, row: 1, url: 'https://example.test/', location: null, display: null }]);
    expect(sheets[0]!.definedNames).toEqual([{ name: 'Area', formula: 'S!$A$1:$B$1' }]);
  });

  it('anchors images and reads chart objects from their cell ranges', async () => {
    const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10"/></svg>';
    const chart = `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${await import('../ods-fixture.js').then(m => m.ODS_NS)} office:version="1.3"><office:body><office:chart>`
      + '<chart:chart chart:class="chart:line"><chart:plot-area><chart:axis chart:dimension="x" chart:name="primary-x"><chart:categories table:cell-range-address="S.A1:S.A2"/></chart:axis>'
      + '<chart:series chart:values-cell-range-address="S.B1:S.B2" chart:label-cell-address="S.C1"/></chart:plot-area></chart:chart></office:chart></office:body></office:document-content>';
    const { sheets } = await open(buildOdsFixture({
      body: '<table:table table:name="S"><table:table-column table:number-columns-repeated="8"/>'
        + '<table:table-row><table:table-cell office:value-type="string"><text:p>Q1</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="3"/><table:table-cell office:value-type="string"><text:p>Sales</text:p></table:table-cell></table:table-row>'
        + '<table:table-row><table:table-cell office:value-type="string"><text:p>Q2</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="5"/>'
        + '<table:table-cell><draw:frame svg:width="2cm" svg:height="1cm" svg:x="0cm" svg:y="0cm"><draw:image xlink:href="Pictures/a.svg"/></draw:frame>'
        + '<draw:frame table:end-cell-address="S.H9" table:end-x="0cm" table:end-y="0cm" svg:width="8cm" svg:height="5cm" svg:x="0cm" svg:y="0cm"><draw:object xlink:href="./Object 1"/></draw:frame></table:table-cell></table:table-row></table:table>',
      extraFiles: [['Pictures/a.svg', svg], ['Object 1/content.xml', chart]],
    }));
    const [image] = sheets[0]!.images;
    expect(image).toMatchObject({ imagePath: 'Pictures/a.svg', mimeType: 'image/svg+xml', fromCol: 2, fromRow: 1 });
    const [anchored] = sheets[0]!.charts;
    expect(anchored!.chart.chartType).toBe('line');
    expect(anchored!.chart.categories).toEqual(['Q1', 'Q2']);
    expect(anchored!.chart.series.map(s => [s.name, s.values])).toEqual([['Sales', [3, 5]]]);
  });

  it('maps conditional formats, validations and sparklines over the cells', async () => {
    const { sheets, workbook } = await open(buildOdsFixture({
      styles: '<style:style style:name="Good" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:background-color="#c6efce"/></style:style>',
      body: '<table:content-validations><table:content-validation table:name="val1" table:condition="of:cell-content-is-in-list(&quot;a&quot;;&quot;b&quot;)" table:allow-empty-cell="true"/></table:content-validations>'
        + '<table:table table:name="S"><table:table-column table:number-columns-repeated="4"/>'
        + '<table:table-row><table:table-cell office:value-type="float" office:value="1"/><table:table-cell office:value-type="float" office:value="5"/>'
        + '<table:table-cell table:content-validation-name="val1" office:value-type="string"><text:p>a</text:p></table:table-cell><table:table-cell/></table:table-row>'
        + '<calcext:conditional-formats><calcext:conditional-format calcext:target-range-address="S.A1:S.B1">'
        + '<calcext:condition calcext:apply-style-name="Good" calcext:value="&gt;2" calcext:base-cell-address="S.A1"/></calcext:conditional-format>'
        + '<calcext:conditional-format calcext:target-range-address="S.A1:S.B1"><calcext:color-scale>'
        + '<calcext:color-scale-entry calcext:value="0" calcext:type="minimum" calcext:color="#ffffff"/><calcext:color-scale-entry calcext:value="0" calcext:type="maximum" calcext:color="#ff0000"/>'
        + '</calcext:color-scale></calcext:conditional-format></calcext:conditional-formats>'
        + '<calcext:sparkline-groups><calcext:sparkline-group calcext:type="column" calcext:color-series="#376092"><calcext:sparklines>'
        + '<calcext:sparkline calcext:cell-address="S.D1" calcext:data-range="S.A1:S.B1"/></calcext:sparklines></calcext:sparkline-group></calcext:sparkline-groups>'
        + '</table:table>',
    }));
    const sheet = sheets[0] as unknown as {
      conditionalFormats: Array<{ sqref: unknown[]; rules: Array<{ type: string; operator?: string; formulas?: string[]; dxfId?: number }> }>;
      dataValidations: Array<{ sqref: string; validationType: string; formula1: string }>;
      sparklineGroups: Array<{ kind: string; sparklines: Array<{ row: number; col: number; values: Array<number | null> }> }>;
    };
    const [cellIs, scale] = sheet.conditionalFormats.map(f => f.rules[0]!);
    expect(cellIs).toMatchObject({ type: 'cellIs', operator: 'greaterThan', formulas: ['2'] });
    expect(workbook.workbookIndex.styles.dxfs[cellIs!.dxfId!]?.fill).toMatchObject({ fgColor: 'C6EFCE' });
    expect(scale?.type).toBe('colorScale');
    expect(sheet.dataValidations).toEqual([expect.objectContaining({ sqref: 'C1', validationType: 'list', formula1: '"a,b"' })]);
    expect(sheet.sparklineGroups[0]).toMatchObject({ kind: 'column', sparklines: [{ row: 1, col: 4, values: [1, 5] }] });
  });

  it('rejects encrypted and corrupt packages, claims only spreadsheets', async () => {
    const source = testOdsSource();
    expect(source.claim(buildOdsFixture({ body: '<table:table table:name="S"/>' }))).toBe(true);
    expect(source.claim(buildOdsFixture({ body: '', mediaType: 'application/vnd.oasis.opendocument.text' }))).toBe(false);
    const template = buildOdsFixture({ body: '<table:table table:name="S"/>', mediaType: 'application/vnd.oasis.opendocument.spreadsheet-template' });
    expect(source.claim(template)).toBe(true);
    await expect(open(buildOdsFixture({ body: '<table:table' }))).rejects.toBeInstanceOf(Error);
  });
});

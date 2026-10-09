import { storedZip } from './test-ooxml-package.ts';

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const SHEET_TYPE = `${REL_NS}/worksheet`;

/** One cache-free clustered bar chart whose only series value source is the
 * unqualified defined name `Revenue` (ECMA-376 Part 1 §21.2.2.123 `numRef`:
 * `numCache` is optional, so values come from the referenced cells only). */
const CHART_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>Revenue</c:f></c:numRef></c:val></c:ser><c:axId val="10"/><c:axId val="20"/></c:barChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>`;

function drawingXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${REL_NS}"><xdr:twoCellAnchor><xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>11</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`;
}

function relationships(entries: ReadonlyArray<readonly [id: string, type: string, target: string]>): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PACKAGE_REL_NS}">${entries.map(([id, type, target]) =>
    `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`).join('')}</Relationships>`;
}

function worksheetXml(sheetData: string, drawing: boolean): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><sheetData>${sheetData}</sheetData>${drawing ? '<drawing r:id="rIdDrawing"/>' : ''}</worksheet>`;
}

/**
 * Authored (not Office-exported) XLSX for issue #1547 chart-ingress controls.
 *
 * Sheets, in workbook order: `Data` (index 0, chart), `Other` (index 1, no
 * drawing) and `Plain` (index 2, chart). `Data!B2:B4` = 11, 12, 13;
 * `Data!C2:C4` = 21, 22, 23; `Data!D2:D4` = 31, 32, 33.
 *
 * Three same-spelling `Revenue` definitions are authored in this order:
 * workbook scope → `Data!$C$2:$C$4`, `localSheetId="0"` → `Data!$B$2:$B$4`,
 * `localSheetId="1"` → `Data!$D$2:$D$4` (ECMA-376 Part 1 §18.2.5). Both
 * charts reference the unqualified name through a cache-free `c:numRef`.
 * Under the existing unqualified-name resolution contract the `Data` chart reads
 * its local B values, `Plain` reads the global C values, and the `Other`
 * local is never visible to either chart.
 */
export function scopedNameChartXlsx(): Uint8Array {
  const dataRows = [2, 3, 4].map((row) =>
    `<row r="${row}"><c r="B${row}"><v>${9 + row}</v></c><c r="C${row}"><v>${19 + row}</v></c><c r="D${row}"><v>${29 + row}</v></c></row>`).join('');
  return storedZip({
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/><Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>`,
    '_rels/.rels': relationships([
      ['rId1', `${REL_NS}/officeDocument`, 'xl/workbook.xml'],
    ]),
    'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/><sheet name="Plain" sheetId="3" r:id="rId3"/></sheets><definedNames><definedName name="Revenue">Data!$C$2:$C$4</definedName><definedName name="Revenue" localSheetId="0">Data!$B$2:$B$4</definedName><definedName name="Revenue" localSheetId="1">Data!$D$2:$D$4</definedName></definedNames></workbook>`,
    'xl/_rels/workbook.xml.rels': relationships([
      ['rId1', SHEET_TYPE, 'worksheets/sheet1.xml'],
      ['rId2', SHEET_TYPE, 'worksheets/sheet2.xml'],
      ['rId3', SHEET_TYPE, 'worksheets/sheet3.xml'],
      ['rId4', `${REL_NS}/styles`, 'styles.xml'],
    ]),
    'xl/styles.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${MAIN_NS}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`,
    'xl/worksheets/sheet1.xml': worksheetXml(dataRows, true),
    'xl/worksheets/_rels/sheet1.xml.rels': relationships([
      ['rIdDrawing', `${REL_NS}/drawing`, '../drawings/drawing1.xml'],
    ]),
    'xl/worksheets/sheet2.xml': worksheetXml('', false),
    'xl/worksheets/sheet3.xml': worksheetXml('', true),
    'xl/worksheets/_rels/sheet3.xml.rels': relationships([
      ['rIdDrawing', `${REL_NS}/drawing`, '../drawings/drawing2.xml'],
    ]),
    'xl/drawings/drawing1.xml': drawingXml(),
    'xl/drawings/_rels/drawing1.xml.rels': relationships([
      ['rIdChart', `${REL_NS}/chart`, '../charts/chart1.xml'],
    ]),
    'xl/drawings/drawing2.xml': drawingXml(),
    'xl/drawings/_rels/drawing2.xml.rels': relationships([
      ['rIdChart', `${REL_NS}/chart`, '../charts/chart2.xml'],
    ]),
    'xl/charts/chart1.xml': CHART_XML,
    'xl/charts/chart2.xml': CHART_XML,
  });
}

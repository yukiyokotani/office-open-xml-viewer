// Self-authored fixture: number formats whose decimal and group tokens the
// viewer paints with the host locale's separators (#2346), next to a quoted
// literal dot that stays a dot.
import { storedZip } from './chart-ex-packages.mjs';

export function numberSeparatorsXlsxBytes() {
  const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const pkg = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const rows = [
    '<row r="1"><c r="A1" s="1"><v>1609.344</v></c></row>',
    '<row r="2"><c r="A2"><v>39.37007874</v></c></row>',
    '<row r="3"><c r="A3" s="2"><v>3.785411784</v></c></row>',
    '<row r="4"><c r="A4" s="3"><v>1234567.891</v></c></row>',
  ].join('');
  return storedZip([
    ['[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`],
    ['_rels/.rels', `<Relationships xmlns="${pkg}"><Relationship Id="r1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ['xl/workbook.xml', `<workbook xmlns="${main}" xmlns:r="${rel}"><sheets><sheet name="Separators" sheetId="1" r:id="s1"/></sheets></workbook>`],
    ['xl/_rels/workbook.xml.rels', `<Relationships xmlns="${pkg}"><Relationship Id="s1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="styles" Type="${rel}/styles" Target="styles.xml"/></Relationships>`],
    ['xl/styles.xml', `<styleSheet xmlns="${main}"><numFmts count="2"><numFmt numFmtId="164" formatCode="#,##0.000"/><numFmt numFmtId="165" formatCode="0.00&quot; l.&quot;"/></numFmts><fonts count="1"><font><sz val="14"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/><xf numFmtId="4" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/></cellXfs></styleSheet>`],
    ['xl/worksheets/sheet1.xml', `<worksheet xmlns="${main}"><cols><col min="1" max="1" width="24" customWidth="1"/></cols><sheetData>${rows}</sheetData></worksheet>`],
  ]);
}

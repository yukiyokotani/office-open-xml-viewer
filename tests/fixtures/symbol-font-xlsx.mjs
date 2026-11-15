// Self-authored fixture: symbol-encoded faces in cells, rich runs, a wrapped
// cell, number-format output and a text box with an `a:sym` run (#2177 for
// XLSX). Wingdings check marks and bullets, Symbol Greek.
import { storedZip } from './chart-ex-packages.mjs';

export function symbolFontXlsxBytes() {
  const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const pkg = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const inline = (ref, value, style) =>
    `<c r="${ref}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${value}</t></is></c>`;
  const rows = [
    `<row r="1">${inline('A1', 'ü þ ý û', 1)}${inline('B1', 'abgd pqr', 2)}</row>`,
    `<row r="2"><c r="A2" t="inlineStr"><is><r><rPr><rFont val="Wingdings"/><sz val="11"/></rPr><t>l</t></r>`
      + `<r><rPr><rFont val="Calibri"/><sz val="11"/></rPr><t xml:space="preserve"> first item</t></r></is></c>`
      + `<c r="B2" s="4"><v>42</v></c></row>`,
    `<row r="3">${inline('A3', 'JJJJ KKKK LLLL JJJJ KKKK LLLL', 3)}${inline('B3', 'a + b = g', 2)}</row>`,
  ].join('');
  const drawing = '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    + '<xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
    + '<xdr:to><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
    + '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="Text 1"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>'
    + '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2400000" cy="800000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
    + '<a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></xdr:spPr>'
    + '<xdr:txBody><a:bodyPr wrap="square"/><a:p><a:r><a:rPr sz="1400"><a:latin typeface="Calibri"/><a:sym typeface="Wingdings"/></a:rPr><a:t>&#xF0FC; done &#xF0FB; open</a:t></a:r></a:p>'
    + '<a:p><a:r><a:rPr sz="1400"><a:latin typeface="Symbol"/></a:rPr><a:t>abg</a:t></a:r></a:p></xdr:txBody>'
    + '</xdr:sp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>';
  return storedZip([
    ['[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`],
    ['_rels/.rels', `<Relationships xmlns="${pkg}"><Relationship Id="r1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ['xl/workbook.xml', `<workbook xmlns="${main}" xmlns:r="${rel}"><sheets><sheet name="Symbols" sheetId="1" r:id="s1"/></sheets></workbook>`],
    ['xl/_rels/workbook.xml.rels', `<Relationships xmlns="${pkg}"><Relationship Id="s1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="styles" Type="${rel}/styles" Target="styles.xml"/></Relationships>`],
    ['xl/styles.xml', `<styleSheet xmlns="${main}"><numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;a = &quot;0"/></numFmts><fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><sz val="14"/><name val="Wingdings"/></font><font><sz val="12"/><name val="Symbol"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0"/><xf numFmtId="0" fontId="2" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyAlignment="1"><alignment wrapText="1"/></xf><xf numFmtId="164" fontId="2" fillId="0" borderId="0" applyNumberFormat="1"/></cellXfs></styleSheet>`],
    ['xl/worksheets/sheet1.xml', `<worksheet xmlns="${main}" xmlns:r="${rel}"><cols><col min="1" max="1" width="18" customWidth="1"/><col min="2" max="2" width="16" customWidth="1"/></cols><sheetData>${rows}</sheetData><drawing r:id="d1"/></worksheet>`],
    ['xl/worksheets/_rels/sheet1.xml.rels', `<Relationships xmlns="${pkg}"><Relationship Id="d1" Type="${rel}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`],
    ['xl/drawings/drawing1.xml', drawing],
  ]);
}

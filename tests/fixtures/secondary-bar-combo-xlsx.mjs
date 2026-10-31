// Self-authored fixture: a clustered column group on the primary axes and a
// second clustered column group on a secondary value axis whose category axis
// is deleted (the usual Excel "column on secondary axis" form), with variants
// that cannot share the primary categories.
import { storedZip } from './chart-ex-packages.mjs';

/**
 * @param {{ deletedAxPos?: 't' | 'b', dateAxis?: boolean, midCat?: boolean,
 *   reversed?: boolean, shortFirst?: boolean }} variant
 */
export function secondaryBarComboXlsxBytes(variant = {}) {
  const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const pkg = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const c = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const xdr = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
  const orientation = variant.reversed ? 'maxMin' : 'minMax';
  const points = (values) => values
    .map((value, index) => `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`).join('');
  const series = (index, values) => `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>`
    + `<c:tx><c:v>S${index}</c:v></c:tx>`
    + `<c:cat><c:strLit><c:ptCount val="3"/>${points(['A', 'B', 'C'])}</c:strLit></c:cat>`
    + `<c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>`
    + `${points(values)}</c:numLit></c:val></c:ser>`;
  const group = (index, values, axes) => `<c:barChart><c:barDir val="col"/>`
    + `<c:grouping val="clustered"/><c:varyColors val="0"/>${series(index, values)}`
    + `<c:gapWidth val="150"/><c:axId val="${axes[0]}"/><c:axId val="${axes[1]}"/></c:barChart>`;
  const primaryCategory = variant.dateAxis ? 'dateAx' : 'catAx';
  const primaryCategoryTail = variant.dateAxis
    ? '<c:auto val="1"/><c:lblOffset val="100"/><c:baseTimeUnit val="days"/>'
    : '<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>';
  const axes = `<c:${primaryCategory}><c:axId val="1"/><c:scaling><c:orientation val="${orientation}"/></c:scaling>`
    + '<c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/>'
    + `<c:tickLblPos val="nextTo"/><c:crossAx val="2"/><c:crosses val="autoZero"/>${primaryCategoryTail}</c:${primaryCategory}>`
    + '<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
    + '<c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/>'
    + '<c:tickLblPos val="nextTo"/><c:crossAx val="1"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>'
    + '<c:valAx><c:axId val="4"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
    + '<c:axPos val="r"/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/>'
    + `<c:crossAx val="3"/><c:crosses val="max"/><c:crossBetween val="${variant.midCat ? 'midCat' : 'between'}"/></c:valAx>`
    + `<c:catAx><c:axId val="3"/><c:scaling><c:orientation val="${orientation}"/></c:scaling><c:delete val="1"/>`
    + `<c:axPos val="${variant.deletedAxPos ?? 'b'}"/><c:numFmt formatCode="General" sourceLinked="1"/>`
    + '<c:tickLblPos val="nextTo"/><c:crossAx val="4"/><c:crosses val="autoZero"/><c:auto val="1"/>'
    + '<c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>';
  const chart = `<c:chartSpace xmlns:c="${c}" xmlns:a="${a}"><c:chart><c:autoTitleDeleted val="1"/>`
    + `<c:plotArea><c:layout/>${group(0, variant.shortFirst ? [10] : [10, 20, 30], [1, 2])}`
    + `${group(1, [300, 100, 200], [3, 4])}${axes}</c:plotArea>`
    + '<c:legend><c:legendPos val="b"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>';
  const anchor = '<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row>'
    + '<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>18</xdr:row>'
    + '<xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart"/>'
    + '<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
    + `<a:graphic><a:graphicData uri="${c}"><c:chart r:id="c1"/></a:graphicData></a:graphic></xdr:graphicFrame>`
    + '<xdr:clientData/></xdr:twoCellAnchor>';
  return storedZip([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
      + '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
      + '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
      + '<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
      + '</Types>'],
    ['_rels/.rels', `<Relationships xmlns="${pkg}"><Relationship Id="r1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ['xl/workbook.xml', `<workbook xmlns="${main}" xmlns:r="${rel}"><sheets><sheet name="Combo" sheetId="1" r:id="s1"/></sheets></workbook>`],
    ['xl/_rels/workbook.xml.rels', `<Relationships xmlns="${pkg}"><Relationship Id="s1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/>`
      + `<Relationship Id="styles" Type="${rel}/styles" Target="styles.xml"/></Relationships>`],
    ['xl/styles.xml', `<styleSheet xmlns="${main}"><fonts count="1"><font><sz val="11"/><name val="Arial"/></font></fonts>`
      + '<fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders>'
      + '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>'],
    ['xl/worksheets/sheet1.xml', `<worksheet xmlns="${main}" xmlns:r="${rel}"><sheetData/><drawing r:id="d1"/></worksheet>`],
    ['xl/worksheets/_rels/sheet1.xml.rels', `<Relationships xmlns="${pkg}"><Relationship Id="d1" Type="${rel}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`],
    ['xl/drawings/drawing1.xml', `<xdr:wsDr xmlns:xdr="${xdr}" xmlns:a="${a}" xmlns:r="${rel}" xmlns:c="${c}">${anchor}</xdr:wsDr>`],
    ['xl/drawings/_rels/drawing1.xml.rels', `<Relationships xmlns="${pkg}"><Relationship Id="c1" Type="${rel}/chart" Target="../charts/chart1.xml"/></Relationships>`],
    ['xl/charts/chart1.xml', chart],
  ]);
}

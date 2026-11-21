import { XLSX } from '../lib/grid.mjs';
import { CHART_CT, columnChart } from '../lib/chart.mjs';
import { A_SERIES } from '../lib/iso216.mjs';
import { R, sheet } from '../lib/xlsx.mjs';
import { relationships as relationshipsXml } from '../lib/opc.mjs';
import { XML_DECL } from '../lib/xml.mjs';

const { MARGIN_COL, MODULE, FIRST_ROW } = XLSX;
const ROWS = A_SERIES.slice(0, 7);
const first = FIRST_ROW + 1;
const last = first + ROWS.length - 1;
const SHEET = 'xlsx-009';

// The chart frame is anchored to cell edges: from the top left of E6 to the
// top left of J(last + 1), so it spans whole columns and rows of the grid.
const anchorCell = (col, row) => `<xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff>`;
const DRAWING = XML_DECL + '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
  + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
  + `<xdr:twoCellAnchor editAs="oneCell"><xdr:from>${anchorCell(4, FIRST_ROW - 1)}</xdr:from><xdr:to>${anchorCell(9, last)}</xdr:to>`
  + '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
  + '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
  + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
  + `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ${R} r:id="rId1"/>`
  + '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>';

export default {
  id: 'xlsx-009',
  slug: 'chart',
  format: 'xlsx',
  title: 'Chart',
  covers: [
    'xl/worksheets/sheet1.xml|x:worksheet>x:drawing',
    'xl/worksheets/sheet1.xml|x:drawing@r:id',
    'xl/drawings/drawing1.xml|xdr:wsDr>xdr:twoCellAnchor',
    'xl/drawings/drawing1.xml|xdr:twoCellAnchor@editAs=oneCell',
    'xl/drawings/drawing1.xml|xdr:from>xdr:col',
    'xl/drawings/drawing1.xml|xdr:twoCellAnchor>xdr:graphicFrame',
    'xl/drawings/drawing1.xml|a:graphicData>c:chart',
    'xl/charts/chart1.xml|c:cat>c:strRef',
    'xl/charts/chart1.xml|c:val>c:numRef',
    'xl/charts/chart1.xml|c:numRef>c:f',
    'xl/charts/chart1.xml|c:numRef>c:numCache',
  ],
  parts() {
    const s = sheet({
      id: SHEET,
      title: 'Chart',
      verify: 'Verify: the chart frame spans E6 to the top of J' + (last + 1) + '; seven black columns plot C' + first + ':C' + last + '.',
      cols: [['A', MARGIN_COL], ['B', MODULE], ['C', MODULE], ['D', MARGIN_COL], ['E', MODULE], ['F', MODULE], ['G', MODULE], ['H', MODULE], ['I', MODULE]],
    });
    s.set(`B${FIRST_ROW}`, 'ISO 216', { font: 'captionBold' });
    s.set(`C${FIRST_ROW}`, 'Width mm', { font: 'captionBold', align: { horizontal: 'right' } });
    ROWS.forEach(([name, w], i) => s.set(`B${first + i}`, name).set(`C${first + i}`, w));
    const chart = columnChart({
      seriesName: 'Width mm',
      cats: ROWS.map(([n]) => n),
      vals: ROWS.map(([, w]) => w),
      max: 900,
      majorUnit: 300,
      labelSz: XLSX.TYPE.caption * 100,
      refs: { name: `'${SHEET}'!$C$${FIRST_ROW}`, cat: `'${SHEET}'!$B$${first}:$B$${last}`, val: `'${SHEET}'!$C$${first}:$C$${last}` },
    });
    s.after('<drawing r:id="rId1"/>');
    s.attach({
      parts: {
        'xl/drawings/drawing1.xml': DRAWING,
        'xl/drawings/_rels/drawing1.xml.rels': relationshipsXml([['rId1', 'chart', '../charts/chart1.xml']]),
        'xl/charts/chart1.xml': chart,
      },
      sheetRels: [['rId1', 'drawing', '../drawings/drawing1.xml']],
      overrides: {
        'xl/drawings/drawing1.xml': 'application/vnd.openxmlformats-officedocument.drawing+xml',
        'xl/charts/chart1.xml': CHART_CT,
      },
    });
    return s.build();
  },
};

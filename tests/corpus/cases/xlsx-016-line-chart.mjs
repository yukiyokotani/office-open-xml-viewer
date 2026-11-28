import { XLSX } from '../lib/grid.mjs';
import { CHART_CT, lineChart } from '../lib/chart.mjs';
import { A_SERIES } from '../lib/iso216.mjs';
import { relationships } from '../lib/opc.mjs';
import { R, sheet } from '../lib/xlsx.mjs';
import { XML_DECL } from '../lib/xml.mjs';

const { MARGIN_COL, MODULE, FIRST_ROW } = XLSX;
const ROWS = A_SERIES.slice(0, 7);
const first = FIRST_ROW + 1;
const last = first + ROWS.length - 1;
const cell = (tag, col, row) => `<xdr:${tag}><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:${tag}>`;
const DRAWING = XML_DECL + '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
  + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
  + `<xdr:twoCellAnchor editAs="oneCell">${cell('from', 4, FIRST_ROW - 1)}${cell('to', 9, last + 6)}`
  + '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
  + '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
  + `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ${R} r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame>`
  + '<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>';

export default {
  id: 'xlsx-016',
  slug: 'line-chart',
  format: 'xlsx',
  title: 'Line chart',
  covers: [
    'xl/charts/chart1.xml|c:plotArea>c:lineChart',
    'xl/charts/chart1.xml|c:chart>c:legend',
    'xl/charts/chart1.xml|c:legend>c:legendPos',
    'xl/charts/chart1.xml|c:valAx>c:majorGridlines',
    'xl/charts/chart1.xml|c:dLbls>c:showVal',
    'xl/charts/chart1.xml|c:ser>c:marker',
  ],
  parts() {
    const s = sheet({
      id: 'xlsx-016',
      title: 'Line chart',
      verify: `Verify: a line through seven labelled markers over hairline gridlines, legend below, framed by E6 and J${last + 7}.`,
      cols: [['A', MARGIN_COL], ['B', MODULE], ['C', MODULE], ['D', MARGIN_COL], ['E', MODULE], ['F', MODULE], ['G', MODULE], ['H', MODULE], ['I', MODULE]],
    });
    s.set(`B${FIRST_ROW}`, 'ISO 216', { font: 'captionBold' });
    s.set(`C${FIRST_ROW}`, 'Height mm', { font: 'captionBold', align: { horizontal: 'right' } });
    ROWS.forEach(([name, , h], i) => s.set(`B${first + i}`, name).set(`C${first + i}`, h));
    const chart = lineChart({
      seriesName: 'Height mm', cats: ROWS.map(([n]) => n), vals: ROWS.map(([, , h]) => h), max: 1200, majorUnit: 400, labelSz: XLSX.TYPE.caption * 100,
    });
    s.after('<drawing r:id="rId1"/>');
    s.attach({
      parts: {
        'xl/drawings/drawing1.xml': DRAWING,
        'xl/drawings/_rels/drawing1.xml.rels': relationships([['rId1', 'chart', '../charts/chart1.xml']]),
        'xl/charts/chart1.xml': chart,
      },
      sheetRels: [['rId1', 'drawing', '../drawings/drawing1.xml']],
      overrides: { 'xl/drawings/drawing1.xml': 'application/vnd.openxmlformats-officedocument.drawing+xml', 'xl/charts/chart1.xml': CHART_CT },
    });
    return s.build();
  },
};

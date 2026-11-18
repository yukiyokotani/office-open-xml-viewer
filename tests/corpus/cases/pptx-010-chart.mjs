import { CHART_CT, columnChart } from '../lib/chart.mjs';
import { field, header, pptxPackage, textBox } from '../lib/pptx.mjs';

// ISO 216 A-series widths in millimetres.
const CATS = ['A0', 'A1', 'A2', 'A3', 'A4', 'A5', 'A6'];
const VALS = [841, 594, 420, 297, 210, 148, 105];
const box = field(1, 6, 2, 3);

const CHART = columnChart({ seriesName: 'Width, mm', cats: CATS, vals: VALS, max: 900, majorUnit: 300 });

export default {
  id: 'pptx-010',
  slug: 'chart',
  format: 'pptx',
  title: 'Chart',
  covers: [
    'ppt/slides/slide1.xml|a:graphicData>c:chart',
    'ppt/slides/slide1.xml|c:chart@r:id',
    'ppt/charts/chart1.xml|c:plotArea>c:barChart',
    'ppt/charts/chart1.xml|c:barChart>c:barDir',
    'ppt/charts/chart1.xml|c:barDir@val=col',
    'ppt/charts/chart1.xml|c:ser>c:cat',
    'ppt/charts/chart1.xml|c:cat>c:strLit',
    'ppt/charts/chart1.xml|c:val>c:numLit',
    'ppt/charts/chart1.xml|c:plotArea>c:catAx',
    'ppt/charts/chart1.xml|c:plotArea>c:valAx',
    'ppt/charts/chart1.xml|c:scaling>c:max',
    'ppt/charts/chart1.xml|c:valAx>c:majorUnit',
  ],
  parts: () => pptxPackage([
    header({
      id: 'pptx-010',
      title: 'Chart',
      verify: [
        'Verify: seven black columns, A0 to A6, over columns 1-6 of rows 2-3, heights 841 to 105 mm',
        'on a 0-900 scale with labels every 300; no title, legend or gridlines.',
      ],
    }),
    '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
      + `<p:xfrm><a:off x="${box.x}" y="${box.y}"/><a:ext cx="${box.cx}" cy="${box.cy}"/></p:xfrm>`
      + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
      + '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId2"/></a:graphicData></a:graphic></p:graphicFrame>',
    textBox(4, 'Caption', field(7, 12, 2, 2), [
      { style: 'caption', bold: true, text: 'ISO 216 A series: width in millimetres' },
      { style: 'caption', text: 'Clustered columns: c:barChart, barDir col.' },
      { style: 'caption', text: 'Literal data (c:strLit, c:numLit): no embedded workbook.' },
      { style: 'caption', text: 'Value axis: min 0, max 900, major unit 300.' },
    ]),
  ], {
    parts: { 'ppt/charts/chart1.xml': CHART },
    rels: [['rId2', 'chart', '../charts/chart1.xml']],
    overrides: { 'ppt/charts/chart1.xml': CHART_CT },
  }),
};

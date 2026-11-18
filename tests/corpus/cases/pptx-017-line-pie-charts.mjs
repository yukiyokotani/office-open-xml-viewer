import { COLOR } from '../lib/grid.mjs';
import { CHART_CT, lineChart, pieChart } from '../lib/chart.mjs';
import { field, header, pptxPackage, textBox } from '../lib/pptx.mjs';

// ISO 216 heights, and A0 cut into A1, A2, A3 and two A4 sheets.
const LINE = { seriesName: 'Height mm', cats: ['A0', 'A1', 'A2', 'A3', 'A4', 'A5', 'A6'], vals: [1189, 841, 594, 420, 297, 210, 148], max: 1200, majorUnit: 400 };
const PIE = {
  seriesName: 'Share of A0',
  cats: ['A1', 'A2', 'A3', 'A4', 'A4'],
  vals: [0.5, 0.25, 0.125, 0.0625, 0.0625],
  fills: [COLOR.ink, COLOR.accent, COLOR.ink, COLOR.accent, COLOR.ink],
};
const frame = (id, name, box, rId) => `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${name}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>`
  + `<p:xfrm><a:off x="${box.x}" y="${box.y}"/><a:ext cx="${box.cx}" cy="${box.cy}"/></p:xfrm>`
  + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
  + `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rId}"/></a:graphicData></a:graphic></p:graphicFrame>`;

export default {
  id: 'pptx-017',
  slug: 'line-pie-charts',
  format: 'pptx',
  title: 'Line and pie charts',
  covers: [
    'ppt/charts/chart1.xml|c:plotArea>c:lineChart',
    'ppt/charts/chart1.xml|c:ser>c:marker',
    'ppt/charts/chart1.xml|c:marker>c:symbol',
    'ppt/charts/chart1.xml|c:ser>c:dLbls',
    'ppt/charts/chart1.xml|c:dLbls>c:showVal',
    'ppt/charts/chart1.xml|c:valAx>c:majorGridlines',
    'ppt/charts/chart1.xml|c:chart>c:legend',
    'ppt/charts/chart1.xml|c:legend>c:legendPos',
    'ppt/charts/chart1.xml|c:legend>c:overlay',
    'ppt/charts/chart2.xml|c:plotArea>c:pieChart',
    'ppt/charts/chart2.xml|c:ser>c:dPt',
    'ppt/charts/chart2.xml|c:dLbls>c:showPercent',
    'ppt/charts/chart2.xml|c:pieChart>c:firstSliceAng',
  ],
  parts: () => pptxPackage([
    header({
      id: 'pptx-017',
      title: 'Line and pie charts',
      verify: [
        'Verify: a line through seven orange markers labelled 1189 to 148 over hairline gridlines, legend below;',
        'a pie of 50, 25, 12.5 and two 6.25 % slices from 12 o’clock, alternating black and orange, legend right.',
      ],
    }),
    frame(3, 'Line chart', field(1, 6, 2, 3), 'rId2'),
    frame(4, 'Pie chart', field(7, 12, 2, 3), 'rId3'),
    textBox(5, 'Caption', field(1, 6, 4, 4), [
      { style: 'caption', bold: true, text: 'ISO 216 A series: height in millimetres' },
      { style: 'caption', text: 'c:lineChart with c:marker, c:dLbls showVal, c:majorGridlines, c:legend legendPos b.' },
    ]),
    textBox(6, 'Caption', field(7, 12, 4, 4), [
      { style: 'caption', bold: true, text: 'A0 cut into A1, A2, A3 and two A4 sheets' },
      { style: 'caption', text: 'c:pieChart with one c:dPt fill per slice, c:dLbls showPercent, legendPos r.' },
    ]),
  ], {
    parts: { 'ppt/charts/chart1.xml': lineChart(LINE), 'ppt/charts/chart2.xml': pieChart(PIE) },
    rels: [['rId2', 'chart', '../charts/chart1.xml'], ['rId3', 'chart', '../charts/chart2.xml']],
    overrides: { 'ppt/charts/chart1.xml': CHART_CT, 'ppt/charts/chart2.xml': CHART_CT },
  }),
};

import { COLOR, DOCX } from '../lib/grid.mjs';
import { CHART_CT, columnChart, lineChart, pieChart } from '../lib/chart.mjs';
import { A_NS, WP_NS } from '../lib/drawing.mjs';
import { R_NS, docxPackage, header, run } from '../lib/docx.mjs';
import { A_SERIES } from '../lib/iso216.mjs';

// Charts embedded in a Word document: chart parts related from
// word/document.xml and placed inline, one grid row high. The column chart
// carries sheet references with cached values, as Word writes them; the other
// two carry literal data.
const at = DOCX.rowLine;
const { emu, span, LINE } = DOCX;
const SIZES = A_SERIES.slice(0, 7);
const LABEL = DOCX.TYPE.caption.sz * 50; // half-points → hundredths of a point

const inlineChart = (id, name, rId, cx, cy) => `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">`
  + `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="${name}"/>`
  + '<wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
  + `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rId}"/>`
  + '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
const HALF = { cx: emu(span(1, 2)), cy: emu(9 * LINE) };

export default {
  id: 'docx-020',
  slug: 'charts',
  format: 'docx',
  title: 'Charts',
  covers: [
    'word/document.xml|wp:inline>a:graphic',
    'word/document.xml|a:graphicData>c:chart',
    'word/document.xml|wp:inline>wp:cNvGraphicFramePr',
    'word/charts/chart1.xml|c:plotArea>c:barChart',
    'word/charts/chart2.xml|c:plotArea>c:lineChart',
    'word/charts/chart3.xml|c:plotArea>c:pieChart',
    'word/charts/chart1.xml|c:val>c:numRef',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-020',
      title: 'Charts',
      verify: [
        'Verify: row 2 holds a black column chart (columns 1-2)',
        'and a line chart with markers (columns 3-4); row 3 a pie',
        'in columns 1-2 with its legend. Labels are 8 pt.',
      ],
    }),
    {
      kind: 'p',
      at: at(2),
      lineAuto: true,
      height: 9,
      tabs: [['left', DOCX.colStart(3)]],
      runs: [inlineChart(1, 'Column chart', 'rId3', HALF.cx, HALF.cy), '<w:r><w:tab/></w:r>', inlineChart(2, 'Line chart', 'rId4', HALF.cx, HALF.cy)],
    },
    {
      kind: 'p',
      at: at(3),
      lineAuto: true,
      height: 9,
      runs: [inlineChart(3, 'Pie chart', 'rId5', HALF.cx, HALF.cy)],
    },
    {
      kind: 'p',
      at: at(4),
      style: 'caption',
      cols: [1, 2],
      runs: [run('Charts in WordprocessingML. ', '<w:b/>'), run('Each chart is a c:chartSpace part related from the '
        + 'document and drawn in a wp:inline frame of two grid columns by nine baselines. The column chart reads '
        + 'its values from c:numRef caches (Sheet1!$B$2:$B$8), as Word writes them with an embedded workbook; '
        + 'the line and pie charts carry literal values.')],
      height: 5,
    },
  ], {
    ns: [R_NS, WP_NS, A_NS],
    parts: {
      'word/charts/chart1.xml': columnChart({
        seriesName: 'Height mm', cats: SIZES.map((s) => s[0]), vals: SIZES.map((s) => s[2]), max: 1200, majorUnit: 400, labelSz: LABEL,
        refs: { name: 'Sheet1!$B$1', cat: 'Sheet1!$A$2:$A$8', val: 'Sheet1!$B$2:$B$8' },
      }),
      'word/charts/chart2.xml': lineChart({
        seriesName: 'Width mm', cats: SIZES.map((s) => s[0]), vals: SIZES.map((s) => s[1]), max: 900, majorUnit: 300, labelSz: LABEL,
      }),
      'word/charts/chart3.xml': pieChart({
        seriesName: 'Area', cats: ['A1', 'A2', 'A3', 'A4', 'A4'], vals: [50, 25, 12.5, 6.25, 6.25], labelSz: LABEL,
        fills: [COLOR.ink, COLOR.accent, COLOR.ink, COLOR.accent, COLOR.ink],
      }),
    },
    rels: [['rId3', 'chart', 'charts/chart1.xml'], ['rId4', 'chart', 'charts/chart2.xml'], ['rId5', 'chart', 'charts/chart3.xml']],
    overrides: { 'word/charts/chart1.xml': CHART_CT, 'word/charts/chart2.xml': CHART_CT, 'word/charts/chart3.xml': CHART_CT },
  }),
};

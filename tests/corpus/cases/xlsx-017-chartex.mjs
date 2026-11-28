import { COLOR, XLSX } from '../lib/grid.mjs';
import {
  CHARTEX_CT, CHARTEX_REL, DRAWING_CT, axes, chartExDrawing, chartExPart, chartExStyleParts, dim, fill, line, seriesText, spPr, txPr, uniqueId,
} from '../lib/chartex.mjs';
import { A_SERIES, B_SERIES, C_SERIES } from '../lib/iso216.mjs';
import { relationships } from '../lib/opc.mjs';
import { sheet } from '../lib/xlsx.mjs';

const { MARGIN_COL, MODULE, FIRST_ROW, TYPE } = XLSX;
const ID = 'xlsx-017';
// Data in B-E, a margin column F, the chart frame over G-K (five modules) and
// 18 rows from row 6: every chart occupies the same whole grid field.
const COLS = [['A', MARGIN_COL], ['B', MODULE], ['C', MODULE], ['D', MODULE], ['E', MODULE], ['F', MARGIN_COL],
  ['G', MODULE], ['H', MODULE], ['I', MODULE], ['J', MODULE], ['K', MODULE]];
const FRAME = { from: [6, FIRST_ROW - 1], to: [11, FIRST_ROW - 1 + 18] };
const HEAD = FIRST_ROW;
const DATA = FIRST_ROW + 1;
const head = { font: 'captionBold' };
const headRight = { font: 'captionBold', align: { horizontal: 'right' } };

const MONTHS = [['January', 31], ['February', 28], ['March', 31], ['April', 30], ['May', 31], ['June', 30],
  ['July', 31], ['August', 31], ['September', 30], ['October', 31], ['November', 30], ['December', 31]];
const quarterOf = (i) => `Q${Math.floor(i / 3) + 1}`;
const halfOf = (i) => (i < 6 ? 'H1' : 'H2');
const range = (name, col, first, last) => `'${name}'!$${col}$${first}:$${col}$${last}`;
const cellRef = (name, ref) => `'${name}'!$${ref.replace(/(\d+)/, '$$$1')}`;
const CAP = TYPE.caption;
const labels = (pos, vis, extra = '') => `<cx:dataLabels pos="${pos}">${txPr(CAP)}`
  + `<cx:visibility seriesName="${vis.s ? 1 : 0}" categoryName="${vis.c ? 1 : 0}" value="${vis.v ? 1 : 0}"/>${extra}</cx:dataLabels>`;

/** One sheet per chart type. Each returns { name, title, verify, requires,
 * fill(api), chart(name) }. The first is sheet 1, named after the case. */
const SHEETS = [
  {
    title: 'Waterfall',
    verify: 'Verify: four rising quarters, a total bar at 366, a cyan fall of 1, a total bar at 365; labels above the bars.',
    rows: [['Jan–Mar', 91], ['Apr–Jun', 91], ['Jul–Sep', 92], ['Oct–Dec', 92], ['Leap year', 366], ['29 February', -1], ['Common year', 365]],
    fill(s) {
      s.set(`B${HEAD}`, 'Period', head).set(`C${HEAD}`, 'Days', headRight);
      this.rows.forEach(([k, v], i) => s.set(`B${DATA + i}`, k).set(`C${DATA + i}`, v));
    },
    chart() {
      // The cache is the data (no formula), as Office writes a literal series.
      const data = dim('str', 'cat', { levels: [this.rows.map(([k]) => k)] })
        + dim('num', 'val', { levels: [this.rows.map(([, v]) => v)], formatCode: 'General' });
      const series = `<cx:series layoutId="waterfall" uniqueId="${uniqueId(1)}"><cx:tx><cx:txData><cx:v>Days</cx:v></cx:txData></cx:tx>`
        + `${spPr(fill(COLOR.ink))}<cx:dataPt idx="5">${spPr(fill(COLOR.accent))}</cx:dataPt>`
        + `${labels('outEnd', { v: true })}<cx:dataId val="0"/>`
        + '<cx:layoutPr><cx:subtotals><cx:idx val="4"/><cx:idx val="6"/></cx:subtotals></cx:layoutPr></cx:series>';
      return chartExPart({
        data: [[0, data]], titleText: 'Days in a leap year', titlePt: TYPE.text, textPt: CAP, plot: series, legend: 't',
        axesXml: axes({ cat: { gapWidth: 0.5 }, val: { min: 0, max: 500, majorUnit: 100, gridlines: true } }),
      });
    },
  },
  {
    name: 'Pareto',
    title: 'Pareto',
    verify: 'Verify: three columns sorted 217, 120, 28 (months summed by length); the cyan line climbs to 100 % on the right axis.',
    fill(s) {
      s.set(`B${HEAD}`, 'Month', head).set(`C${HEAD}`, 'Length', head).set(`D${HEAD}`, 'Days', headRight);
      MONTHS.forEach(([m, d], i) => s.set(`B${DATA + i}`, m).set(`C${DATA + i}`, `${d} days`).set(`D${DATA + i}`, d));
    },
    chart(name) {
      const last = DATA + MONTHS.length - 1;
      // Formula-only dimensions: the consumer reads the cells.
      const data = dim('str', 'cat', { f: range(name, 'C', DATA, last) }) + dim('num', 'val', { f: range(name, 'D', DATA, last) });
      const series = `<cx:series layoutId="clusteredColumn" uniqueId="${uniqueId(2)}">${seriesText(cellRef(name, `D${HEAD}`), 'Days')}`
        + `${spPr(fill(COLOR.ink))}${labels('outEnd', { v: true })}<cx:dataId val="0"/><cx:layoutPr><cx:aggregation/></cx:layoutPr><cx:axisId val="1"/></cx:series>`
        + `<cx:series layoutId="paretoLine" ownerIdx="0" uniqueId="${uniqueId(3)}">${spPr(line(COLOR.accent, 19050))}<cx:axisId val="2"/></cx:series>`;
      return chartExPart({
        data: [[0, data]], titleText: 'Days by month length', titlePt: TYPE.text, textPt: CAP, plot: series,
        axesXml: axes({
          cat: { gapWidth: 0 },
          val: { min: 0, max: 250, majorUnit: 50, gridlines: true },
          val2: { min: 0, max: 1, units: 'percentage', numFmt: '0%' },
        }),
      });
    },
  },
  {
    name: 'Box and whisker',
    title: 'Box and whisker',
    verify: 'Verify: three boxes of the A, B and C widths (mm); A exclusive quartiles with a mean line, B cyan, C unfilled; × marks the means.',
    fill(s) {
      s.set(`B${HEAD}`, 'Size', head).set(`C${HEAD}`, 'A width', headRight).set(`D${HEAD}`, 'B width', headRight).set(`E${HEAD}`, 'C width', headRight);
      A_SERIES.forEach((_, i) => s.set(`B${DATA + i}`, String(i)).set(`C${DATA + i}`, A_SERIES[i][1])
        .set(`D${DATA + i}`, B_SERIES[i][1]).set(`E${DATA + i}`, C_SERIES[i][1]));
    },
    chart(name) {
      const last = DATA + A_SERIES.length - 1;
      const cols = ['C', 'D', 'E'];
      const data = cols.map((c, i) => [i, dim('num', 'val', { f: range(name, c, DATA, last) })]);
      const looks = [
        { fill: fill(COLOR.paper), vis: '<cx:visibility meanLine="1" meanMarker="1" nonoutliers="0" outliers="1"/>', q: 'exclusive' },
        { fill: fill(COLOR.accent), vis: '<cx:visibility meanLine="0" meanMarker="1" nonoutliers="0" outliers="1"/>', q: 'inclusive' },
        { fill: '<a:noFill/>', vis: '<cx:visibility meanLine="0" meanMarker="1" nonoutliers="1" outliers="1"/>', q: 'inclusive' },
      ];
      const series = cols.map((c, i) => `<cx:series layoutId="boxWhisker" uniqueId="${uniqueId(4 + i)}">`
        + `${seriesText(cellRef(name, `${c}${HEAD}`), ['A width', 'B width', 'C width'][i])}${spPr(looks[i].fill + line(COLOR.ink))}`
        + `<cx:dataId val="${i}"/><cx:layoutPr>${looks[i].vis}<cx:statistics quartileMethod="${looks[i].q}"/></cx:layoutPr></cx:series>`).join('');
      return chartExPart({
        data, titleText: 'ISO 216 and ISO 269 widths', titlePt: TYPE.text, textPt: CAP, plot: series, legend: 't',
        axesXml: axes({ cat: { gapWidth: 1 }, val: { min: 0, max: 1000, majorUnit: 250, gridlines: true, ticks: 'cross' } }),
      });
    },
  },
  {
    name: 'Treemap',
    title: 'Treemap',
    verify: 'Verify: four quarter groups with banner captions, each split into three month tiles by days; tiles show month and days.',
    fill(s) {
      s.set(`B${HEAD}`, 'Quarter', head).set(`C${HEAD}`, 'Month', head).set(`D${HEAD}`, 'Days', headRight);
      MONTHS.forEach(([m, d], i) => s.set(`B${DATA + i}`, quarterOf(i)).set(`C${DATA + i}`, m).set(`D${DATA + i}`, d));
    },
    chart(name) {
      const last = DATA + MONTHS.length - 1;
      const data = dim('str', 'cat', { f: `'${name}'!$B$${DATA}:$C$${last}` }) + dim('num', 'size', { f: range(name, 'D', DATA, last) });
      const series = `<cx:series layoutId="treemap" uniqueId="${uniqueId(7)}">${seriesText(cellRef(name, `D${HEAD}`), 'Days')}`
        + `${spPr(fill(COLOR.accent) + line(COLOR.paper, 12700))}`
        + `${labels('inEnd', { c: true, v: true }, '<cx:separator>, </cx:separator>')}<cx:dataId val="0"/>`
        + '<cx:layoutPr><cx:parentLabelLayout val="banner"/></cx:layoutPr></cx:series>';
      return chartExPart({ data: [[0, data]], titleText: 'Days by quarter and month', titlePt: TYPE.text, textPt: CAP, plot: series });
    },
  },
  {
    name: 'Sunburst',
    title: 'Sunburst',
    verify: 'Verify: an inner ring of two halves, a middle ring of four quarters, an outer ring of twelve months, angles by days.',
    fill(s) {
      s.set(`B${HEAD}`, 'Half', head).set(`C${HEAD}`, 'Quarter', head).set(`D${HEAD}`, 'Month', head).set(`E${HEAD}`, 'Days', headRight);
      MONTHS.forEach(([m, d], i) => s.set(`B${DATA + i}`, halfOf(i)).set(`C${DATA + i}`, quarterOf(i)).set(`D${DATA + i}`, m).set(`E${DATA + i}`, d));
    },
    chart(name) {
      const last = DATA + MONTHS.length - 1;
      const data = dim('str', 'cat', { f: `'${name}'!$B$${DATA}:$D$${last}` }) + dim('num', 'size', { f: range(name, 'E', DATA, last) });
      const series = `<cx:series layoutId="sunburst" uniqueId="${uniqueId(8)}">${seriesText(cellRef(name, `E${HEAD}`), 'Days')}`
        + `${spPr(fill(COLOR.accent) + line(COLOR.paper, 12700))}${labels('ctr', { c: true })}<cx:dataId val="0"/></cx:series>`;
      return chartExPart({ data: [[0, data]], titleText: 'Days by half, quarter and month', titlePt: TYPE.text, textPt: CAP, plot: series });
    },
  },
  {
    name: 'Funnel',
    title: 'Funnel',
    verify: 'Verify: seven centred cyan bars, A0 841 at the top to A6 105, each half as wide every second step; values inside.',
    requires: 'cx2',
    fill(s) {
      s.set(`B${HEAD}`, 'ISO 216', head).set(`C${HEAD}`, 'Width mm', headRight);
      A_SERIES.slice(0, 7).forEach(([k, w], i) => s.set(`B${DATA + i}`, k).set(`C${DATA + i}`, w));
    },
    chart(name) {
      const rows = A_SERIES.slice(0, 7);
      const last = DATA + rows.length - 1;
      // Formula and cache together, as Office writes a series it can refresh.
      const data = dim('str', 'cat', { f: range(name, 'B', DATA, last), levels: [rows.map(([k]) => k)] })
        + dim('num', 'val', { f: range(name, 'C', DATA, last), levels: [rows.map(([, w]) => w)], formatCode: 'General' });
      const series = `<cx:series layoutId="funnel" uniqueId="${uniqueId(9)}">${seriesText(cellRef(name, `C${HEAD}`), 'Width mm')}`
        + `${spPr(fill(COLOR.accent))}${labels('ctr', { v: true })}<cx:dataId val="0"/></cx:series>`;
      return chartExPart({
        data: [[0, data]], titleText: 'ISO 216 A widths', titlePt: TYPE.text, textPt: CAP, plot: series,
        axesXml: axes({ cat: { gapWidth: 0.25 } }),
      });
    },
  },
  {
    name: 'Region map',
    title: 'Region map',
    verify: 'Verify: a world map; eight countries shaded from white (code 1) to cyan (code 91) by country calling code.',
    requires: 'cx4',
    rows: [['Canada', 1], ['Egypt', 20], ['France', 33], ['Mexico', 52], ['Brazil', 55], ['Australia', 61], ['Japan', 81], ['India', 91]],
    fill(s) {
      s.set(`B${HEAD}`, 'Country', head).set(`C${HEAD}`, 'Calling code', headRight);
      this.rows.forEach(([k, v], i) => s.set(`B${DATA + i}`, k).set(`C${DATA + i}`, v));
    },
    chart(name) {
      const last = DATA + this.rows.length - 1;
      const data = dim('str', 'cat', { f: range(name, 'B', DATA, last) }) + dim('num', 'colorVal', { f: range(name, 'C', DATA, last) });
      const series = `<cx:series layoutId="regionMap" uniqueId="${uniqueId(10)}">${seriesText(cellRef(name, `C${HEAD}`), 'Calling code')}`
        + `${spPr(line(COLOR.ink, 3175))}<cx:valueColors><cx:minColor>${`<a:srgbClr val="${COLOR.paper}"/>`}</cx:minColor>`
        + `<cx:maxColor><a:srgbClr val="${COLOR.accent}"/></cx:maxColor></cx:valueColors>`
        + '<cx:valueColorPositions count="2"><cx:min><cx:extremeValue/></cx:min><cx:max><cx:extremeValue/></cx:max></cx:valueColorPositions>'
        + '<cx:dataId val="0"/><cx:layoutPr><cx:regionLabelLayout val="none"/></cx:layoutPr></cx:series>';
      return chartExPart({ data: [[0, data]], titleText: 'Country calling codes', titlePt: TYPE.text, textPt: CAP, plot: series, legend: 'r' });
    },
  },
];

export default {
  id: ID,
  slug: 'chartex',
  format: 'xlsx',
  title: 'Chart extensions',
  covers: [
    'xl/drawings/drawing1.xml|xdr:twoCellAnchor>mc:AlternateContent',
    'xl/drawings/drawing1.xml|a:graphicData>cx:chart',
    'xl/drawings/drawing1.xml|mc:AlternateContent>mc:Fallback',
    'xl/charts/chartEx1.xml|cx:chartSpace>cx:chartData',
    'xl/charts/chartEx1.xml|cx:chartSpace>cx:txPr',
    'xl/charts/chartEx1.xml|cx:title>cx:tx',
    'xl/charts/chartEx1.xml|cx:tx>cx:rich',
    'xl/charts/chartEx1.xml|cx:strDim>cx:lvl',
    'xl/charts/chartEx1.xml|cx:series@layoutId=waterfall',
    'xl/charts/chartEx1.xml|cx:series>cx:dataPt',
    'xl/charts/chartEx1.xml|cx:dataLabels@pos=outEnd',
    'xl/charts/chartEx1.xml|cx:dataLabels>cx:visibility',
    'xl/charts/chartEx1.xml|cx:layoutPr>cx:subtotals',
    'xl/charts/chartEx1.xml|cx:legend@pos=t',
    'xl/charts/chartEx1.xml|cx:axis>cx:majorGridlines',
    'xl/charts/chartEx2.xml|cx:series@layoutId=clusteredColumn',
    'xl/charts/chartEx2.xml|cx:series@layoutId=paretoLine',
    'xl/charts/chartEx2.xml|cx:series@ownerIdx',
    'xl/charts/chartEx2.xml|cx:layoutPr>cx:aggregation',
    'xl/charts/chartEx2.xml|cx:units@unit=percentage',
    'xl/charts/chartEx2.xml|cx:strDim>cx:f',
    'xl/charts/chartEx2.xml|cx:txData>cx:f',
    'xl/charts/chartEx3.xml|cx:series@layoutId=boxWhisker',
    'xl/charts/chartEx3.xml|cx:statistics@quartileMethod=exclusive',
    'xl/charts/chartEx3.xml|cx:statistics@quartileMethod=inclusive',
    'xl/charts/chartEx3.xml|cx:visibility@meanLine',
    'xl/charts/chartEx3.xml|cx:visibility@outliers',
    'xl/charts/chartEx3.xml|cx:majorTickMarks@type=cross',
    'xl/charts/chartEx4.xml|cx:series@layoutId=treemap',
    'xl/charts/chartEx4.xml|cx:numDim@type=size',
    'xl/charts/chartEx4.xml|cx:layoutPr>cx:parentLabelLayout',
    'xl/charts/chartEx4.xml|cx:dataLabels>cx:separator',
    'xl/charts/chartEx5.xml|cx:series@layoutId=sunburst',
    'xl/charts/chartEx5.xml|cx:dataLabels@pos=ctr',
    'xl/charts/chartEx6.xml|cx:series@layoutId=funnel',
    'xl/charts/chartEx6.xml|cx:numDim>cx:lvl',
    'xl/charts/chartEx7.xml|cx:series@layoutId=regionMap',
    'xl/charts/chartEx7.xml|cx:numDim@type=colorVal',
    'xl/charts/chartEx7.xml|cx:series>cx:valueColors',
    'xl/charts/chartEx7.xml|cx:layoutPr>cx:regionLabelLayout',
    'xl/charts/chartEx7.xml|cx:legend@pos=r',
  ],
  parts() {
    const first = SHEETS[0];
    const s = sheet({ id: ID, title: 'Chart extensions: waterfall', verify: first.verify, cols: COLS });
    const parts = {};
    const overrides = {};
    SHEETS.forEach((spec, i) => {
      const n = i + 1;
      const name = i === 0 ? ID : spec.name;
      const api = i === 0 ? s : s.addSheet({ name, cols: COLS });
      if (i > 0) {
        api.set('B2', `${ID}, sheet ${n}`, { font: 'captionBold' })
          .set('B3', `Chart extensions: ${spec.title.toLowerCase()}`, { font: 'title' }).height(3, XLSX.TITLE_ROW_PT)
          .set('B4', spec.verify, { font: 'caption' });
      }
      spec.fill(api);
      api.after('<drawing r:id="rId1"/>');
      const rel = [['rId1', 'drawing', `../drawings/drawing${n}.xml`]];
      if (i === 0) s.attach({ sheetRels: rel }); else api.rels(rel);
      parts[`xl/drawings/drawing${n}.xml`] = chartExDrawing({ ...FRAME, requires: spec.requires ?? 'cx1', name: spec.title });
      parts[`xl/drawings/_rels/drawing${n}.xml.rels`] = relationships([['rId1', CHARTEX_REL, `../charts/chartEx${n}.xml`]]);
      parts[`xl/charts/chartEx${n}.xml`] = spec.chart(name);
      const style = chartExStyleParts('xl/charts', n);
      Object.assign(parts, style.parts);
      Object.assign(overrides, style.overrides);
      overrides[`xl/drawings/drawing${n}.xml`] = DRAWING_CT;
      overrides[`xl/charts/chartEx${n}.xml`] = CHARTEX_CT;
    });
    s.attach({ parts, overrides });
    return s.build();
  },
};

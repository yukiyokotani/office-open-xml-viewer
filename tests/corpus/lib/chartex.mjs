// Chart extensions ([MS-ODRAWXML] §2.24, namespace cx 2014/chartex): the
// chartEx part, the spreadsheet drawing that anchors it and the package
// scaffolding. Text is the format's caption type in the corpus font, black on
// white; a case passes the series, axes and layout properties of its chart
// type as XML fragments built with the helpers below.

import { COLOR, FONT, LANG, OUTLINED_INSET } from './grid.mjs';
import { relationships } from './opc.mjs';
import { XML_DECL, esc } from './xml.mjs';

export const CHARTEX_NS = 'http://schemas.microsoft.com/office/drawing/2014/chartex';
export const CHARTEX_CT = 'application/vnd.ms-office.chartex+xml';
export const CHARTEX_REL = 'http://schemas.microsoft.com/office/2014/relationships/chartEx';
export const CHART_STYLE_CT = 'application/vnd.ms-office.chartstyle+xml';
export const CHART_COLORS_CT = 'application/vnd.ms-office.chartcolorstyle+xml';
const CHART_STYLE_REL = 'http://schemas.microsoft.com/office/2011/relationships/chartStyle';
const CHART_COLORS_REL = 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle';
export const DRAWING_CT = 'application/vnd.openxmlformats-officedocument.drawing+xml';

/** The markup-compatibility namespace a consumer must understand for each
 * layout: Office writes waterfall, pareto, box and whisker, treemap, sunburst
 * and histogram under cx1, funnel under cx2 and region map under cx4. */
const REQUIRES = {
  cx1: 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex',
  cx2: 'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex',
  cx4: 'http://schemas.microsoft.com/office/drawing/2016/5/10/chartex',
};

export const fill = (rgb) => `<a:solidFill><a:srgbClr val="${rgb}"/></a:solidFill>`;
export const line = (rgb, w = 6350) => `<a:ln w="${w}">${fill(rgb)}</a:ln>`;
export const spPr = (inner) => `<cx:spPr>${inner}</cx:spPr>`;

/** Text properties: size in points, bold, colour. */
export const txPr = (pt, { b = false, color = COLOR.ink } = {}) => '<cx:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>'
  + `<a:defRPr sz="${pt * 100}"${b ? ' b="1"' : ''}>${fill(color)}<a:latin typeface="${FONT.latin}"/></a:defRPr>`
  + `</a:pPr><a:endParaRPr lang="${LANG.latin}"/></a:p></cx:txPr>`;

/** Chart title as rich text at the top, left-aligned with the plot. */
export const title = (text, pt) => '<cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:rich><a:bodyPr/><a:lstStyle/>'
  + `<a:p><a:pPr><a:defRPr sz="${pt * 100}" b="1"/></a:pPr><a:r><a:rPr lang="${LANG.latin}" sz="${pt * 100}" b="1">${fill(COLOR.ink)}`
  + `<a:latin typeface="${FONT.latin}"/></a:rPr><a:t>${esc(text)}</a:t></a:r></a:p></cx:rich></cx:tx></cx:title>`;

/** Series name read from a cell, with its cached text. */
export const seriesText = (f, v) => `<cx:tx><cx:txData><cx:f>${esc(f)}</cx:f><cx:v>${esc(v)}</cx:v></cx:txData></cx:tx>`;

const pts = (list) => list.map((v, i) => `<cx:pt idx="${i}">${esc(v)}</cx:pt>`).join('');

/** One data dimension. f: the sheet formula; levels: cached values, one list
 * per level, leaf level first (omit for a formula-only dimension, which the
 * consumer reads from the cells). Without f the cache is the data. */
export function dim(kind, type, { f, levels, formatCode } = {}) {
  const tag = kind === 'str' ? 'cx:strDim' : 'cx:numDim';
  const lvls = (levels ?? []).map((list) => `<cx:lvl ptCount="${list.length}"${formatCode ? ` formatCode="${esc(formatCode)}"` : ''}>${pts(list)}</cx:lvl>`).join('');
  return `<${tag} type="${type}">${f ? `<cx:f>${esc(f)}</cx:f>` : ''}${lvls}</${tag}>`;
}

/** Category and value axes; extra XML per axis is placed after the scaling. */
export function axes({ cat, val, val2 }) {
  const out = [];
  if (cat) out.push(`<cx:axis id="0"${cat.hidden ? ' hidden="1"' : ''}><cx:catScaling${cat.gapWidth !== undefined ? ` gapWidth="${cat.gapWidth}"` : ''}/>`
    + `${cat.extra ?? ''}<cx:tickLabels/>${spPr(line(COLOR.ink))}</cx:axis>`);
  const valAxis = (id, v) => `<cx:axis id="${id}"><cx:valScaling${v.max !== undefined ? ` max="${v.max}"` : ''}${v.min !== undefined ? ` min="${v.min}"` : ''}`
    + `${v.majorUnit !== undefined ? ` majorUnit="${v.majorUnit}"` : ''}/>${v.units ? `<cx:units unit="${v.units}"/>` : ''}`
    + `${v.gridlines ? `<cx:majorGridlines>${spPr(line(COLOR.ink, 3175))}</cx:majorGridlines>` : ''}`
    + `${v.ticks ? `<cx:majorTickMarks type="${v.ticks}"/>` : ''}<cx:tickLabels/>`
    + `${v.numFmt ? `<cx:numFmt formatCode="${esc(v.numFmt)}" sourceLinked="0"/>` : ''}</cx:axis>`;
  if (val) out.push(valAxis(1, val));
  if (val2) out.push(valAxis(2, val2));
  return out.join('');
}

/** A whole chartEx part. data: [[id, dims XML], ...]; plot: plotAreaRegion
 * content (series); axesXml: axes(); legend: position or undefined. */
export function chartExPart({ data, titleText, titlePt, plot, axesXml = '', legend, textPt }) {
  return XML_DECL + `<cx:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" `
    + `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:cx="${CHARTEX_NS}">`
    + `<cx:chartData>${data.map(([id, dims]) => `<cx:data id="${id}">${dims}</cx:data>`).join('')}</cx:chartData>`
    + `<cx:chart>${title(titleText, titlePt)}<cx:plotArea><cx:plotAreaRegion>${plot}</cx:plotAreaRegion>${axesXml}</cx:plotArea>`
    + `${legend ? `<cx:legend pos="${legend}" align="ctr" overlay="0"/>` : ''}</cx:chart>`
    + `${spPr(`${fill(COLOR.paper)}<a:ln><a:noFill/></a:ln>`)}${txPr(textPt)}</cx:chartSpace>`;
}

/** Spreadsheet drawing with one chartEx frame anchored to cell edges
 * (zero-based column and row of the top-left and bottom-right corners). */
export function chartExDrawing({ from: [c0, r0], to: [c1, r1], requires = 'cx1', name = 'Chart' }) {
  const cell = (col, row) => `<xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff>`;
  return XML_DECL + '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
    + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    + `<xdr:twoCellAnchor editAs="oneCell"><xdr:from>${cell(c0, r0)}</xdr:from><xdr:to>${cell(c1, r1)}</xdr:to>`
    + '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
    + `<mc:Choice xmlns:${requires}="${REQUIRES[requires]}" Requires="${requires}">`
    + `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="${esc(name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>`
    + '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
    + `<a:graphic><a:graphicData uri="${CHARTEX_NS}">`
    + `<cx:chart xmlns:cx="${CHARTEX_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>`
    + '</a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice>'
    // Office always writes a fallback: a text box for consumers that do not
    // understand the Requires namespace. Its geometry comes from the anchor.
    + '<mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr>'
    + `<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${fill(COLOR.paper)}<a:ln w="6350">${fill(COLOR.ink)}</a:ln></xdr:spPr>`
    + `<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" lIns="${OUTLINED_INSET}" tIns="${OUTLINED_INSET}" rIns="${OUTLINED_INSET}" bIns="${OUTLINED_INSET}"/><a:lstStyle/><a:p><a:r><a:rPr lang="${LANG.latin}" sz="900"><a:latin typeface="${FONT.latin}"/></a:rPr>`
    + `<a:t>Fallback: this ${esc(name.toLowerCase())} chart needs a consumer that understands ${requires}.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback>`
    + '</mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>';
}

/** Deterministic series GUID: {00000000-0000-0000-0000-<n as 12 hex digits>}. */
export const uniqueId = (n) => `{00000000-0000-0000-0000-${n.toString(16).toUpperCase().padStart(12, '0')}}`;

const CS_NS = 'xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const CS_ELEMENTS = ['axisTitle', 'categoryAxis', 'chartArea', 'dataLabel', 'dataLabelCallout', 'dataPoint', 'dataPoint3D',
  'dataPointLine', 'dataPointMarker', 'dataPointWireframe', 'dataTable', 'downBar', 'dropLine', 'errorBar', 'floor',
  'gridlineMajor', 'gridlineMinor', 'hiLoLine', 'leaderLine', 'legend', 'plotArea', 'plotArea3D', 'seriesAxis', 'seriesLine',
  'title', 'trendline', 'trendlineLabel', 'upBar', 'valueAxis', 'wall'];
const CS_ENTRY = '<cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>';

/** The chart style and colour style parts of one chartEx part, with the
 * relationships from it and the content types. Excel refuses a chartEx part
 * without these relationships (it repairs the file by deleting the drawings).
 * Every entry is empty (index 0) and the colours are the theme accents, so
 * the style adds nothing to the chart's own formatting. dir: the package
 * directory of the chart parts (xl/charts, ppt/charts, word/charts); n: the
 * chartEx part number. Returns { parts, overrides } to merge into the package. */
export function chartExStyleParts(dir, n) {
  return {
    parts: {
      [`${dir}/_rels/chartEx${n}.xml.rels`]: relationships([['rId1', CHART_STYLE_REL, `style${n}.xml`], ['rId2', CHART_COLORS_REL, `colors${n}.xml`]]),
      [`${dir}/style${n}.xml`]: `${XML_DECL}<cs:chartStyle ${CS_NS} id="410">${CS_ELEMENTS.map((e) => `<cs:${e}>${CS_ENTRY}</cs:${e}>`).join('')}</cs:chartStyle>`,
      [`${dir}/colors${n}.xml`]: `${XML_DECL}<cs:colorStyle ${CS_NS} meth="cycle" id="10">`
        + `${[1, 2, 3, 4, 5, 6].map((i) => `<a:schemeClr val="accent${i}"/>`).join('')}<cs:variation/></cs:colorStyle>`,
    },
    overrides: { [`${dir}/style${n}.xml`]: CHART_STYLE_CT, [`${dir}/colors${n}.xml`]: CHART_COLORS_CT },
  };
}

// DrawingML chart part: one clustered column series, black on white, caption
// type for labels, fixed value axis. Shared by the PPTX and XLSX specimens.

import { COLOR, FONT, LANG, PPTX } from './grid.mjs';
import { XML_DECL } from './xml.mjs';

const pts = (list) => `<c:ptCount val="${list.length}"/>${list.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}`;
const textPr = (sz) => `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${sz}"><a:solidFill><a:srgbClr val="${COLOR.ink}"/></a:solidFill>`
  + `<a:latin typeface="${FONT.latin}"/></a:defRPr></a:pPr><a:endParaRPr lang="${LANG.latin}"/></a:p></c:txPr>`;
const RULE = `<c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="${COLOR.ink}"/></a:solidFill></a:ln></c:spPr>`;
const axis = (tag, id, cross, pos, extra, TEXT) => `<c:${tag}><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/>${extra.scaling ?? ''}</c:scaling>`
  + `<c:delete val="0"/><c:axPos val="${pos}"/>${extra.numFmt ?? ''}<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>`
  + `${extra.line ? RULE : '<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>'}${TEXT}<c:crossAx val="${cross}"/><c:crosses val="autoZero"/>${extra.tail ?? ''}</c:${tag}>`;

/** cats, vals: arrays; refs?: { name, cat, val } sheet formulas, in which case
 * the data are written as references with caches instead of literals.
 * labelSz: axis label size in hundredths of a point (the format's caption). */
export function columnChart({ seriesName, cats, vals, max, majorUnit, refs, labelSz = PPTX.TYPE.caption.sz }) {
  const TEXT = textPr(labelSz);
  const cat = refs
    ? `<c:strRef><c:f>${refs.cat}</c:f><c:strCache>${pts(cats)}</c:strCache></c:strRef>`
    : `<c:strLit>${pts(cats)}</c:strLit>`;
  const val = refs
    ? `<c:numRef><c:f>${refs.val}</c:f><c:numCache><c:formatCode>General</c:formatCode>${pts(vals)}</c:numCache></c:numRef>`
    : `<c:numLit><c:formatCode>General</c:formatCode>${pts(vals)}</c:numLit>`;
  const tx = refs
    ? `<c:tx><c:strRef><c:f>${refs.name}</c:f><c:strCache>${pts([seriesName])}</c:strCache></c:strRef></c:tx>`
    : `<c:tx><c:v>${seriesName}</c:v></c:tx>`;
  return XML_DECL + '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
    + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    + '<c:roundedCorners val="0"/><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>'
    + '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
    + `<c:ser><c:idx val="0"/><c:order val="0"/>${tx}`
    + `<c:spPr><a:solidFill><a:srgbClr val="${COLOR.ink}"/></a:solidFill></c:spPr><c:invertIfNegative val="0"/>`
    + `<c:cat>${cat}</c:cat><c:val>${val}</c:val>`
    + '</c:ser><c:gapWidth val="50"/><c:axId val="1"/><c:axId val="2"/></c:barChart>'
    + axis('catAx', 1, 2, 'b', { line: true, tail: '<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/>' }, TEXT)
    + axis('valAx', 2, 1, 'l', {
      scaling: `<c:max val="${max}"/><c:min val="0"/>`,
      numFmt: '<c:numFmt formatCode="General" sourceLinked="0"/>',
      tail: `<c:crossBetween val="between"/><c:majorUnit val="${majorUnit}"/>`,
    }, TEXT)
    + '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>'
    + '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
    + '</c:chartSpace>';
}

export const CHART_CT = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';

const space = (plot, extra = '') => XML_DECL + '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
  + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
  + `<c:roundedCorners val="0"/><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>${plot}`
  + `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>${extra}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>`
  + '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:chartSpace>';
const lit = (cats, vals) => `<c:cat><c:strLit>${pts(cats)}</c:strLit></c:cat>`
  + `<c:val><c:numLit><c:formatCode>General</c:formatCode>${pts(vals)}</c:numLit></c:val>`;
const fill = (rgb) => `<a:solidFill><a:srgbClr val="${rgb}"/></a:solidFill>`;
const legend = (pos, TEXT) => `<c:legend><c:legendPos val="${pos}"/><c:overlay val="0"/>${TEXT}</c:legend>`;

/** Line chart with markers, value labels above the points, horizontal
 * gridlines and a legend below. */
export function lineChart({ seriesName, cats, vals, max, majorUnit, labelSz = PPTX.TYPE.caption.sz }) {
  const TEXT = textPr(labelSz);
  const plot = '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>'
    + `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>${seriesName}</c:v></c:tx>`
    + `<c:spPr><a:ln w="19050" cap="rnd">${fill(COLOR.ink)}<a:round/></a:ln></c:spPr>`
    + `<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr>${fill(COLOR.accent)}<a:ln><a:noFill/></a:ln></c:spPr></c:marker>`
    + `<c:dLbls><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${TEXT}<c:dLblPos val="t"/><c:showLegendKey val="0"/><c:showVal val="1"/>`
    + '<c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>'
    + `${lit(cats, vals)}<c:smooth val="0"/></c:ser><c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>`
    + axis('catAx', 1, 2, 'b', { line: true, tail: '<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/>' }, TEXT)
    + axis('valAx', 2, 1, 'l', {
      scaling: `<c:max val="${max}"/><c:min val="0"/>`,
      numFmt: '<c:numFmt formatCode="General" sourceLinked="0"/>',
      tail: `<c:crossBetween val="between"/><c:majorUnit val="${majorUnit}"/>`,
    }, TEXT).replace('<c:axPos val="l"/>', `<c:axPos val="l"/><c:majorGridlines><c:spPr><a:ln w="3175">${fill(COLOR.ink)}</a:ln></c:spPr></c:majorGridlines>`);
  return space(plot, legend('b', TEXT));
}

/** Pie chart: one fill per point (c:dPt), percentage labels, legend right. */
export function pieChart({ seriesName, cats, vals, fills, labelSz = PPTX.TYPE.caption.sz }) {
  const TEXT = textPr(labelSz);
  const dPts = fills.map((rgb, i) => `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/><c:spPr>${fill(rgb)}`
    + `<a:ln w="12700">${fill(COLOR.paper)}</a:ln></c:spPr></c:dPt>`).join('');
  const plot = '<c:pieChart><c:varyColors val="1"/>'
    + `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>${seriesName}</c:v></c:tx>${dPts}`
    + `<c:dLbls><c:numFmt formatCode="0.00%" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${TEXT}<c:dLblPos val="outEnd"/>`
    + '<c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/><c:showLeaderLines val="0"/></c:dLbls>'
    + `${lit(cats, vals)}</c:ser><c:firstSliceAng val="0"/></c:pieChart>`;
  return space(plot, legend('r', TEXT));
}

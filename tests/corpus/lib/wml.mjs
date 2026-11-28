// Raw WordprocessingML pieces for specimens whose subject the block composer
// cannot express (exact line rules, raised objects, anchors with effect
// extents, vertical sections). Lengths: twips for WordprocessingML, EMU for
// DrawingML.

import { COLOR } from './grid.mjs';
import { png } from './png.mjs';
import { esc } from './xml.mjs';

export const WP14_NS = 'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"';
export const VML_NS = ['xmlns:v="urn:schemas-microsoft-com:vml"', 'xmlns:o="urn:schemas-microsoft-com:office:office"'];

export const t = (text) => `<w:t xml:space="preserve">${esc(text)}</w:t>`;
/** A text run; rPr is raw run-property XML in schema order. */
export const r = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}${t(text)}</w:r>`;
/** A paragraph from raw pPr and content XML. */
export const p = (pPr, content = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${content}</w:p>`;
/** Line spacing: rule 'auto' (line in 240ths), 'exact' or 'atLeast' (twips). */
export const spacing = (line, rule) => `<w:spacing w:line="${line}" w:lineRule="${rule}"/>`;

/** A 32 × 32 picture that shows its orientation: a frame with a filled
 * square in the top-left corner and a bar along the top edge. */
export const ORIENT_PNG = png(32, 32, [COLOR.paper, COLOR.ink], [
  [0, 0, 32, 2, 1], [0, 30, 32, 2, 1], [0, 0, 2, 32, 1], [30, 0, 2, 32, 1], [2, 2, 12, 12, 1], [14, 2, 16, 4, 1],
]);

/** pic:pic graphic; rot in 60000ths of a degree. */
export function picGraphic({ id, rId, cx, cy, rot = 0, flipH = false, flipV = false }) {
  const xfrm = `${rot ? ` rot="${rot}"` : ''}${flipH ? ' flipH="1"' : ''}${flipV ? ' flipV="1"' : ''}`;
  return '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
    + `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr>`
    + `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
    + `<pic:spPr><a:xfrm${xfrm}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`
    + '</pic:pic></a:graphicData></a:graphic>';
}

/** Inline picture run; position raises (+) or lowers (−) it, in half-points. */
export function inlinePic({ position = 0, ...g }) {
  return `<w:r>${position ? `<w:rPr><w:position w:val="${position}"/></w:rPr>` : ''}<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">`
    + `<wp:extent cx="${g.cx}" cy="${g.cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${g.id}" name="Picture ${g.id}"/>`
    + `${picGraphic(g)}</wp:inline></w:drawing></w:r>`;
}

/** Inline wps rectangle run (a shape, not a picture) with the accent fill. */
export function inlineShape({ id, cx, cy, position = 0 }) {
  return `<w:r>${position ? `<w:rPr><w:position w:val="${position}"/></w:rPr>` : ''}<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>`
    + `<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>`
    + `<wp:docPr id="${id}" name="Shape ${id}"/><wp:cNvGraphicFramePr/>`
    + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/>'
    + `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`
    + `<a:solidFill><a:srgbClr val="${COLOR.accent}"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr>`
    + '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
    + '</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback/></mc:AlternateContent></w:r>';
}

/** Legacy VML picture run, size in points. */
export function vmlPic({ id, rId, wPt, hPt, position = 0 }) {
  return `<w:r>${position ? `<w:rPr><w:position w:val="${position}"/></w:rPr>` : ''}<w:pict>`
    + `<v:shape id="vml${id}" type="#_x0000_t75" style="width:${wPt}pt;height:${hPt}pt"><v:imagedata r:id="${rId}" o:title="Picture ${id}"/></v:shape></w:pict></w:r>`;
}

/** Position element. p: { from, offset } | { from, align } | { from, pct } (pct
 * in thousandths of a percent, wp14, with an offset fallback). */
export function position(axis, pos) {
  const open = `<wp:position${axis} relativeFrom="${pos.from}">`;
  const close = `</wp:position${axis}>`;
  if (pos.pct !== undefined) {
    return `<mc:AlternateContent><mc:Choice Requires="wp14">${open}<wp14:pct${axis === 'H' ? 'PosHOffset' : 'PosVOffset'}>${pos.pct}</wp14:pct${axis === 'H' ? 'PosHOffset' : 'PosVOffset'}>${close}</mc:Choice>`
      + `<mc:Fallback>${open}<wp:posOffset>${pos.fallback ?? 0}</wp:posOffset>${close}</mc:Fallback></mc:AlternateContent>`;
  }
  return `${open}${pos.align ? `<wp:align>${pos.align}</wp:align>` : `<wp:posOffset>${pos.offset}</wp:posOffset>`}${close}`;
}

/** Anchored drawing run. graphic: picGraphic(...) XML, or { shape: true } for
 * an accent rectangle. effect: [l, t, r, b] EMU. wrap: raw wrap element. */
export function anchor({
  id, h, v, cx, cy, graphic, wrap = '<wp:wrapNone/>', effect = [0, 0, 0, 0], behindDoc = false, relH = 1, layoutInCell = true,
}) {
  const [l, tp, rt, b] = effect;
  const shape = !graphic;
  const body = shape
    ? '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/>'
      + `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`
      + `<a:solidFill><a:srgbClr val="${COLOR.accent}"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr>`
      + '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
      + '</wps:wsp></a:graphicData></a:graphic>'
    : graphic;
  const drawing = '<w:drawing>'
    + `<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${relH}" behindDoc="${behindDoc ? 1 : 0}" locked="0" layoutInCell="${layoutInCell ? 1 : 0}" allowOverlap="1">`
    + `<wp:simplePos x="0" y="0"/>${position('H', h)}${position('V', v)}`
    + `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="${l}" t="${tp}" r="${rt}" b="${b}"/>${wrap}`
    + `<wp:docPr id="${id}" name="${shape ? 'Shape' : 'Picture'} ${id}"/><wp:cNvGraphicFramePr/>${body}`
    + '</wp:anchor></w:drawing>';
  return shape
    ? `<w:r><mc:AlternateContent><mc:Choice Requires="wps">${drawing}</mc:Choice><mc:Fallback/></mc:AlternateContent></w:r>`
    : `<w:r>${drawing}</w:r>`;
}

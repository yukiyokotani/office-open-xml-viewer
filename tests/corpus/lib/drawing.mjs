// DrawingML picture markup shared by the WordprocessingML specimens.

export const WP_NS = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
export const A_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
export const PIC_NS = 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';

/** a:graphic holding a picture that stretches the image over cx × cy EMU. */
export function pictureGraphic(id, name, rId, cx, cy) {
  return '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
    + `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>`
    + `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
    + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`
    + '</pic:pic></a:graphicData></a:graphic>';
}

export function inlinePicture({ id, name, rId, cx, cy }) {
  return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/>`
    + `<wp:docPr id="${id}" name="${name}"/>${pictureGraphic(id, name, rId, cx, cy)}</wp:inline></w:drawing></w:r>`;
}

/** Picture anchored to the margin at (x, y) EMU with square wrapping; distL
 * keeps wrapped text one gutter away. */
export function anchoredPicture({ id, name, rId, x, y, cx, cy, distL }) {
  return '<w:r><w:drawing>'
    + `<wp:anchor distT="0" distB="0" distL="${distL}" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="0">`
    + '<wp:simplePos x="0" y="0"/>'
    + `<wp:positionH relativeFrom="margin"><wp:posOffset>${x}</wp:posOffset></wp:positionH>`
    + `<wp:positionV relativeFrom="margin"><wp:posOffset>${y}</wp:posOffset></wp:positionV>`
    + `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>`
    + '<wp:wrapSquare wrapText="bothSides"/>'
    + `<wp:docPr id="${id}" name="${name}"/>${pictureGraphic(id, name, rId, cx, cy)}`
    + '</wp:anchor></w:drawing></w:r>';
}

export const WPS_NS = 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';
export const MC_NS = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';

/** A wps shape anchored in its paragraph's run. Lengths are EMU.
 * h / v: { from, offset } or { from, align }; wrap: raw wrap element
 * ('<wp:wrapNone/>' by default). geometry: prst name; adj: a:gd list;
 * fill: hex or null; outline: line XML or '' for none; text: txbxContent
 * body or ''; insets: [l, t, r, b]. */
export function wpsShape({
  id, x, y, h, v, cx, cy, prst = 'rect', adj = '', fill = null, outline = '', text = '', insets = [0, 0, 0, 0],
  relH = 1, behindDoc = false, layoutInCell = true, wrap = '<wp:wrapNone/>', dist = [0, 0, 0, 0], connector = false, flipV = false,
}) {
  const pos = (axis, p) => `<wp:position${axis} relativeFrom="${p.from}">${p.align ? `<wp:align>${p.align}</wp:align>` : `<wp:posOffset>${p.offset}</wp:posOffset>`}</wp:position${axis}>`;
  const [dT, dB, dL, dR] = dist;
  return '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>'
    + `<wp:anchor distT="${dT}" distB="${dB}" distL="${dL}" distR="${dR}" simplePos="0" relativeHeight="${relH}" behindDoc="${behindDoc ? 1 : 0}" locked="0" layoutInCell="${layoutInCell ? 1 : 0}" allowOverlap="1">`
    + `<wp:simplePos x="0" y="0"/>${pos('H', h ?? { from: 'margin', offset: x })}${pos('V', v ?? { from: 'paragraph', offset: y })}`
    + `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${wrap}`
    + `<wp:docPr id="${id}" name="Shape ${id}"/><wp:cNvGraphicFramePr/>`
    + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
    + (connector ? '<wps:cNvCnPr/>' : '<wps:cNvSpPr/>')
    + `<wps:spPr><a:xfrm${flipV ? ' flipV="1"' : ''}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`
    + `<a:prstGeom prst="${prst}"><a:avLst>${adj}</a:avLst></a:prstGeom>`
    + (connector ? '' : fill ? `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>` : '<a:noFill/>')
    + (outline || '<a:ln><a:noFill/></a:ln>') + '</wps:spPr>'
    + (text ? `<wps:txbx><w:txbxContent>${text}</w:txbxContent></wps:txbx>` : '')
    + `<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="${insets[0]}" tIns="${insets[1]}" rIns="${insets[2]}" bIns="${insets[3]}" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>`
    + '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback/></mc:AlternateContent></w:r>';
}

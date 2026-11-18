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

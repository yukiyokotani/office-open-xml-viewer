import { COLOR, FONT, OUTLINED_INSET, XLSX } from '../lib/grid.mjs';
import { relationships } from '../lib/opc.mjs';
import { gridPicture } from '../lib/png.mjs';
import { R, sheet } from '../lib/xlsx.mjs';
import { XML_DECL, esc } from '../lib/xml.mjs';

const { MARGIN_COL, MODULE, FIRST_ROW } = XLSX;
// Both objects are anchored to cell corners (zero offsets), so their edges
// are column and row boundaries of the sheet grid.
const cell = (tag, col, row) => `<xdr:${tag}><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:${tag}>`;
const top = FIRST_ROW; // 0-based row index of the first content row (row FIRST_ROW + 1)
const PICTURE = gridPicture({ cols: 2, rows: 1, field: 60, gutterX: 0, fieldH: 60, gutterY: 0, accent: [[1, 0]] });

const SHAPE = '<xdr:twoCellAnchor editAs="twoCell">' + cell('from', 1, top) + cell('to', 3, top + 6)
  + '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="Rectangle"/><xdr:cNvSpPr/></xdr:nvSpPr>'
  + '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
  + `<a:noFill/><a:ln w="12700"><a:solidFill><a:srgbClr val="${COLOR.ink}"/></a:solidFill></a:ln></xdr:spPr>`
  + `<xdr:txBody><a:bodyPr wrap="square" lIns="${OUTLINED_INSET}" tIns="${OUTLINED_INSET}" rIns="${OUTLINED_INSET}" bIns="${OUTLINED_INSET}" anchor="t"/><a:lstStyle/>`
  + `<a:p><a:r><a:rPr lang="en-GB" sz="${XLSX.TYPE.text * 100}"><a:solidFill><a:srgbClr val="${COLOR.ink}"/></a:solidFill><a:latin typeface="${FONT.latin}"/></a:rPr>`
  + `<a:t>${esc('xdr:sp anchored from B7 to the top left of D13: a rectangle with a text body, 1 pt outline, insets 12 pt (one gutter).')}</a:t></a:r></a:p></xdr:txBody></xdr:sp>`
  + '<xdr:clientData/></xdr:twoCellAnchor>';
const PIC = '<xdr:twoCellAnchor editAs="oneCell">' + cell('from', 4, top) + cell('to', 6, top + 6)
  + '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="3" name="Picture"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
  + `<xdr:blipFill><a:blip ${R} r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>`
  + '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>'
  + '<xdr:clientData/></xdr:twoCellAnchor>';
const DRAWING = XML_DECL + '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
  + `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${SHAPE}${PIC}</xdr:wsDr>`;

export default {
  id: 'xlsx-015',
  slug: 'drawing-shapes',
  format: 'xlsx',
  title: 'Shapes and pictures',
  covers: [
    'xl/drawings/drawing1.xml|xdr:twoCellAnchor>xdr:sp',
    'xl/drawings/drawing1.xml|xdr:sp>xdr:txBody',
    'xl/drawings/drawing1.xml|xdr:twoCellAnchor@editAs=twoCell',
    'xl/drawings/drawing1.xml|xdr:twoCellAnchor>xdr:pic',
    'xl/drawings/drawing1.xml|xdr:pic>xdr:blipFill',
    'xl/drawings/drawing1.xml|a:blip@r:embed',
  ],
  parts() {
    const s = sheet({
      id: 'xlsx-015',
      title: 'Shapes and pictures',
      verify: 'Verify: an outlined rectangle with text over B7:C12 and a two-field picture over E7:F12, both on cell edges.',
      cols: [['A', MARGIN_COL], ['B', MODULE], ['C', MODULE], ['D', MARGIN_COL], ['E', MODULE], ['F', MODULE]],
      gridLines: true,
    });
    s.set(`B${FIRST_ROW}`, 'Shape', { font: 'captionBold' });
    s.set(`E${FIRST_ROW}`, 'Picture (PNG from the builder)', { font: 'captionBold' });
    s.after('<drawing r:id="rId1"/>');
    s.attach({
      parts: {
        'xl/drawings/drawing1.xml': DRAWING,
        'xl/drawings/_rels/drawing1.xml.rels': relationships([['rId1', 'image', '../media/image1.png']]),
        'xl/media/image1.png': PICTURE.bytes,
      },
      sheetRels: [['rId1', 'drawing', '../drawings/drawing1.xml']],
      overrides: { 'xl/drawings/drawing1.xml': 'application/vnd.openxmlformats-officedocument.drawing+xml' },
      defaults: { png: 'image/png' },
    });
    return s.build();
  },
};

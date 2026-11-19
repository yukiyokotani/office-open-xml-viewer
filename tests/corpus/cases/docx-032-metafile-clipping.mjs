import { COLOR, DOCX } from '../lib/grid.mjs';
import { A_NS, PIC_NS, WP_NS } from '../lib/drawing.mjs';
import { R_NS, docxPackage, header } from '../lib/docx.mjs';
import { MM_ANISOTROPIC, NULL_BRUSH, NULL_PEN, PX_PER_MM, R, RGN_AND, RGN_COPY, emf } from '../lib/emf.mjs';

// Enhanced Metafile clipping. Every picture is one grid column by nine
// baselines and plays the same drawing: a thin outline of the frame, drawn
// before any clip, then a cyan fill of the frame crossed by two thick black
// diagonals, drawn under the clip the caption names. What stays cyan is the
// clip region. Device units are 0.1 mm (37.5 × 38.1 mm = 375 × 381 px).
const { span, gutterAfter, emu, LINE } = DOCX;
const TW_PER_MM = 1440 / 25.4;
const W = Math.round((span(1, 1) / TW_PER_MM) * PX_PER_MM);
const H = Math.round(((9 * LINE) / TW_PER_MM) * PX_PER_MM);
const CX = emu(span(1, 1));
const CY = emu(9 * LINE);

// Object slots: 1 cyan brush, 2 thick black pen, 3 thin black pen.
const OBJECTS = [R.brush(1, COLOR.accent), R.pen(2, 8, COLOR.ink), R.pen(3, 2, COLOR.ink)];
const outline = (w = W, h = H) => [R.select(3), R.select(NULL_BRUSH), R.rectangle([1, 1, w - 1, h - 1])];
const fill = (w = W, h = H, m = 0) => [R.select(NULL_PEN), R.select(1), R.rectangle([-m, -m, w + m, h + m])];
const cross = (w = W, h = H, m = 0) => [R.select(2), R.moveTo(-m, -m), R.lineTo(w + m, h + m), R.moveTo(w + m, -m), R.lineTo(-m, h + m)];
const pattern = (w, h, m) => [...fill(w, h, m), ...cross(w, h, m)];
const QX = Math.round(W / 4);
const QY = Math.round(H / 4);
const CENTRE = [QX, QY, W - QX, H - QY];

const PICTURES = [
  ['A. IntersectClipRect', 'cyan only in the centre', [...outline(), R.intersectClipRect(CENTRE), ...pattern()]],
  ['B. ExcludeClipRect', 'cyan except the centre', [...outline(), R.excludeClipRect(CENTRE), ...pattern()]],
  ['C. SelectClipPath, AND', 'cyan inside the ellipse', [...outline(), R.beginPath(), R.ellipse(CENTRE.map((v, i) => (i < 2 ? v - QX / 2 : v + QX / 2))), R.endPath(), R.selectClipPath(RGN_AND), ...pattern()]],
  ['D. ExtSelectClipRgn, COPY', 'two bands: top, bottom', [...outline(), R.extSelectClipRgn(RGN_COPY, [[0, 0, W, QY], [0, H - QY, W, H]]), ...pattern()]],
  ['E. SaveDC, RestoreDC', 'cyan centre, full cross', [...outline(), R.saveDC(), R.intersectClipRect(CENTRE), ...fill(), R.restoreDC(-1), ...cross()]],
  ['F. OffsetClipRgn', 'clip moved to bottom right', [...outline(), R.intersectClipRect([0, 0, 2 * QX, 2 * QY]), R.offsetClipRgn(W - 2 * QX, H - 2 * QY), ...pattern()]],
  ['G. Clip in logical units', 'window 1000 × 1000 units', [...outline(), R.setMapMode(MM_ANISOTROPIC), R.setWindowExt(1000, 1000), R.setViewportExt(W, H), R.intersectClipRect([250, 250, 750, 750]), ...pattern(1000, 1000)]],
  ['H. Ink beyond the frame', 'cut at the picture edges', [...pattern(W, H, 60), ...outline()]],
].map(([title, note, records], i) => ({ title, note, rId: `rId${3 + i}`, file: `media/image${i + 1}.emf`, bytes: emf({ width: W, height: H, handles: 3, records: [...OBJECTS, ...records] }) }));

// I: the drawing with no clip, cropped by a:srcRect to the bottom right
// three quarters of its frame and stretched over the same field.
const CROP = {
  title: 'I. a:srcRect crop', note: 'bottom right 3/4, stretched', rId: 'rId11', file: 'media/image9.emf',
  srcRect: '<a:srcRect l="25000" t="25000"/>',
  bytes: emf({ width: W, height: H, handles: 3, records: [...OBJECTS, ...outline(), ...pattern()] }),
};
const MEDIA = [...PICTURES, CROP];

function inline(id, { title, rId, srcRect = '' }) {
  return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${CX}" cy="${CY}"/>`
    + `<wp:docPr id="${id}" name="${title}"/>`
    + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
    + `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="${title}"/><pic:cNvPicPr/></pic:nvPicPr>`
    + `<pic:blipFill><a:blip r:embed="${rId}"/>${srcRect}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
    + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${CX}" cy="${CY}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`
    + '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
}

const colW = (c) => span(c, c) + (c < 4 ? gutterAfter(c) : 0);
const textP = (text, bold) => `<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r>${bold ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const cell = (c, content) => `<w:tc><w:tcPr><w:tcW w:w="${colW(c)}" w:type="dxa"/></w:tcPr>${content || '<w:p/>'}</w:tc>`;
const row = (items, firstId) => '<w:tr>'
  + [1, 2, 3, 4].map((c, i) => {
    const p = items[i];
    return cell(c, p ? `<w:p><w:pPr><w:spacing w:line="${LINE}" w:lineRule="auto"/></w:pPr>${inline(firstId + i, p)}</w:p>${textP(p.title, true)}${textP(p.note)}` : '');
  }).join('') + '</w:tr>';
const SPACER = `<w:tr><w:trPr><w:trHeight w:val="${LINE}" w:hRule="exact"/></w:trPr>${[1, 2, 3, 4].map((c) => cell(c)).join('')}</w:tr>`;
const TABLE = '<w:tbl><w:tblPr>'
  + `<w:tblW w:w="${DOCX.LIVE_W}" w:type="dxa"/><w:tblLayout w:type="fixed"/>`
  + '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
  + `<w:tblGrid>${[1, 2, 3, 4].map((c) => `<w:gridCol w:w="${colW(c)}"/>`).join('')}</w:tblGrid>`
  + row(PICTURES.slice(0, 4), 1) + SPACER + row(PICTURES.slice(4, 8), 5) + SPACER + row([CROP], 9)
  + '</w:tbl>';

export default {
  id: 'docx-032',
  slug: 'metafile-clipping',
  format: 'docx',
  title: 'Metafile clipping',
  covers: [
    'word/document.xml|pic:blipFill>a:srcRect',
    'word/document.xml|a:blip@r:embed',
    'word/media/|*.emf',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-032',
      title: 'Metafile clipping',
      verify: [
        'Verify: each EMF draws a thin frame outline, then a cyan fill',
        'and a black cross under the clip named below it; the cyan',
        'area is the clip region, and nothing is drawn outside a frame.',
      ],
    }),
    { kind: 'raw', at: DOCX.rowLine(2), xml: TABLE },
  ], {
    ns: [R_NS, WP_NS, A_NS, PIC_NS],
    parts: Object.fromEntries(MEDIA.map((p) => [`word/${p.file}`, p.bytes])),
    rels: MEDIA.map((p) => [p.rId, 'image', p.file]),
    defaults: { emf: 'image/x-emf' },
  }),
};

import { DOCX } from '../lib/grid.mjs';
import { A_NS, MC_NS, PIC_NS, WPS_NS, WP_NS } from '../lib/drawing.mjs';
import { R_NS, docxPackage, header } from '../lib/docx.mjs';
import { tiff } from '../lib/tiff.mjs';
import { COLOR } from '../lib/grid.mjs';
import { ORIENT_PNG, VML_NS, inlinePic, inlineShape, p, r, spacing, vmlPic } from '../lib/wml.mjs';

// Inline objects that are rotated, flipped, raised or lowered, and a
// PackBits TIFF:
//   A. A 216 × 72 pt picture rotated 90° between text: the line reserves the
//      rotated footprint (72 × 216 pt). Then the same picture rotated 180°,
//      and flipped horizontally.
//   B. A 24 pt picture raised 5 pt and lowered 15 pt (w:position) beside
//      text, under single, atLeast 24 pt, exact 24 pt and exact 48 pt line
//      spacing; then a shape and a VML picture raised the same way.
//   C. The raised and lowered picture in a table cell and in a text box,
//      on exact 24 pt lines (no clipping there).
//   D. A picture stored as a PackBits-compressed TIFF.
const { emu, LINE, span } = DOCX;
const PT = 20; // twips per point
const BIG = { cx: 2743200, cy: 914400 }; // 216 × 72 pt
const SMALL = emu(24 * PT);
const PB = tiff(64, 32, [COLOR.paper, COLOR.ink], [[0, 0, 64, 2, 1], [0, 30, 64, 2, 1], [0, 0, 2, 32, 1], [62, 0, 2, 32, 1], [8, 8, 16, 16, 1], [40, 8, 16, 16, 1]], 'packbits');
const caption = (text, extra = {}) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1, ...extra });
const raw = (xml) => ({ kind: 'raw', height: NaN, xml });
const label = (text) => raw(p(spacing(LINE, 'exact'), r(text, '<w:sz w:val="16"/>')));

let nextId = 1;
const small = (position) => inlinePic({ id: nextId++, rId: 'rId5', cx: SMALL, cy: SMALL, position });
const raisedLine = (pPr) => p(pPr, `${r('Base ')}${small(10)}${r(' up 5 pt, down 15 pt ')}${small(-30)}${r(' base.')}`);
const ruleRow = (name, pPr) => [label(name), raw(raisedLine(pPr))];

const cellTable = (content) => '<w:tbl><w:tblPr>'
  + `<w:tblW w:w="${span(1, 2)}" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders>`
  + '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
  + `<w:tblGrid><w:gridCol w:w="${span(1, 2)}"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="${span(1, 2)}" w:type="dxa"/></w:tcPr>${content}</w:tc></w:tr></w:tbl>`;

const textBox = (content) => '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>'
  + '<wp:inline distT="0" distB="0" distL="0" distR="0">'
  + `<wp:extent cx="${emu(span(1, 2))}" cy="${emu(4 * LINE)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="90" name="Text box 90"/><wp:cNvGraphicFramePr/>`
  + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>'
  + `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(span(1, 2))}" cy="${emu(4 * LINE)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>`
  + '<a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>'
  + `<wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx>`
  + '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
  + '</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback/></mc:AlternateContent></w:r>';

const parts = () => {
  nextId = 1;
  const EXACT24 = spacing(24 * PT, 'exact');
  const big = (extra) => inlinePic({ id: nextId++, rId: 'rId5', ...BIG, ...extra });
  return docxPackage([
    ...header({
      id: 'docx-052',
      title: 'Inline object position',
      verify: [
        'Verify: A 90° picture takes a 216 pt tall line; B/C',
        'pictures move 5 pt up, 15 pt down as in the PDF;',
        'D, a PackBits TIFF, shows a frame with two squares.',
      ],
    }),
    caption('A. Picture rotated 90°, 180°, flipped', { at: DOCX.rowLine(2), before: 0 }),
    raw(p(spacing(LINE, 'auto'), `${r('Before ')}${big({ rot: 5400000 })}${r(' after the 90° picture.')}`)),
    raw(p(spacing(LINE, 'auto'), `${r('Before ')}${big({ rot: 10800000 })}${r(' after 180°.')}`)),
    raw(p(spacing(LINE, 'auto'), `${r('Before ')}${big({ flipH: true })}${r(' after the flip.')}`)),
    raw(p(spacing(LINE, 'auto'), r('The next line after A.'))),
    { kind: 'p', style: 'caption', bold: true, lines: ['B. Raised and lowered under each line rule'], height: NaN, pageBreakBefore: true },
    ...ruleRow('B1 single (auto)', spacing(LINE, 'auto')),
    ...ruleRow('B2 atLeast 24 pt', spacing(24 * PT, 'atLeast')),
    ...ruleRow('B3 exact 24 pt', EXACT24),
    ...ruleRow('B4 exact 48 pt', spacing(48 * PT, 'exact')),
    label('B5 shape, exact 24 pt'),
    raw(p(EXACT24, `${r('Base ')}${inlineShape({ id: 80, cx: SMALL, cy: SMALL, position: 10 })}${r(' up, down ')}${inlineShape({ id: 81, cx: SMALL, cy: SMALL, position: -30 })}${r(' base.')}`)),
    label('B6 VML picture, exact 24 pt'),
    raw(p(EXACT24, `${r('Base ')}${vmlPic({ id: 82, rId: 'rId5', wPt: 24, hPt: 24, position: 10 })}${r(' up, down ')}${vmlPic({ id: 83, rId: 'rId5', wPt: 24, hPt: 24, position: -30 })}${r(' base.')}`)),
    caption('C. In a table cell and a text box, exact 24 pt'),
    raw(cellTable(raisedLine(EXACT24))),
    raw(p(spacing(LINE, 'auto'), textBox(raisedLine(EXACT24)))),
    caption('D. PackBits TIFF'),
    raw(p(spacing(LINE, 'auto'), inlinePic({ id: 84, rId: 'rId6', cx: emu(span(1, 1)), cy: emu(span(1, 1)) / 2 }))),
  ], {
    parts: { 'word/media/image1.png': ORIENT_PNG, 'word/media/image2.tiff': PB },
    rels: [['rId5', 'image', 'media/image1.png'], ['rId6', 'image', 'media/image2.tiff']],
    defaults: { png: 'image/png', tiff: 'image/tiff' },
    ns: [R_NS, WP_NS, A_NS, PIC_NS, WPS_NS, MC_NS, ...VML_NS],
  });
};

export default {
  id: 'docx-052',
  slug: 'inline-object-position',
  format: 'docx',
  title: 'Inline object position',
  covers: [
    'word/document.xml|a:xfrm@rot',
    'word/document.xml|a:xfrm@flipH',
    'word/document.xml|w:rPr>w:position',
    'word/document.xml|w:spacing@w:lineRule=exact',
    'word/document.xml|w:spacing@w:lineRule=atLeast',
    'word/document.xml|w:pict>v:shape',
    'word/document.xml|w:txbxContent>w:p',
    'word/media/|*.tiff',
  ],
  parts,
};

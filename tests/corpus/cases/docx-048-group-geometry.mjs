import { COLOR, DOCX } from '../lib/grid.mjs';
import { A_NS, MC_NS, WPS_NS, WP_NS } from '../lib/drawing.mjs';
import { docxPackage, header } from '../lib/docx.mjs';
import { esc } from '../lib/xml.mjs';

// Shape groups (wpg:wgp) whose child coordinates do not match the page:
//   A. A group whose child extent (a:chExt) is twice its height (a:ext), so
//      every child is drawn at half its height; a custom-geometry triangle
//      (a:custGeom) among the children; and one child placed at a negative
//      x, outside the group's bounds, which Word still draws.
//   B. A behindDoc group anchored in a table cell (layoutInCell),
//      positioned from the page's left edge and two lines above its
//      paragraph, as wide as the page, with custom-geometry corners. With
//      layoutInCell Word measures from the cell, so the band starts at the
//      table's left edge and runs past the right of the page.
//   C. One large group of twelve outlined rectangles and twelve text boxes
//      laid out as cards, with a child offset (a:chOff) that is not zero, so
//      every child position is shifted by it.
const { LINE, PAGE_W, colStart, span, emu } = DOCX;
const WPG_NS = 'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"';
const LN = `<a:ln w="12700"><a:solidFill><a:srgbClr val="${COLOR.ink}"/></a:solidFill></a:ln>`;
const NO_LN = '<a:ln><a:noFill/></a:ln>';
const fill = (c) => `<a:solidFill><a:srgbClr val="${c}"/></a:solidFill>`;
const xfrm = (x, y, cx, cy) => `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`;
const r = (text) => `<w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
const BODY = '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>';

let nextId = 1;
const rect = (x, y, cx, cy, look) => `<wps:wsp><wps:cNvPr id="${nextId}" name="Child ${nextId++}"/><wps:cNvSpPr/>`
  + `<wps:spPr>${xfrm(x, y, cx, cy)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${look}</wps:spPr>${BODY}</wps:wsp>`;
/** A right triangle with its right angle at the top left (or bottom right). */
const corner = (x, y, cx, cy, look, flip = false) => `<wps:wsp><wps:cNvPr id="${nextId}" name="Corner ${nextId++}"/><wps:cNvSpPr/>`
  + `<wps:spPr>${xfrm(x, y, cx, cy)}<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>`
  + `<a:pathLst><a:path w="${cx}" h="${cy}">`
  + (flip
    ? `<a:moveTo><a:pt x="${cx}" y="0"/></a:moveTo><a:lnTo><a:pt x="${cx}" y="${cy}"/></a:lnTo><a:lnTo><a:pt x="0" y="${cy}"/></a:lnTo>`
    : `<a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="${cx}" y="0"/></a:lnTo><a:lnTo><a:pt x="0" y="${cy}"/></a:lnTo>`)
  + `<a:close/></a:path></a:pathLst></a:custGeom>${look}</wps:spPr>${BODY}</wps:wsp>`;
const textBox = (x, y, cx, cy, text) => `<wps:wsp><wps:cNvPr id="${nextId}" name="Text ${nextId++}"/><wps:cNvSpPr txBox="1"/>`
  + `<wps:spPr>${xfrm(x, y, cx, cy)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>${NO_LN}</wps:spPr>`
  + `<wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr>${r(text)}</w:p></w:txbxContent></wps:txbx>${BODY}</wps:wsp>`;

/** The anchored group: ext and chExt in EMU, children already written. */
function group({ h, v, cx, cy, chOff = [0, 0], chExt = [cx, cy], children, behindDoc = false, layoutInCell = true }) {
  const pos = (axis, p) => `<wp:position${axis} relativeFrom="${p.from}"><wp:posOffset>${p.offset}</wp:posOffset></wp:position${axis}>`;
  return '<w:r><mc:AlternateContent><mc:Choice Requires="wpg"><w:drawing>'
    + `<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="${behindDoc ? 1 : 0}" locked="0" layoutInCell="${layoutInCell ? 1 : 0}" allowOverlap="1">`
    + `<wp:simplePos x="0" y="0"/>${pos('H', h)}${pos('V', v)}`
    + `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>`
    + `<wp:docPr id="${nextId}" name="Group ${nextId++}"/><wp:cNvGraphicFramePr/>`
    + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp><wpg:cNvGrpSpPr/>'
    + `<wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/><a:chOff x="${chOff[0]}" y="${chOff[1]}"/><a:chExt cx="${chExt[0]}" cy="${chExt[1]}"/></a:xfrm></wpg:grpSpPr>`
    + children.join('')
    + '</wpg:wgp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback/></mc:AlternateContent></w:r>';
}

/** An anchor paragraph holding `runs`, then empty lines to reserve the block. */
const block = (runs, lines) => [
  { kind: 'p', lines: [''], runs: [runs], height: 1 },
  ...(lines > 1 ? [{ kind: 'p', lines: Array(lines - 1).fill(''), height: lines - 1 }] : []),
];
const caption = (text, extra = {}) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1, ...extra });

const parts = () => {
  nextId = 1;
  // A. Group over columns 3-4, 8 lines tall; child space 16 lines tall.
  const A_W = emu(span(3, 4));
  const A_H = emu(8 * LINE);
  const c3 = emu(span(3, 3));
  const c4x = emu(colStart(4) - colStart(3));
  const c4 = emu(span(4, 4));
  const outside = emu(colStart(3) - colStart(2));
  const A = group({
    h: { from: 'margin', offset: emu(colStart(3)) }, v: { from: 'paragraph', offset: 0 },
    cx: A_W, cy: A_H, chExt: [A_W, 2 * A_H],
    children: [
      rect(0, 0, c3, 2 * A_H, fill(COLOR.accent) + LN), // full height: 8 lines
      rect(c4x, 0, c4, A_H, fill(COLOR.accent) + LN), // half of child space: 4 lines
      corner(c4x, A_H, c4, A_H, fill(COLOR.ink) + NO_LN), // bottom 4 lines
      rect(-outside, emu(4 * LINE), emu(span(2, 2)), emu(8 * LINE), fill(COLOR.ink) + NO_LN), // column 2, lines 3-6
    ],
  });

  // B. Page-wide band behind a one-cell table, anchored in the cell.
  const BAND_H = emu(4 * LINE);
  const B_GROUP = group({
    h: { from: 'page', offset: 0 }, v: { from: 'paragraph', offset: emu(-2 * LINE) },
    cx: emu(PAGE_W), cy: BAND_H, behindDoc: true,
    children: [
      rect(0, 0, emu(PAGE_W), BAND_H, fill(COLOR.accent) + NO_LN),
      corner(0, 0, emu(span(1, 1)), BAND_H, fill(COLOR.ink) + NO_LN),
      corner(emu(PAGE_W - span(1, 1)), 0, emu(span(1, 1)), BAND_H, fill(COLOR.ink) + NO_LN, true),
    ],
  });
  const CELL_W = span(1, 4);
  const B = '<w:tbl><w:tblPr>'
    + `<w:tblW w:w="${CELL_W}" w:type="dxa"/><w:tblBorders>${['top', 'left', 'bottom', 'right'].map((n) => `<w:${n} w:val="single" w:sz="4" w:space="0" w:color="${COLOR.ink}"/>`).join('')}</w:tblBorders>`
    + '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
    + `<w:tblGrid><w:gridCol w:w="${CELL_W}"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="${4 * LINE}" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="${CELL_W}" w:type="dxa"/></w:tcPr>`
    + `<w:p>${r('First cell line.')}</w:p><w:p>${r('Second line.')}</w:p>`
    + `<w:p>${B_GROUP}${r('Third line anchors the band: page left, two lines up, layoutInCell.')}</w:p>`
    + `<w:p>${r('Fourth line.')}</w:p></w:tc></w:tr></w:tbl>`;

  // C. Twelve cards in one group, child space shifted by chOff.
  const OFF = [emu(4 * LINE), emu(2 * LINE)];
  const C_W = emu(span(1, 4));
  const C_H = emu(11 * LINE);
  const cards = [];
  for (let row = 0; row < 3; row += 1) {
    for (let col = 1; col <= 4; col += 1) {
      const x = OFF[0] + emu(colStart(col));
      const y = OFF[1] + emu(row * 4 * LINE);
      const n = row * 4 + col;
      cards.push(rect(x, y, emu(span(col, col)), emu(3 * LINE), (n % 2 ? fill(COLOR.accent) : fill(COLOR.paper)) + LN));
      cards.push(textBox(x, y + emu(LINE), emu(span(col, col)), emu(LINE), `Card ${n}`));
    }
  }
  const C = group({ h: { from: 'margin', offset: 0 }, v: { from: 'paragraph', offset: 0 }, cx: C_W, cy: C_H, chOff: OFF, chExt: [C_W, C_H], children: cards });

  return docxPackage([
    ...header({
      id: 'docx-048',
      title: 'Group geometry',
      verify: [
        'Verify: A, left bar 8 lines, right box 4 lines over a',
        'triangle, black bar in column 2; B band starts at the',
        'cell edge, two lines up; C cards 1-12 on the grid.',
      ],
    }),
    caption('A. chExt twice the height; a child outside the group', { at: DOCX.rowLine(2), before: 0 }),
    ...block(A, 8),
    caption('B. Page-wide band anchored in a table cell'),
    { kind: 'p', lines: [''], height: 1 },
    { kind: 'p', lines: [''], height: 1 },
    { kind: 'raw', xml: B, height: 4 },
    { kind: 'p', lines: [''], height: 1 },
    caption('C. One group of twelve cards, child offset not zero'),
    ...block(C, 11),
  ], {
    ns: [A_NS, WP_NS, WPS_NS, WPG_NS, MC_NS],
  });
};

export default {
  id: 'docx-048',
  slug: 'group-geometry',
  format: 'docx',
  title: 'Group geometry',
  covers: [
    'word/document.xml|a:graphicData>wpg:wgp',
    'word/document.xml|a:xfrm>a:chExt',
    'word/document.xml|a:xfrm>a:chOff',
    'word/document.xml|wps:spPr>a:custGeom',
    'word/document.xml|a:path>a:lnTo',
    'word/document.xml|wps:cNvSpPr@txBox',
    'word/document.xml|wp:anchor@behindDoc',
    'word/document.xml|wp:positionH@relativeFrom=page',
    'word/document.xml|w:tc>w:p',
  ],
  parts,
};

import { DOCX } from '../lib/grid.mjs';
import { A_NS, MC_NS, WPS_NS, WP_NS } from '../lib/drawing.mjs';
import { R_NS, W, WML, docxPackage, header } from '../lib/docx.mjs';
import { XML_DECL } from '../lib/xml.mjs';
import { p, r } from '../lib/wml.mjs';

// HTML automatic spacing (w:beforeAutospacing and
// w:afterAutospacing) and a hidden paragraph in a table cell.
//   Body start: the first paragraph (the id line) and the first header
//     paragraph have automatic space before, which Word drops there.
//   A. Three paragraphs with automatic spacing between two markers.
//   B. A list whose items alternate between two w:num entries that share
//      one abstract numbering: no automatic space between the items.
//   C. A list alternating between two different abstract numberings.
//   D. Automatic spacing at the start of a table cell, E. of a text box,
//      F. on the first paragraph of page 2.
//   G. A cell whose first paragraph is hidden, mark included (it takes no
//      height), beside a cell with the visible paragraph alone.
const { emu, LINE, span } = DOCX;
const AUTO = '<w:spacing w:beforeAutospacing="1" w:afterAutospacing="1"/>';
const raw = (xml) => ({ kind: 'raw', height: NaN, xml });
const auto = (text, pPr = '') => p(`${pPr}${AUTO}`, r(text));
const marker = (text) => raw(p('', r(text)));
const caption = (text, extra = {}) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1, ...extra });
const numPr = (id) => `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${id}"/></w:numPr>`;

const IND = `<w:pPr><w:ind w:left="${DOCX.colStart(2) - DOCX.colStart(1)}" w:hanging="${DOCX.GUTTER}"/></w:pPr>`;
const lvl = (fmt) => `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>${IND}</w:lvl>`;
const NUMBERING = XML_DECL + `<w:numbering ${W}>`
  + `<w:abstractNum w:abstractNumId="1">${lvl('decimal')}</w:abstractNum>`
  + `<w:abstractNum w:abstractNumId="2">${lvl('upperLetter')}</w:abstractNum>`
  + '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>'
  + '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
  + '<w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num></w:numbering>';
const HEADER = XML_DECL + `<w:hdr ${W}>${auto('Header line with automatic space before (dropped).')}</w:hdr>`;

const BORDERS = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideV'].map((n) => `<w:${n} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('') + '</w:tblBorders>';
const table = (cells) => '<w:tbl><w:tblPr>'
  + `<w:tblW w:w="${span(1, 2 * cells.length)}" w:type="dxa"/>${BORDERS}<w:tblLayout w:type="fixed"/>`
  + '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
  + `<w:tblGrid>${cells.map((_, i) => `<w:gridCol w:w="${span(2 * i + 1, 2 * i + 2)}"/>`).join('')}</w:tblGrid>`
  + `<w:tr>${cells.map((c, i) => `<w:tc><w:tcPr><w:tcW w:w="${span(2 * i + 1, 2 * i + 2)}" w:type="dxa"/></w:tcPr>${c}</w:tc>`).join('')}</w:tr></w:tbl>`;
const HIDDEN = '<w:vanish/>';
const hiddenPara = p(`<w:rPr>${HIDDEN}</w:rPr>`, r('hidden paragraph', HIDDEN));

const textBox = (content) => '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>'
  + `<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${emu(span(1, 2))}" cy="${emu(6 * LINE)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>`
  + '<wp:docPr id="1" name="Text box 1"/><wp:cNvGraphicFramePr/>'
  + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>'
  + `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(span(1, 2))}" cy="${emu(6 * LINE)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>`
  + '<a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>'
  + `<wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx>`
  + '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
  + '</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback/></mc:AlternateContent></w:r>';

const furniture = () => {
  const blocks = header({
    id: 'docx-056',
    title: 'HTML automatic spacing',
    verify: [
      'Verify: automatic gaps are 14 pt, none at the start of',
      'the body, a cell, a text box or the header, none inside',
      'list B; G cells are both one line tall.',
    ],
  });
  blocks[0] = { ...blocks[0], pPrMid: '<w:spacing w:beforeAutospacing="1"/>' };
  return blocks;
};

export default {
  id: 'docx-056',
  slug: 'html-auto-spacing',
  format: 'docx',
  title: 'HTML automatic spacing',
  covers: [
    'word/document.xml|w:spacing@w:beforeAutospacing',
    'word/document.xml|w:spacing@w:afterAutospacing',
    'word/header1.xml|w:spacing@w:beforeAutospacing',
    'word/numbering.xml|w:num>w:abstractNumId',
    'word/document.xml|w:rPr>w:vanish',
  ],
  parts: () => docxPackage([
    ...furniture(),
    caption('A. Three paragraphs with automatic spacing', { at: DOCX.rowLine(2), before: 0 }),
    marker('▼ A starts'),
    ...['A1', 'A2', 'A3'].map((n) => raw(auto(`Paragraph ${n} with automatic space before and after.`))),
    marker('▲ A ends'),
    caption('B. One abstract numbering, two w:num entries', { height: NaN }),
    ...[1, 2, 1, 2].map((id, i) => raw(auto(`Item B${i + 1} (numId ${id})`, numPr(id)))),
    marker('▲ B ends'),
    caption('C. Two abstract numberings', { height: NaN }),
    ...[1, 3, 1, 3].map((id, i) => raw(auto(`Item C${i + 1} (numId ${id})`, numPr(id)))),
    marker('▲ C ends'),
    caption('D. Table cell start; E. text box start', { height: NaN }),
    raw(table([auto('D1 first in the cell') + auto('D2 second in the cell')])),
    raw(p('', textBox(auto('E1 first in the text box') + auto('E2 second in the text box')))),
    caption('G. Hidden first paragraph in a cell', { height: NaN }),
    raw(table([hiddenPara + p('', r('G1 visible, after a hidden one')), p('', r('G2 visible alone'))])),
    raw(auto('F. Page 2 starts with this paragraph, automatic spacing.', '<w:pageBreakBefore/>')),
    marker('▲ F ends'),
  ], {
    firstSection: { refs: '<w:headerReference w:type="default" r:id="rId4"/>', headerDist: DOCX.HEADER_DIST },
    parts: { 'word/numbering.xml': NUMBERING, 'word/header1.xml': HEADER },
    rels: [['rId3', 'numbering', 'numbering.xml'], ['rId4', 'header', 'header1.xml']],
    overrides: { 'word/numbering.xml': `${WML}.numbering+xml`, 'word/header1.xml': `${WML}.header+xml` },
    ns: [R_NS, WP_NS, A_NS, WPS_NS, MC_NS],
  }),
};

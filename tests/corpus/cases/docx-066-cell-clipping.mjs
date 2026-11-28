import { DOCX } from '../lib/grid.mjs';
import { A_NS, MC_NS, PIC_NS, WPS_NS, WP_NS } from '../lib/drawing.mjs';
import { R_NS, WML, docxPackage, header } from '../lib/docx.mjs';
import { XML_DECL } from '../lib/xml.mjs';
import { ORIENT_PNG, inlinePic, p, r, spacing } from '../lib/wml.mjs';

// Content larger than its table cell or text box. Word cuts it at the cell
// or box edge; nothing paints into the next row or cell.
//   A. Rows of exact height one baseline (w:hRule="exact") under a fixed
//      layout: A1 holds three lines of text, A2 one exact 36 pt line. Only
//      the first line of A1 shows; the row below stays clear.
//      A comment range spans the hidden second and third lines of A1.
//   B. A fixed-width cell (one grid column) holding an inline picture three
//      columns wide: the picture is cut at the cell's right edge and the
//      neighbouring cell's text stays visible.
//   C. An inline text box two baselines tall in a cell, holding four lines
//      (no autofit): the box shows two and the cell below stays clear.
const { emu, LINE, span } = DOCX;
const PT = 20;
const BORDERS = '<w:tblBorders>'
  + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((n) => `<w:${n} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('')
  + '</w:tblBorders>';
const MAR0 = '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>';
const raw = (xml) => ({ kind: 'raw', height: NaN, xml });
const caption = (text, extra = {}) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1, height: NaN, ...extra });

const W1 = span(1, 1);
const W3 = span(2, 4);
const tc = (w, content) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
const tr = (cells, trPr = '') => `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}${cells.join('')}</w:tr>`;
const table = (rows) => '<w:tbl><w:tblPr>'
  + `<w:tblW w:w="${W1 + W3}" w:type="dxa"/>${BORDERS}<w:tblLayout w:type="fixed"/>${MAR0}</w:tblPr>`
  + `<w:tblGrid><w:gridCol w:w="${W1}"/><w:gridCol w:w="${W3}"/></w:tblGrid>${rows.join('')}</w:tbl>`;
const EXACT_ROW = `<w:trHeight w:val="${LINE}" w:hRule="exact"/>`;
const line = (text) => p(spacing(LINE, 'exact'), r(text));
const BR = '<w:r><w:br/></w:r>';

// A1: three lines in one paragraph; the comment covers lines 2 and 3.
const A1 = p(spacing(LINE, 'exact'), r('A1 line 1 of 3')
  + '<w:commentRangeStart w:id="0"/>' + BR + r('A1 line 2, hidden') + BR + r('A1 line 3, hidden')
  + '<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>');
const A2 = p(spacing(36 * PT, 'exact'), r('A2 one exact 36 pt line'));
const TABLE_A = table([
  tr([tc(W1, line('A1')), tc(W3, A1)], EXACT_ROW),
  tr([tc(W1, line('A2')), tc(W3, A2)], EXACT_ROW),
  tr([tc(W1, line('Next row')), tc(W3, line('Clear: nothing from A1 or A2 here.'))]),
]);

const PIC = { cx: emu(W3), cy: emu(4 * LINE) };
const TABLE_B = table([
  tr([tc(W1, p(spacing(LINE, 'auto'), inlinePic({ id: 1, rId: 'rId5', ...PIC }))), tc(W3, line('B neighbour cell: text stays visible.'))]),
  tr([tc(W1, line('Next row')), tc(W3, line('Clear.'))]),
]);

const BOX = { cx: emu(W3), cy: emu(2 * LINE) };
const textBox = (content) => '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>'
  + '<wp:inline distT="0" distB="0" distL="0" distR="0">'
  + `<wp:extent cx="${BOX.cx}" cy="${BOX.cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="2" name="Text box 2"/><wp:cNvGraphicFramePr/>`
  + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>'
  + `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${BOX.cx}" cy="${BOX.cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>`
  + `<wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx>`
  + '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
  + '</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback/></mc:AlternateContent></w:r>';
const TABLE_C = table([
  tr([tc(W1, line('C box')), tc(W3, p(spacing(LINE, 'auto'), textBox(
    [1, 2, 3, 4].map((n) => line(n <= 2 ? `C box line ${n} of 4` : `C box line ${n}, hidden`)).join(''),
  )))]),
  tr([tc(W1, line('Next row')), tc(W3, line('Clear: no box lines here.'))]),
]);

const COMMENTS = XML_DECL + `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`
  + '<w:comment w:id="0" w:author="Corpus" w:initials="C"><w:p><w:r><w:t xml:space="preserve">'
  + 'Range over the hidden lines of A1.</w:t></w:r></w:p></w:comment></w:comments>';

export default {
  id: 'docx-066',
  slug: 'cell-clipping',
  format: 'docx',
  title: 'Cell clipping',
  covers: [
    'word/document.xml|w:trHeight@w:hRule=exact',
    'word/document.xml|w:tblLayout@w:type=fixed',
    'word/document.xml|w:tc>w:p',
    'word/document.xml|w:drawing>wp:inline',
    'word/document.xml|w:txbxContent>w:p',
    'word/document.xml|w:p>w:commentRangeStart',
    'word/document.xml|w:r>w:commentReference',
    'word/comments.xml|w:comments>w:comment',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-066',
      title: 'Cell clipping',
      verify: [
        'Verify: content past a cell or box edge is cut there;',
        'every "Next row" line and the B neighbour stay clear,',
        'no "hidden" line shows. Compare the PDF.',
      ],
    }),
    caption('A. Exact one-baseline rows, fixed layout', { at: DOCX.rowLine(2), before: 0 }),
    raw(TABLE_A),
    caption('B. Picture three columns wide in a one-column cell'),
    raw(TABLE_B),
    caption('C. Text box of two baselines holding four lines'),
    raw(TABLE_C),
    { kind: 'p', lines: ['After the tables.'], height: NaN },
  ], {
    parts: { 'word/media/image1.png': ORIENT_PNG, 'word/comments.xml': COMMENTS },
    rels: [['rId5', 'image', 'media/image1.png'], ['rId6', 'comments', 'comments.xml']],
    overrides: { 'word/comments.xml': `${WML}.comments+xml` },
    defaults: { png: 'image/png' },
    ns: [R_NS, WP_NS, A_NS, PIC_NS, WPS_NS, MC_NS],
  }),
};

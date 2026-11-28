import { COLOR, DOCX } from '../lib/grid.mjs';
import { docxPackage, header } from '../lib/docx.mjs';
import { esc } from '../lib/xml.mjs';

// Text frames (w:framePr) and a fixed-width table that continues on the
// next page, as leaflets and forms lay out boxes:
//   A. A framed paragraph two columns wide and three lines high (exact),
//      with text wrapping around it; two consecutive paragraphs with the
//      same framePr, which Word sets in one frame (aligned right); a frame
//      with wrap notBeside, so no text runs beside it; and a frame anchored
//      to the page, centred, at a fixed distance from the top margin.
//   B. A floating table (w:tblpPr, anchored to the margin and the text) of
//      fixed width, near the foot of page 1, whose rows continue on page 2.
//      Each value cell holds a nested table whose grid is wider than the
//      cell. The continued rows must start at the top of page 2 in the same
//      columns as on page 1.
//   C. The same table in the flow, not floating, on page 2.
const { LINE, GUTTER, colStart, span, lineTop } = DOCX;
const t = (text) => `<w:t xml:space="preserve">${esc(text)}</w:t>`;
const r = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}${t(text)}</w:r>`;
const BOX = `<w:pBdr>${['top', 'left', 'bottom', 'right'].map((n) => `<w:${n} w:val="single" w:sz="4" w:space="0" w:color="${COLOR.ink}"/>`).join('')}</w:pBdr>`;
const SHADE = `<w:shd w:val="clear" w:color="auto" w:fill="${COLOR.accent}"/>`;
const P = (runs, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
const raw = (xml) => ({ kind: 'raw', xml, height: NaN });
const caption = (text, extra = {}) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1, ...extra });
const FILL = 'Ordinary text continues here so the paragraph runs to several lines and shows how it flows around the frame placed beside it.';

const frame = (attrs) => `<w:framePr ${attrs}/>`;
const F1 = frame(`w:w="${span(1, 2)}" w:h="${3 * LINE}" w:hRule="exact" w:hSpace="${GUTTER}" w:wrap="around" w:vAnchor="text" w:hAnchor="margin" w:x="0" w:y="0"`);
const F2 = frame(`w:w="${span(4, 4)}" w:hSpace="${GUTTER}" w:wrap="around" w:vAnchor="text" w:hAnchor="margin" w:xAlign="right" w:y="0"`);
const F3 = frame(`w:w="${span(2, 3)}" w:wrap="notBeside" w:vAnchor="text" w:hAnchor="margin" w:x="${colStart(2)}" w:y="0"`);
const F4 = frame(`w:w="${span(2, 3)}" w:h="${2 * LINE}" w:hRule="exact" w:wrap="around" w:vAnchor="margin" w:hAnchor="page" w:xAlign="center" w:y="${lineTop(56)}"`);

// Tables: label column 1, value columns 2-4; nested grid wider than its cell.
const BORDERS = (names) => `<w:tblBorders>${names.map((n) => `<w:${n} w:val="single" w:sz="4" w:space="0" w:color="${COLOR.ink}"/>`).join('')}</w:tblBorders>`;
const MAR = '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>';
const LABEL_W = span(1, 1);
const VALUE_W = span(2, 4);
const NESTED_GRID = [span(2, 2), span(2, 4)];
const nested = (n) => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${BORDERS(['insideV'])}<w:tblLayout w:type="fixed"/>${MAR}</w:tblPr>`
  + `<w:tblGrid>${NESTED_GRID.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`
  + `<w:tr>${NESTED_GRID.map((w, i) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${P(r(i ? `${100 + n}-000${n % 10}` : 'Postal code'))}</w:tc>`).join('')}</w:tr></w:tbl>`;
const tableXml = (floating, rows) => '<w:tbl><w:tblPr>'
  + (floating ? `<w:tblpPr w:leftFromText="0" w:rightFromText="0" w:vertAnchor="text" w:horzAnchor="margin" w:tblpX="0" w:tblpY="0"/><w:tblOverlap w:val="never"/>` : '')
  + `<w:tblW w:w="${LABEL_W + VALUE_W}" w:type="dxa"/>${BORDERS(['top', 'left', 'bottom', 'right', 'insideH', 'insideV'])}<w:tblLayout w:type="fixed"/>${MAR}</w:tblPr>`
  + `<w:tblGrid><w:gridCol w:w="${LABEL_W}"/><w:gridCol w:w="${VALUE_W}"/></w:tblGrid>`
  + Array.from({ length: rows }, (_, k) => `<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="${2 * LINE}" w:hRule="exact"/></w:trPr>`
    + `<w:tc><w:tcPr><w:tcW w:w="${LABEL_W}" w:type="dxa"/></w:tcPr>${P(r(`Row ${k + 1}`, '<w:b/>'))}</w:tc>`
    + `<w:tc><w:tcPr><w:tcW w:w="${VALUE_W}" w:type="dxa"/></w:tcPr>${nested(k + 1)}${P('')}</w:tc></w:tr>`).join('')
  + '</w:tbl>';

export default {
  id: 'docx-049',
  slug: 'frames-floating-table',
  format: 'docx',
  title: 'Frames and a floating table',
  covers: [
    'word/document.xml|w:pPr>w:framePr',
    'word/document.xml|w:framePr@w:wrap=around',
    'word/document.xml|w:framePr@w:wrap=notBeside',
    'word/document.xml|w:framePr@w:xAlign=right',
    'word/document.xml|w:framePr@w:hAnchor=page',
    'word/document.xml|w:framePr@w:hRule=exact',
    'word/document.xml|w:tblPr>w:tblpPr',
    'word/document.xml|w:tblPr>w:tblOverlap',
    'word/document.xml|w:tc>w:tbl',
    'word/document.xml|w:trPr>w:cantSplit',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-049',
      title: 'Frames and a floating table',
      verify: [
        'Verify: A text wraps beside the left and right frames,',
        'not beside the third; the fourth sits centred at line',
        '56. B rows past page 1 resume at the top of page 2.',
      ],
    }),
    caption('A. Frames: wrap around, merged pair, notBeside, page', { at: DOCX.rowLine(2), before: 0 }),
    raw(P(r('F1 frame: two columns wide, three lines high (exact), text wraps around it.'), F1 + BOX + SHADE)),
    raw(P(r(`${FILL} ${FILL}`))),
    raw(P(r('F2 first paragraph, right-aligned frame.'), F2 + BOX)),
    raw(P(r('F2 second paragraph, the same frame.'), F2 + BOX)),
    raw(P(r(`${FILL} ${FILL}`))),
    raw(P(r('F3 frame with wrap notBeside: nothing runs beside it.'), F3 + BOX)),
    raw(P(r(FILL))),
    raw(P(r('F4 page-anchored frame, centred, at line 56 of the margin.'), F4 + BOX + SHADE)),
    raw(P(r('The paragraph that anchors F4 is this one, before B.'))),
    caption('B. Floating fixed-width table continuing on page 2'),
    raw(tableXml(true, 16)),
    raw(P(r('After the floating table.'))),
    caption('C. The same table in the flow'),
    raw(tableXml(false, 4)),
    raw(P(r('After the table in the flow.'))),
  ]),
};

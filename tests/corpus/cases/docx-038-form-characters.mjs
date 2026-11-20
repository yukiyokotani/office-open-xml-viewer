import { DOCX, FONT } from '../lib/grid.mjs';
import { docxPackage, header } from '../lib/docx.mjs';
import { esc } from '../lib/xml.mjs';

// Character-level devices of Japanese forms, where alignment depends on
// exact glyph advances:
//   A. Fit text (w:fitText): a two-character label stretched to the width
//      of four, and a twelve-character label squeezed into eight, written
//      as two runs that share one w:id so they fit as one span. Character
//      scale (w:w) at 66 % and 150 %.
//   B. Check boxes: content controls (w14:checkbox, unchecked and checked)
//      and typed squares "□" followed by a full-width space. A label that
//      continues on a second paragraph is indented by w:firstLineChars 200,
//      so it lines up under the text after "□　", at 9 pt and at 12 pt.
//   C. Labels padded with full-width spaces to four characters in table
//      cells ("氏　　名"), and a date line spaced the same way: the columns
//      line up and no cell gains an extra line.
//   D. A timetable aligned once with a left tab stop at column 2 and once
//      with a full-width space after the time: the names after the tab
//      start on the column edge, those after the space one full width on.
const { LINE, colStart, span } = DOCX;
const JA = `<w:rFonts w:ascii="${FONT.eastAsia}" w:eastAsia="${FONT.eastAsia}" w:hAnsi="${FONT.eastAsia}"/>`;
const SZ12 = '<w:sz w:val="24"/>';
const CHAR = 180; // one full-width character at 9 pt, in twips
const t = (text) => `<w:t xml:space="preserve">${esc(text)}</w:t>`;
const r = (text, rPr = '') => `<w:r><w:rPr>${JA}${rPr}</w:rPr>${t(text)}</w:r>`;
const p = (runs, pPr = '', height = 1) => ({ kind: 'raw', height, xml: `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>` });
const caption = (text, extra = {}) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1, ...extra });
const fit = (width, id) => `<w:fitText w:val="${width}" w:id="${id}"/>`;

const box = (checked) => '<w:sdt><w:sdtPr><w14:checkbox>'
  + `<w14:checked w14:val="${checked ? 1 : 0}"/>`
  + `<w14:checkedState w14:val="25A0" w14:font="${FONT.eastAsia}"/><w14:uncheckedState w14:val="25A1" w14:font="${FONT.eastAsia}"/>`
  + `</w14:checkbox></w:sdtPr><w:sdtContent>${r(checked ? '■' : '□')}</w:sdtContent></w:sdt>`;

const tc = (w, body) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${body}</w:tc>`;
const cp = (runs) => `<w:p>${runs}</w:p>`;
const BORDERS = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
  .map((n) => `<w:${n} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('') + '</w:tblBorders>';
const W1 = colStart(2) - colStart(1);
const W2 = span(2, 3);
const FORM = `<w:tbl><w:tblPr><w:tblW w:w="${W1 + W2}" w:type="dxa"/>${BORDERS}<w:tblLayout w:type="fixed"/>`
  + '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
  + `<w:tblGrid><w:gridCol w:w="${W1}"/><w:gridCol w:w="${W2}"/></w:tblGrid>`
  + [['氏　　名', '用紙　太郎'], ['住　　所', '東京都千代田区一丁目一番一号'], ['電話番号', '０３－００００－００００'], ['日　　付', '令和　　年　　月　　日']]
    .map(([k, v]) => `<w:tr><w:trPr><w:trHeight w:val="${LINE}" w:hRule="exact"/></w:trPr>${tc(W1, cp(r(k)))}${tc(W2, cp(r(v)))}</w:tr>`).join('')
  + '</w:tbl>';

export default {
  id: 'docx-038',
  slug: 'form-characters',
  format: 'docx',
  title: 'Form characters and alignment',
  covers: [
    'word/document.xml|w:rPr>w:fitText',
    'word/document.xml|w:fitText@w:id',
    'word/document.xml|w:rPr>w:w',
    'word/document.xml|w:sdtPr>w14:checkbox',
    'word/document.xml|w14:checkbox>w14:checkedState',
    'word/document.xml|w14:checkbox>w14:uncheckedState',
    'word/document.xml|w:ind@w:firstLineChars',
    'word/document.xml|w:tabs>w:tab',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-038',
      title: 'Form characters and alignment',
      verify: [
        'Verify: fitted labels fill their widths; B labels and their',
        'continuations line up after the squares; C columns align',
        'with no extra lines; D tabbed names start at column 2.',
      ],
    }),
    caption('A. Fit text: 2 characters to 4, 12 to 8 (two runs, one id); w:w', { at: DOCX.rowLine(2), before: 0 }),
    p(r('氏名', fit(4 * CHAR, 1)) + r('：四字分の幅に広げる')),
    p(r('法人の場合は', fit(8 * CHAR, 2)) + r('代表者の氏名', fit(8 * CHAR, 2)) + r('：八字分の幅に詰める')),
    p(r('文字幅66%', '<w:w w:val="66"/>') + r('　標準　') + r('文字幅150%', '<w:w w:val="150"/>')),
    caption('B. Check boxes: content controls, then typed squares at 9 and 12 pt'),
    p(box(false) + r('　未選択の内容制御') + r('　') + box(true) + r('　選択済みの内容制御')),
    p(r('□　会議冒頭の傍聴')),
    p(r('□　下水道幹線の老朽化対策')),
    p(r('工事現場（続きの段落、二字下げ）'), '<w:ind w:firstLineChars="200"/>'),
    p(r('□　十二ポイントの項目', SZ12), '<w:spacing w:line="480" w:lineRule="exact"/>', 2),
    p(r('続きの段落、二字下げ', SZ12), '<w:spacing w:line="480" w:lineRule="exact"/><w:ind w:firstLineChars="200"/>', 2),
    caption('C. Labels padded with full-width spaces in cells'),
    { kind: 'raw', xml: FORM, height: 4 },
    caption('D. Timetable: tab stop, then full-width spaces'),
    p(r('10:00') + `<w:r><w:tab/></w:r>` + r('開会'), `<w:tabs><w:tab w:val="left" w:pos="${colStart(2) - colStart(1)}"/></w:tabs>`),
    p(r('10:15') + `<w:r><w:tab/></w:r>` + r('報告と質疑'), `<w:tabs><w:tab w:val="left" w:pos="${colStart(2) - colStart(1)}"/></w:tabs>`),
    p(r('10:00　開会')),
    p(r('10:15　報告と質疑')),
  ], {
    ns: ['xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"', 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"', 'mc:Ignorable="w14"'],
  }),
};

import { DOCX, FONT } from '../lib/grid.mjs';
import { docxPackage, header } from '../lib/docx.mjs';
import { esc } from '../lib/xml.mjs';

// Japanese punctuation (yakumono) under punctuation compression
// (w:characterSpacingControl compressPunctuation), and line breaking of long
// Latin tokens inside Japanese text.
//   A. Pairs of brackets, commas and full stops in a row: ）。 」、「 （「 」）
//      mid-line, at the end of a line before a manual break, and at the
//      start of the next line. Each pair closes up; no glyph may overlap
//      its neighbour.
//   B. The same pair in right-aligned and centred lines.
//   C. Run boundaries: the same sentence in one run, split into two runs
//      after the comma with Latin following, and split with Japanese
//      following. The split must not change the spacing.
//   D. Spaces kept inside brackets (xml:space="preserve") in a justified
//      paragraph that wraps naturally.
//   E. A URL and an e-mail address in a justified Japanese paragraph in a
//      two-column-wide field: the URL is longer than the line and must break
//      inside itself; nothing may run past the column edge.
// A-C break lines in the text; D and E wrap naturally and say so.
const JA = `<w:rFonts w:ascii="${FONT.eastAsia}" w:eastAsia="${FONT.eastAsia}" w:hAnsi="${FONT.eastAsia}"/>`;
const t = (text) => `<w:t xml:space="preserve">${esc(text)}</w:t>`;
const r = (text, rPr = JA) => `<w:r><w:rPr>${rPr}</w:rPr>${t(text)}</w:r>`;
const rLines = (texts) => `<w:r><w:rPr>${JA}</w:rPr>${texts.map((x, i) => `${i ? '<w:br/>' : ''}${t(x)}`).join('')}</w:r>`;
const IND_2 = `<w:ind w:right="${DOCX.LIVE_W - DOCX.colEnd(2)}"/>`;
const IND_4 = '';

const para = (runs, { jc, ind = IND_4, before = 0 } = {}) => ({
  kind: 'raw',
  height: NaN,
  xml: `<w:p><w:pPr>${before ? `<w:spacing w:before="${before * DOCX.LINE}"/>` : ''}${ind}${jc ? `<w:jc w:val="${jc}"/>` : ''}</w:pPr>${runs}</w:p>`,
});
const fixed = (runs, n, opts = {}) => ({ ...para(runs, opts), height: n });
const caption = (text, before = 1) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before });

export default {
  id: 'docx-036',
  slug: 'japanese-punctuation',
  format: 'docx',
  title: 'Japanese punctuation and breaks',
  covers: [
    'word/settings.xml|w:characterSpacingControl@w:val=compressPunctuation',
    'word/document.xml|w:pPr>w:jc',
    'word/document.xml|w:jc@w:val=right',
    'word/document.xml|w:jc@w:val=center',
    'word/document.xml|w:jc@w:val=both',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-036',
      title: 'Japanese punctuation and breaks',
      verify: [
        'Verify: paired marks close up and never overlap, also at line',
        'ends and starts; split runs set like one run; D and E wrap',
        'naturally and the URL breaks inside the column.',
      ],
    }),
    { ...caption('A. Pairs mid-line, before a line break and after it', 0), at: DOCX.rowLine(2) },
    fixed(rLines([
      '前文である（予定の資料でも構わない）。後文が続く。',
      '「規格」、「寸法」、「用紙」の三語を並べる。',
      '用紙の寸法（「A4」「B5」など）を示す行である。',
      'この行は括弧と句点で終わる（A4判）。',
      '（「B5判」）から始まる行である。',
    ]), 5),
    caption('B. Right-aligned and centred lines'),
    fixed(rLines(['右揃えの行は（括弧）。で終わる。']), 1, { jc: 'right' }),
    fixed(rLines(['中央揃えの行は「鉤括弧」、で区切る。']), 1, { jc: 'center' }),
    caption('C. One run, then two runs split after the comma'),
    fixed(r('確認した後に、E-mailで送信する。'), 1),
    fixed(r('確認した後に、') + r('E-mailで送信する。'), 1),
    fixed(r('確認した後に、') + r('次の手順へ進む。'), 1),
    caption('D. Kept spaces inside brackets, justified, natural wrap'),
    para(r('用紙の寸法は（ 幅 × 高さ ）の順に書き、単位は（ mm ）とする。A列の各判は面積が半分ずつになり、縦横比は（ 1 : 1.414 ）で変わらない。'), { jc: 'both', ind: IND_2 }),
    caption('E. URL and address in justified text, natural wrap'),
    para(r('寸法の一覧はhttps://www.example.org/standards/paper-sizes/iso-216-a-series-and-b-series.htmlに掲載されており、問い合わせはpaper.sizes.reference@example.orgで受け付ける。'), { jc: 'both', ind: IND_2 }),
  ], {
    settings: '<w:characterSpacingControl w:val="compressPunctuation"/>',
  }),
};

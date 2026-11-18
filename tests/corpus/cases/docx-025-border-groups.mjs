import { DOCX, FONT } from '../lib/grid.mjs';
import { docxPackage, header, run } from '../lib/docx.mjs';

// How Word groups bordered paragraphs: consecutive paragraphs with the same
// borders and the same indents share one box (w:between draws a rule between
// them); a paragraph with a different indent starts its own box. Text wraps at
// the paragraph indents, and the side rules sit w:space points outside them.
// Typed bullets and circled numbers open lines in a Japanese box.
const SZ = 4; // ½ pt
const SPACE = 4; // pt
const edge = (name) => `<w:${name} w:val="single" w:sz="${SZ}" w:space="${SPACE}" w:color="000000"/>`;
const BOX = `<w:pBdr>${['top', 'left', 'bottom', 'right'].map(edge).join('')}</w:pBdr>`;
const BOX_BETWEEN = `<w:pBdr>${['top', 'left', 'bottom', 'right', 'between'].map(edge).join('')}</w:pBdr>`;
const JA = `<w:rFonts w:ascii="${FONT.eastAsia}" w:eastAsia="${FONT.eastAsia}" w:hAnsi="${FONT.eastAsia}"/>`;
const IND = { 'w:left': DOCX.GUTTER, 'w:right': DOCX.GUTTER };
const IND_DEEP = { 'w:left': DOCX.colStart(2), 'w:right': DOCX.GUTTER };
const caption = (text) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1 });
const boxed = (text, extra = {}) => ({ kind: 'p', pBdr: BOX, ind: IND, text, ...extra });

export default {
  id: 'docx-025',
  slug: 'border-groups',
  format: 'docx',
  title: 'Border groups',
  covers: [
    'word/document.xml|w:pBdr>w:between',
    'word/document.xml|w:pPr>w:pBdr',
    'word/document.xml|w:pPr>w:ind',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-025',
      title: 'Border groups',
      verify: [
        'Verify: A is one box round three paragraphs; B adds rules',
        'between them; C splits where the indent changes; in D the',
        'rules run straight past the ● and ① that open lines.',
      ],
    }),
    caption('A. One box: same borders, same indents'),
    boxed('Three paragraphs carry the same w:pBdr and the same indents of one gutter on each side. Word draws them as one box, with no rule between them.'),
    boxed('The side rules stand 4 pt outside the indents (w:space), so the text wraps at the indents and never reaches the right rule, however long the line runs before it breaks.'),
    boxed('Third paragraph of the box.'),
    caption('B. Rules between paragraphs: w:between'),
    { kind: 'p', pBdr: BOX_BETWEEN, ind: IND, text: 'The same box with a w:between border.' },
    { kind: 'p', pBdr: BOX_BETWEEN, ind: IND, text: 'A rule separates each paragraph from the next inside the box.' },
    caption('C. A different indent starts a new box'),
    boxed('This paragraph is indented one gutter.'),
    boxed('This one is indented to grid column 2, so it is drawn in a box of its own whose left rule moves with the indent.', { ind: IND_DEEP }),
    boxed('Back to one gutter: a third box.'),
    caption('D. Typed bullets and circled numbers'),
    { kind: 'p', pBdr: BOX, ind: IND, runs: [run('● 枠の中の行頭に、記号を文字として入力した段落である。罫線は記号の位置に関係なく、段落の字下げに沿ってまっすぐ引かれる。', JA)] },
    { kind: 'p', pBdr: BOX, ind: IND, runs: [run('①　丸数字で始まる行。②　同じ段落の中の二つ目の丸数字。③　三つ目。', JA)] },
    { kind: 'p', pBdr: BOX, ind: IND, runs: [run('●①　記号と丸数字が続けて行頭に来る場合も、左の罫線はその左に引かれる。', JA)] },
  ]),
};

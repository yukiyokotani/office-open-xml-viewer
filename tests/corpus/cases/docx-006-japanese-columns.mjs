import { DOCX } from '../lib/grid.mjs';
import { docxPackage, headerSection } from '../lib/docx.mjs';

const COLS = { num: 2, space: DOCX.gutterAfter(2) };
// The document grid repeats the 12 pt baseline: line pitch 240 twips.
const DOC_GRID = `<w:docGrid w:type="lines" w:linePitch="${DOCX.LINE}"/>`;

const PROSE = [
  '版面とは、紙面のうち文字や図版が置かれる領域をいう。この見本の版面は幅165 mm、高さ254 mmであり、'
    + '12ポイントの行送りで60行を収める。本文は9ポイントのNoto Sans JPで組み、欧文と数字はRobotoで組む。',
  '段組は、版面を複数の段に分けて一行の長さを短くする方法である。行が長すぎると、読み手の視線は行末から'
    + '次の行頭へ戻るときに迷いやすい。ここでは版面を幅80 mmの二段に分け、段間を5 mmとした。',
  '和文の文字は正方形の仮想ボディを持ち、字送りを一定に保てば、文字は行の上に等間隔に並ぶ。'
    + '句読点、閉じ括弧、長音符号、小書きの仮名が行頭に来ないように、'
    + 'また開き括弧が行末に残らないように行を調整する規則を「禁則処理」という。'
    + 'この見本ではWordの既定の禁則処理に任せ、段落の配置も既定の左揃えのままとしている。',
  '行送りは、ある行の基準線から次の行の基準線までの距離である。本文の行送り12 ptは文字の大きさ9 ptの'
    + '約1.33倍にあたる。行間はその差の3 ptであり、二つの段の行は同じ基準線の上にそろう。',
];

export default {
  id: 'docx-006',
  slug: 'japanese-columns',
  format: 'docx',
  title: 'Japanese text in columns',
  covers: [
    'word/styles.xml|w:rFonts@w:eastAsia',
    'word/styles.xml|w:lang@w:eastAsia',
    'word/document.xml|w:sectPr>w:docGrid',
    'word/document.xml|w:docGrid@w:type=lines',
    'word/document.xml|w:docGrid@w:linePitch',
    'word/document.xml|w:cols@w:num',
  ],
  parts: () => docxPackage([
    ...headerSection({
      id: 'docx-006',
      title: 'Japanese text in columns',
      verify: [
        'Verify: no line starts with a comma, full stop, closing',
        'bracket, long-vowel mark or small kana; none ends with an',
        'opening bracket; Latin and digits are spaced from kana.',
      ],
    }, { cols: COLS, docGrid: DOC_GRID }),
    ...PROSE.map((text, i) => ({ kind: 'p', text, before: i ? 1 : 0 })),
    { kind: 'section', type: 'continuous', docGrid: DOC_GRID },
    {
      kind: 'p',
      before: 1,
      style: 'caption',
      text: `Sections 2 and 3 carry w:docGrid type lines with a line pitch of ${DOCX.LINE} twips. `
        + 'w:kinsoku, w:autoSpaceDE, w:autoSpaceDN and w:jc are omitted, so Word applies its defaults: '
        + 'Japanese line-breaking rules on, extra space between Latin or digits and Japanese text, alignment left.',
    },
  ], { firstSection: { docGrid: DOC_GRID } }),
};

import { DOCX } from '../lib/grid.mjs';
import { docxPackage, run } from '../lib/docx.mjs';

// tbRl: characters run top to bottom, lines advance right to left. The page
// grid turns with the text: line pitch stays 240 twips, now measured across
// the 165 mm width of the live area.
const SECTION = {
  textDirection: 'tbRl',
  docGrid: `<w:docGrid w:type="lines" w:linePitch="${DOCX.LINE}"/>`,
};

// Tate-chu-yoko: a two-digit run set horizontally inside the vertical line.
const tcy = (digits, id) => run(digits, `<w:eastAsianLayout w:id="${id}" w:vert="1"/>`);

// Short original sentences, each about the glyphs it contains.
const PROSE = [
  [run('句読点の「、」と「。」は、縦組みでは字面の右上に置かれる。')],
  [run('丸括弧（　）、亀甲括弧〔　〕、かぎ括弧「　」、二重かぎ括弧『　』は九十度回転する。')],
  [run('長音符号は「データ」「ルーラー」のように縦向きの線になる。')],
  [run('波ダッシュ〜、三点リーダー…、二点リーダー‥も縦組み用の字形に替わる。')],
  [run('コロン：とセミコロン；も縦組み用の字形を持つ。')],
  [run('小書きの仮名は「ちょっと」「しゅっぱつ」のように右上に寄る。')],
  [run('欧文のRoboto 2026は横に倒れ、縦中横の'), tcy('12', 1), run('月'), tcy('24', 2), run('日は正立する。')],
];

export default {
  id: 'docx-007',
  slug: 'vertical-japanese',
  format: 'docx',
  title: 'Vertical writing',
  covers: [
    'word/document.xml|w:sectPr>w:textDirection',
    'word/document.xml|w:textDirection@w:val=tbRl',
    'word/document.xml|w:docGrid@w:type=lines',
    'word/document.xml|w:rPr>w:eastAsianLayout',
    'word/document.xml|w:eastAsianLayout@w:vert',
  ],
  parts: () => docxPackage([
    { kind: 'p', style: 'caption', bold: true, lines: ['docx-007'] },
    { kind: 'p', style: 'title', lines: ['Vertical writing'] },
    {
      kind: 'p',
      style: 'caption',
      before: 1,
      lines: [
        'Verify: 、。 sit top-right; brackets, ー, 〜, …, ‥, ： and ； turn vertical;',
        'small kana shift top-right; Latin rotates; 12 and 24 stay upright (tate-chu-yoko).',
      ],
    },
    ...PROSE.map((runs, i) => ({ kind: 'p', runs, before: i ? 0 : 1 })),
  ], { firstSection: SECTION }),
};

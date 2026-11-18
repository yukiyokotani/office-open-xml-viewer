import { LANG } from '../lib/grid.mjs';
import { field, header, pptxPackage, textBox } from '../lib/pptx.mjs';

// Running text, so the text size 10/12: horizontal in five grid columns
// (about 36 full-width characters), vertical in four grid columns by three rows.
const ja = (text) => ({ style: 'caption', text, lang: LANG.eastAsia, altLang: LANG.latin });

export default {
  id: 'pptx-005',
  slug: 'japanese-text',
  format: 'pptx',
  title: 'Japanese text',
  covers: [
    'ppt/slides/slide1.xml|a:rPr@altLang',
    'ppt/slides/slide1.xml|a:bodyPr@vert=eaVert',
    'ppt/theme/theme1.xml|a:minorFont>a:ea',
    'ppt/presentation.xml|a:defRPr>a:ea',
  ],
  parts: () => pptxPackage([
    header({
      id: 'pptx-005',
      title: 'Japanese text',
      verify: [
        'Verify: the theme east-Asian font sets the kana and kanji; no line starts with 、。」',
        'or ー; in the vertical box 、。 sit top-right and Latin and digits rotate.',
      ],
    }),
    textBox(3, 'Horizontal', field(1, 5, 2, 4), [
      ja('文字の大きさは10ポイント、行送りは12ポイントである。和文の書体はテーマの東アジア用フォント（+mn-ea）から、欧文の書体はラテン用フォントから選ばれる。'),
      ja('この文字枠の幅は5グリッド列で、一行におよそ36字が並ぶ。句読点や閉じ括弧が行頭に来ないように、PowerPointは行を調整する。'),
      ja('「Roboto」と「Noto Sans JP」は、どちらもSIL Open Font Licenseで公開されている書体である。'),
    ]),
    textBox(4, 'Vertical', field(9, 12, 2, 4), [
      ja('縦書きの文字枠では、文字は上から下へ進み、行は右から左へ進む。'),
      ja('句読点「、」「。」は右上に寄り、長音ーは縦になる。'),
      ja('欧文Robotoと数字2026は横に倒れる。'),
    ], '', { vert: 'eaVert' }),
  ]),
};

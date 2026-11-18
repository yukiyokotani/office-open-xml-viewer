import { COLOR, FONT } from '../lib/grid.mjs';
import { THEME, field, header, pptxPackage, shape, textBox } from '../lib/pptx.mjs';

// A theme with a full style matrix: three solid fills (the placeholder colour,
// a tint and a saturated shade of it), three line styles (solid, dash, dot)
// and an outer shadow effect. Gradients are left to pptx-009. The shapes carry no fill or line of their own; p:style picks
// the matrix entries and fills the placeholder colour phClr with accent1.
const ph = (mods = '') => `<a:schemeClr val="phClr">${mods}</a:schemeClr>`;
const fill = (mods = '') => `<a:solidFill>${ph(mods)}</a:solidFill>`;
const FILLS = fill() + fill('<a:tint val="40000"/>') + fill('<a:shade val="60000"/><a:satMod val="150000"/>');
const LINES = [[12700, 'solid'], [25400, 'dash'], [38100, 'sysDot']]
  .map(([w, dash]) => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>${ph()}</a:solidFill><a:prstDash val="${dash}"/><a:miter lim="800000"/></a:ln>`).join('');
const SHADOW = '<a:effectLst><a:outerShdw blurRad="76200" dist="76200" dir="2700000" algn="tl" rotWithShape="0">'
  + `<a:srgbClr val="${COLOR.ink}"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>`;
const EFFECTS = '<a:effectStyle><a:effectLst/></a:effectStyle>'.repeat(2) + `<a:effectStyle>${SHADOW}</a:effectStyle>`;
const BG_FILLS = fill() + fill('<a:tint val="95000"/>') + fill('<a:shade val="95000"/>');
const FMT_SCHEME = '<a:fmtScheme name="Corpus matrix">'
  + `<a:fillStyleLst>${FILLS}</a:fillStyleLst><a:lnStyleLst>${LINES}</a:lnStyleLst>`
  + `<a:effectStyleLst>${EFFECTS}</a:effectStyleLst><a:bgFillStyleLst>${BG_FILLS}</a:bgFillStyleLst></a:fmtScheme>`;

const theme = THEME
  .replace(/<a:dk1>.*?<\/a:dk1>/, `<a:dk1><a:sysClr val="windowText" lastClr="${COLOR.ink}"/></a:dk1>`)
  .replace(/<a:lt1>.*?<\/a:lt1>/, `<a:lt1><a:sysClr val="window" lastClr="${COLOR.paper}"/></a:lt1>`)
  .replace(/<\/a:majorFont>/, `<a:font script="Jpan" typeface="${FONT.eastAsia}"/></a:majorFont>`)
  .replace(/<\/a:minorFont>/, `<a:font script="Jpan" typeface="${FONT.eastAsia}"/></a:minorFont>`)
  .replace(/<a:fmtScheme[\s\S]*<\/a:fmtScheme>/, FMT_SCHEME)
  .replace('</a:themeElements>', '</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/>');
if (theme === THEME || !theme.includes('Corpus matrix') || !theme.includes('sysClr')) throw new Error('pptx-019: theme rewrite failed');

const style = (fill, ln, effect) => `<p:style><a:lnRef idx="${ln}"><a:schemeClr val="tx1"/></a:lnRef>`
  + `<a:fillRef idx="${fill}"><a:schemeClr val="accent1"/></a:fillRef>`
  + `<a:effectRef idx="${effect}"><a:schemeClr val="accent1"/></a:effectRef>`
  + '<a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style>';

const SHAPES = [
  { name: 'Matrix 1', ref: [1, 1, 0], lines: ['fillRef 1, lnRef 1', 'solid accent fill', '1 pt solid black line'] },
  { name: 'Matrix 2', ref: [2, 2, 0], lines: ['fillRef 2, lnRef 2', 'tint 40 % of the accent', '2 pt dashed black line'] },
  { name: 'Matrix 3', ref: [3, 3, 0], lines: ['fillRef 3, lnRef 3', 'shade 60 %, saturation 150 %', '3 pt dotted black line'] },
  { name: 'Matrix effect', ref: [1, 0, 3], lines: ['fillRef 1, effectRef 3', 'solid accent fill, no line', 'shadow 6 pt down-right, 40 % black'] },
];

export default {
  id: 'pptx-019',
  slug: 'theme-style-matrix',
  format: 'pptx',
  title: 'Theme style matrix',
  covers: [
    'ppt/theme/theme1.xml|a:schemeClr>a:tint',
    'ppt/theme/theme1.xml|a:schemeClr>a:shade',
    'ppt/theme/theme1.xml|a:schemeClr>a:satMod',
    'ppt/theme/theme1.xml|a:ln>a:prstDash',
    'ppt/theme/theme1.xml|a:effectLst>a:outerShdw',
    'ppt/theme/theme1.xml|a:outerShdw>a:srgbClr',
    'ppt/theme/theme1.xml|a:srgbClr>a:alpha',
    'ppt/theme/theme1.xml|a:dk1>a:sysClr',
    'ppt/theme/theme1.xml|a:majorFont>a:font',
    'ppt/theme/theme1.xml|a:theme>a:objectDefaults',
    'ppt/theme/theme1.xml|a:theme>a:extraClrSchemeLst',
    'ppt/slides/slide1.xml|p:style>a:lnRef',
    'ppt/slides/slide1.xml|p:style>a:effectRef',
    'ppt/slides/slide1.xml|p:bg>p:bgRef',
  ],
  parts: () => pptxPackage([
    header({
      id: 'pptx-019',
      title: 'Theme style matrix',
      verify: [
        'Verify: four accent shapes, three columns by two rows each, styled only by the theme matrix:',
        'accent, pale tint and dark shade with 1/2/3 pt lines, then a shadowed accent.',
      ],
    }),
    ...SHAPES.flatMap((s, i) => {
      const c1 = 1 + 3 * i;
      const sp = shape(3 + 2 * i, s.name, field(c1, c1 + 2, 2, 3), 'rect', '', '');
      return [
        sp.replace('</p:spPr></p:sp>', `</p:spPr>${style(...s.ref)}</p:sp>`),
        textBox(4 + 2 * i, `Caption ${i + 1}`, field(c1, c1 + 2, 4, 4), s.lines.map((text, j) => ({ style: 'caption', bold: j === 0, text }))),
      ];
    }),
  ], {
    slideBg: '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>',
    parts: { 'ppt/theme/theme1.xml': theme },
  }),
};

import { COLOR, DOCX, FONT, LANG } from '../lib/grid.mjs';
import { W, WML, docxPackage, header, run } from '../lib/docx.mjs';
import { THEME } from '../lib/pptx.mjs';
import { XML_DECL } from '../lib/xml.mjs';

const at = DOCX.rowLine;
const TAB = [['left', DOCX.colStart(3)]];
// Accent 1 at themeShade 80 (hex): each channel × 0x80 / 0xFF.
const shade = (rgb, s) => rgb.match(/../g).map((h) => Math.round((parseInt(h, 16) * s) / 255).toString(16).padStart(2, '0')).join('').toUpperCase();
const SHADED = shade(COLOR.accent, 0x80);

const LINES = [
  ['w:rFonts asciiTheme minorHAnsi', '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/>', `The theme's minor Latin font, ${FONT.latin}.`],
  ['w:rFonts asciiTheme majorHAnsi', '<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>', `The theme's major Latin font, also ${FONT.latin}.`],
  ['w:rFonts eastAsiaTheme minorEastAsia', '<w:rFonts w:eastAsiaTheme="minorEastAsia"/>', 'テーマの東アジア用フォントで組む和文。'],
  ['w:color themeColor accent1', `<w:color w:val="${COLOR.accent}" w:themeColor="accent1"/>`, `Accent 1 of the theme, ${COLOR.accent}.`],
  ['w:color themeColor accent1, themeShade 80', `<w:color w:val="${SHADED}" w:themeColor="accent1" w:themeShade="80"/>`, `Accent 1 shaded to half, ${SHADED}.`],
  ['w:color themeColor text1', `<w:color w:val="${COLOR.ink}" w:themeColor="text1"/>`, 'Text 1, mapped to dark 1 by w:clrSchemeMapping.'],
];

const FONT_TABLE = XML_DECL + `<w:fonts ${W}>`
  + `<w:font w:name="${FONT.latin}"><w:panose1 w:val="02000000000000000000"/><w:charset w:val="00"/><w:family w:val="auto"/><w:pitch w:val="variable"/></w:font>`
  + `<w:font w:name="${FONT.eastAsia}"><w:altName w:val="Noto Sans CJK JP"/><w:charset w:val="80"/><w:family w:val="swiss"/><w:pitch w:val="variable"/></w:font>`
  + '</w:fonts>';

const CLR_MAP = '<w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" '
  + 'w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/>';

export default {
  id: 'docx-014',
  slug: 'theme-fonts',
  format: 'docx',
  title: 'Theme fonts and colours',
  covers: [
    'word/document.xml|w:rFonts@w:asciiTheme',
    'word/document.xml|w:rFonts@w:eastAsiaTheme',
    'word/document.xml|w:color@w:themeColor',
    'word/document.xml|w:color@w:themeShade',
    'word/theme/theme1.xml|a:fontScheme>a:minorFont',
    'word/theme/theme1.xml|a:clrScheme>a:accent1',
    'word/settings.xml|w:settings>w:themeFontLang',
    'word/settings.xml|w:settings>w:clrSchemeMapping',
    'word/fontTable.xml|w:fonts>w:font',
    'word/fontTable.xml|w:font>w:altName',
    'word/fontTable.xml|w:font>w:panose1',
    'word/fontTable.xml|w:font>w:charset',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-014',
      title: 'Theme fonts and colours',
      verify: [
        'Verify: every sample resolves through the theme part:',
        `${FONT.latin} and ${FONT.eastAsia}, accent orange, its half shade, black.`,
      ],
    }),
    { kind: 'p', at: at(2), style: 'caption', bold: true, tabs: TAB, lines: ['Run property\tSample'] },
    ...LINES.map(([name, rPr, sample]) => ({ kind: 'p', tabs: TAB, height: 1, runs: [run(`${name}\t`), run(sample, rPr)] })),
    {
      kind: 'p', at: at(3), style: 'caption', cols: [1, 3],
      lines: [
        `The font table lists both faces; ${FONT.eastAsia} carries w:altName Noto Sans CJK JP,`,
        `the name under which many systems install it. Theme languages: ${LANG.latin}, ${LANG.eastAsia}.`,
      ],
    },
  ], {
    settingsAfter: `<w:themeFontLang w:val="${LANG.latin}" w:eastAsia="${LANG.eastAsia}"/>${CLR_MAP}`,
    parts: { 'word/theme/theme1.xml': THEME, 'word/fontTable.xml': FONT_TABLE },
    rels: [['rId3', 'theme', 'theme/theme1.xml'], ['rId4', 'fontTable', 'fontTable.xml']],
    overrides: {
      'word/theme/theme1.xml': 'application/vnd.openxmlformats-officedocument.theme+xml',
      'word/fontTable.xml': `${WML}.fontTable+xml`,
    },
  }),
};

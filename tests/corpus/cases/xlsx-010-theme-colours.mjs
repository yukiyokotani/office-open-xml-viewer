import { FONT, XLSX } from '../lib/grid.mjs';
import { THEME } from '../lib/pptx.mjs';
import { sheet } from '../lib/xlsx.mjs';

const { MARGIN_COL, MODULE, FIRST_ROW } = XLSX;
// SpreadsheetML theme colour indices: 0 lt1, 1 dk1, 2 lt2, 3 dk2, 4 accent1, ...
const SAMPLES = [
  ['font scheme minor', { scheme: 'minor' }, `Theme minor font, ${FONT.latin}`],
  ['font color theme 1 (dk1)', { scheme: 'minor', themeColor: [1] }, 'Dark 1: black'],
  ['font color theme 4 (accent1)', { scheme: 'minor', themeColor: [4] }, 'Accent 1: cyan'],
  ['fill theme 4', { fill: 'theme:4' }, null],
  ['fill theme 4, tint -0.5', { fill: 'theme:4:-0.5' }, null],
  ['fill theme 4, tint 0.8', { fill: 'theme:4:0.8' }, null],
];

export default {
  id: 'xlsx-010',
  slug: 'theme-colours',
  format: 'xlsx',
  title: 'Theme fonts and colours',
  covers: [
    'xl/theme/theme1.xml|a:clrScheme>a:accent1',
    'xl/theme/theme1.xml|a:fontScheme>a:minorFont',
    'xl/styles.xml|x:font>x:scheme',
    'xl/styles.xml|x:color@theme',
    'xl/styles.xml|x:fgColor@theme',
    'xl/styles.xml|x:fgColor@tint',
  ],
  parts() {
    const s = sheet({
      id: 'xlsx-010',
      title: 'Theme fonts and colours',
      verify: 'Verify: the font is the theme minor font; fills from top: accent cyan, a darker cyan (tint -0.5), a pale cyan (tint 0.8).',
      cols: [['A', MARGIN_COL], ['B', 2 * MODULE], ['C', MODULE], ['D', MODULE], ['E', MODULE]],
    });
    s.theme(THEME);
    s.set(`B${FIRST_ROW}`, 'Style', { font: 'captionBold' });
    s.set(`C${FIRST_ROW}`, 'Sample', { font: 'captionBold' });
    SAMPLES.forEach(([label, style, text], i) => {
      const r = FIRST_ROW + 1 + 2 * i;
      s.set(`B${r}`, label);
      if (text) s.set(`C${r}`, text, style);
      else for (const col of ['C', 'D', 'E']) s.set(`${col}${r}`, null, style);
    });
    return s.build();
  },
};

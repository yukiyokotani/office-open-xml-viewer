import { COLOR, XLSX } from '../lib/grid.mjs';
import { sheet } from '../lib/xlsx.mjs';

const { MARGIN_COL, MODULE, FIRST_ROW } = XLSX;
const SAMPLES = [
  [`fill ${COLOR.ink}, font ${COLOR.paper}`, { fill: COLOR.ink, color: COLOR.paper }, 'White on black'],
  [`fill ${COLOR.accent}, font ${COLOR.ink}`, { fill: COLOR.accent }, 'Black on cyan'],
  [`fill ${COLOR.accent}, font ${COLOR.paper}`, { fill: COLOR.accent, color: COLOR.paper }, 'White on cyan'],
  [`font ${COLOR.accent}, no fill`, { color: COLOR.accent }, 'Cyan on white'],
];

export default {
  id: 'xlsx-005',
  slug: 'fills-font-colour',
  format: 'xlsx',
  title: 'Fills and font colour',
  covers: [
    'xl/styles.xml|x:fill>x:patternFill',
    'xl/styles.xml|x:patternFill@patternType=solid',
    'xl/styles.xml|x:patternFill>x:fgColor',
    'xl/styles.xml|x:fgColor@rgb',
    'xl/styles.xml|x:font>x:color',
    'xl/styles.xml|x:color@rgb',
    'xl/styles.xml|x:xf@fillId',
    'xl/styles.xml|x:xf@applyFill',
  ],
  parts() {
    const s = sheet({
      id: 'xlsx-005',
      title: 'Fills and font colour',
      verify: 'Verify: each fill covers C to E of its row exactly; the text colour is as its label states.',
      cols: [['A', MARGIN_COL], ['B', 2 * MODULE], ['C', MODULE], ['D', MODULE], ['E', MODULE]],
    });
    s.set(`B${FIRST_ROW}`, 'Solid pattern fill and font colour, RGB', { font: 'captionBold' });
    SAMPLES.forEach(([text, style, sample], i) => {
      const r = FIRST_ROW + 1 + 2 * i;
      s.set(`B${r}`, text);
      s.set(`C${r}`, sample, style);
      if (style.fill) for (const col of ['D', 'E']) s.set(`${col}${r}`, null, { fill: style.fill });
    });
    return s.build();
  },
};

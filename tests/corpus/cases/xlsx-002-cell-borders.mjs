import { XLSX } from '../lib/grid.mjs';
import { sheet } from '../lib/xlsx.mjs';

const { MARGIN_COL, MODULE, FIRST_ROW } = XLSX;

// ST_BorderStyle values, in order of visual weight (ECMA-376 Part 1, §18.18.3).
const STYLES = ['hair', 'dotted', 'dashed', 'thin', 'medium', 'thick', 'double'];
// Specimens sit on every other row from FIRST_ROW + 1; after one empty row
// comes the caption, then the box.
const BOX_TOP = XLSX.FIRST_ROW + 2 * STYLES.length + 2;
const BOX = `C${BOX_TOP}:E${BOX_TOP + 2}`;

export default {
  id: 'xlsx-002',
  slug: 'cell-borders',
  format: 'xlsx',
  title: 'Cell borders',
  covers: [
    'xl/styles.xml|x:border>x:left',
    'xl/styles.xml|x:border>x:right',
    'xl/styles.xml|x:border>x:top',
    'xl/styles.xml|x:border>x:bottom',
    ...STYLES.map((style) => `xl/styles.xml|x:bottom@style=${style}`),
    'xl/styles.xml|x:xf@borderId',
  ],
  parts() {
    const s = sheet({
      id: 'xlsx-002',
      title: 'Cell borders',
      verify: `Verify: each rule runs along the bottom edge of C to E in its row; the box encloses ${BOX}.`,
      cols: [['A', MARGIN_COL], ['B', MODULE], ['C', MODULE], ['D', MODULE], ['E', MODULE]],
    });
    s.set(`B${FIRST_ROW}`, 'Bottom edge of C:E, by border style', { font: 'captionBold' });
    STYLES.forEach((style, i) => {
      const r = FIRST_ROW + 1 + 2 * i;
      s.set(`B${r}`, style);
      for (const col of ['C', 'D', 'E']) s.set(`${col}${r}`, null, { border: { bottom: style } });
    });
    const boxTop = BOX_TOP;
    s.set(`B${boxTop - 1}`, `All four edges, thin, around ${BOX}`, { font: 'captionBold' });
    for (let r = boxTop; r < boxTop + 3; r += 1) {
      for (const col of ['C', 'D', 'E']) {
        const border = {};
        if (r === boxTop) border.top = 'thin';
        if (r === boxTop + 2) border.bottom = 'thin';
        if (col === 'C') border.left = 'thin';
        if (col === 'E') border.right = 'thin';
        s.set(`${col}${r}`, null, { border });
      }
    }
    return s.build();
  },
};

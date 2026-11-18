import { COLOR, DOCX } from '../lib/grid.mjs';
import { docxPackage, header } from '../lib/docx.mjs';

const at = DOCX.rowLine;
// A 1 pt rule (w:sz 8) with 11 pt of space (w:space) takes exactly one 12 pt
// baseline above or below the text, so bordered blocks stay on the grid.
const SZ = 8;
const SPACE = DOCX.LINE / 20 - SZ / 8;
const edge = (name, color = COLOR.ink) => `<w:${name} w:val="single" w:sz="${SZ}" w:space="${SPACE}" w:color="${color}"/>`;
const pBdr = (edges, color) => `<w:pBdr>${edges.map((e) => edge(e, color)).join('')}</w:pBdr>`;
const SHD = `<w:shd w:val="clear" w:color="auto" w:fill="${COLOR.accent}"/>`;
const BOX = ['top', 'left', 'bottom', 'right'];
const caption = (row, text) => ({ kind: 'p', at: at(row), style: 'caption', bold: true, lines: [text] });

export default {
  id: 'docx-012',
  slug: 'borders-shading',
  format: 'docx',
  title: 'Paragraph borders and shading',
  covers: [
    'word/document.xml|w:pPr>w:pBdr',
    'word/document.xml|w:pBdr>w:top',
    'word/document.xml|w:pBdr>w:bottom',
    'word/document.xml|w:pBdr>w:left',
    'word/document.xml|w:pBdr>w:right',
    'word/document.xml|w:top@w:space',
    'word/document.xml|w:pPr>w:shd',
    'word/document.xml|w:shd@w:fill',
    'word/document.xml|w:shd@w:val=clear',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-012',
      title: 'Borders and shading',
      verify: [
        'Verify: each 1 pt rule sits one baseline from its text;',
        'side rules fall in the gutters; the orange fill spans the box.',
      ],
    }),
    caption(2, `Top border: w:sz ${SZ} (1 pt), w:space ${SPACE} pt, one baseline above the text`),
    { kind: 'p', before: 1, extraLines: 1, cols: [2, 3], pBdr: pBdr(['top']), lines: ['This paragraph has a top border only.'] },
    { kind: 'p', before: 1, extraLines: 1, cols: [2, 3], pBdr: pBdr(['bottom']), lines: ['This paragraph has a bottom border only.'] },
    caption(3, 'Box: top, left, bottom and right; the side rules sit in the gutters'),
    {
      kind: 'p', before: 1, extraLines: 2, cols: [2, 3], pBdr: pBdr(BOX),
      lines: ['The text keeps the edges of columns 2 and 3;', 'the side rules stand 12 pt outside, in the gutters.'],
    },
    caption(4, `Shading: w:shd w:fill ${COLOR.accent}, inside a box of the same colour`),
    {
      kind: 'p', before: 1, extraLines: 2, cols: [2, 3], pBdr: pBdr(BOX, COLOR.accent), shd: SHD,
      lines: ['The fill covers the text and the border space:', 'one baseline above and below the text.'],
    },
    caption(5, 'Shading without borders: the fill ends at the text lines'),
    { kind: 'p', before: 1, cols: [2, 3], shd: SHD, lines: ['A shaded paragraph with no border is as tall as its lines.'] },
  ]),
};

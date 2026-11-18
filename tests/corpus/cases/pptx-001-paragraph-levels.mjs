import { PPTX } from '../lib/grid.mjs';
import { field, header, pptxPackage, textBox } from '../lib/pptx.mjs';

const LEVELS = [0, 1, 2, 3, 4];
const box = field(1, 12, 2, 3);

// The text box's own list style carries everything that varies by level: the
// left margin steps one column module (column + gutter) per level, and every
// level keeps the 20/24 pt text size with one empty 12 pt line before it.
const lstStyle = LEVELS.map((lvl) => `<a:lvl${lvl + 1}pPr marL="${lvl * PPTX.STEP}" indent="0">`
  + `<a:lnSpc><a:spcPts val="${PPTX.TYPE.text.leading}"/></a:lnSpc>`
  + `<a:spcBef><a:spcPts val="${lvl === 0 ? 0 : PPTX.LINE / 127}"/></a:spcBef>`
  + `<a:defRPr sz="${PPTX.TYPE.text.sz}"/></a:lvl${lvl + 1}pPr>`).join('');

export default {
  id: 'pptx-001',
  slug: 'paragraph-levels',
  format: 'pptx',
  title: 'Paragraph levels',
  covers: [
    'ppt/slides/slide1.xml|a:lstStyle>a:lvl1pPr',
    'ppt/slides/slide1.xml|a:lstStyle>a:lvl5pPr',
    'ppt/slides/slide1.xml|a:lvl2pPr@marL',
    'ppt/slides/slide1.xml|a:lvl2pPr>a:defRPr',
    'ppt/slides/slide1.xml|a:pPr@lvl',
    'ppt/slides/slide1.xml|a:lnSpc>a:spcPts',
    'ppt/slides/slide1.xml|a:spcBef>a:spcPts',
    'ppt/slides/slide1.xml|a:bodyPr@lIns',
    'ppt/slides/slide1.xml|a:bodyPr@anchor=t',
  ],
  parts: () => pptxPackage([
    header({
      id: 'pptx-001',
      title: 'Paragraph levels',
      verify: [
        'Verify: level n starts on the left edge of column n + 1;',
        'each paragraph advances 36 pt, a 24 pt line after a 12 pt space.',
      ],
    }),
    textBox(3, 'Levels', box, LEVELS.map((lvl) => ({
      own: false,
      lvl,
      text: `lvl ${lvl} · lvl${lvl + 1}pPr · marL ${lvl * PPTX.STEP} · column ${lvl + 1}`,
    })), lstStyle),
    textBox(4, 'Caption', field(1, 6, 4, 4), [
      { style: 'caption', text: `Text box: offset ${box.x}, ${box.y}; extent ${box.cx} by ${box.cy} EMU; insets 0.` },
      { style: 'caption', text: `Margins and spacing come from the box's own list style: ${PPTX.STEP} EMU per level.` },
    ]),
  ]),
};

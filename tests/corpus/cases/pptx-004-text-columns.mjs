import { PPTX } from '../lib/grid.mjs';
import { field, header, pptxPackage, textBox } from '../lib/pptx.mjs';

// Running text is set in the text size, 10/12. With spcCol equal to the grid
// gutter, each text column below is four grid columns wide (about 55-60
// characters): 2 columns in an 8-column box, 3 columns in the full width.
const two = field(1, 8, 2, 2);
const three = { ...field(1, 12, 3, 3), cy: 6 * PPTX.LINE };
const COL = PPTX.w(1, 4);

export default {
  id: 'pptx-004',
  slug: 'text-columns',
  format: 'pptx',
  title: 'Text columns',
  covers: [
    'ppt/slides/slide1.xml|a:bodyPr@numCol',
    'ppt/slides/slide1.xml|a:bodyPr@spcCol',
  ],
  parts: () => pptxPackage([
    header({
      id: 'pptx-004',
      title: 'Text columns',
      verify: [
        'Verify: text columns start on grid columns 1 and 5 (row 2) and 1, 5 and 9 (row 3), each four grid',
        'columns wide; each column fills to the foot of its box before the next begins; nothing overflows.',
      ],
    }),
    textBox(3, 'Two columns', two, [{
      style: 'caption',
      text: 'Running text on a slide is set at 10 point on a 12 point line, the size of the captions; 20 point is '
        + 'kept for single lines and labels. This box is eight grid columns wide and one grid row, nine baselines, '
        + `tall. It holds two text columns of four grid columns each, separated by spcCol ${PPTX.GUTTER} EMU, one `
        + 'gutter. A line of four grid columns carries between fifty and sixty characters, a comfortable measure '
        + 'for continuous reading. PowerPoint does not balance text columns: it fills the first column to the foot '
        + 'of the box and only then continues at the top of the second, which begins on the left edge of grid '
        + 'column 5. These last lines are the ones that have crossed into the second column.',
    }], '', { numCol: 2, spcCol: PPTX.GUTTER }),
    textBox(4, 'Three columns', three, [{
      style: 'caption',
      text: 'This box spans all twelve grid columns and holds numCol 3 with the same spacing, so each text column is '
        + `four grid columns and three gutters wide: ${COL} EMU, exactly the width of grid columns 1 to 4. The second `
        + 'and third text columns therefore begin on the left edges of grid columns 5 and 9. The box is six baselines '
        + 'tall, half a grid row, and the text is long enough to fill the first two columns and to begin the third. '
        + 'Each line sits on the 12 point baseline of the slide grid; with zero insets the first line of every column '
        + 'starts at the top edge of the box. The column width follows from the box: subtract the spacing between '
        + 'the columns from the box width and divide by the number of columns. A wider box at the same numCol widens '
        + 'every column; a larger spcCol widens only the gaps. Nothing on this slide is positioned by eye.',
    }], '', { numCol: 3, spcCol: PPTX.GUTTER }),
    textBox(5, 'Caption', field(1, 6, 4, 4), [
      { style: 'caption', text: 'Column width = (box width − (numCol − 1) × spcCol) ÷ numCol' },
      { style: 'caption', text: `Row 2: (${two.cx} − ${PPTX.GUTTER}) ÷ 2 = ${COL} EMU` },
      { style: 'caption', text: `Row 3: (${three.cx} − 2 × ${PPTX.GUTTER}) ÷ 3 = ${COL} EMU` },
    ]),
  ]),
};

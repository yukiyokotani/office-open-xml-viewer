import { COLOR } from '../lib/grid.mjs';
import { NO_LINE, field, header, line, pptxPackage, shape, solid, textBox } from '../lib/pptx.mjs';

const LINE_W = 12700; // 1 pt

// One shape per field of three columns by two rows: geometry × paint.
const SHAPES = [
  { prst: 'rect', fill: solid(COLOR.ink), ln: NO_LINE, paint: [`solidFill ${COLOR.ink}`, 'ln noFill'] },
  { prst: 'rect', fill: '<a:noFill/>', ln: line(LINE_W), paint: ['noFill', `ln ${LINE_W} EMU, ${COLOR.ink}`] },
  { prst: 'ellipse', fill: '<a:noFill/>', ln: line(LINE_W), paint: ['noFill', `ln ${LINE_W} EMU, ${COLOR.ink}`] },
  { prst: 'ellipse', fill: solid(COLOR.accent), ln: NO_LINE, paint: [`solidFill ${COLOR.accent}`, 'ln noFill'] },
];

export default {
  id: 'pptx-002',
  slug: 'preset-shapes',
  format: 'pptx',
  title: 'Preset shapes',
  covers: [
    'ppt/slides/slide1.xml|a:prstGeom@prst=rect',
    'ppt/slides/slide1.xml|a:prstGeom@prst=ellipse',
    'ppt/slides/slide1.xml|p:spPr>a:solidFill',
    'ppt/slides/slide1.xml|p:spPr>a:noFill',
    'ppt/slides/slide1.xml|a:ln@w',
    'ppt/slides/slide1.xml|a:ln>a:solidFill',
    'ppt/slides/slide1.xml|a:ln>a:noFill',
    'ppt/slides/slide1.xml|a:solidFill>a:srgbClr',
  ],
  parts: () => pptxPackage([
    header({
      id: 'pptx-002',
      title: 'Preset shapes',
      verify: [
        'Verify: each shape fills a field of three columns by two rows;',
        'a 1 pt line is centred on the geometry edge.',
      ],
    }),
    ...SHAPES.flatMap((s, i) => {
      const c1 = 1 + 3 * i;
      const box = field(c1, c1 + 2, 2, 3);
      return [
        shape(3 + 2 * i, `${s.prst} ${i + 1}`, box, s.prst, s.fill, s.ln),
        textBox(4 + 2 * i, `Caption ${i + 1}`, field(c1, c1 + 2, 4, 4), [
          { style: 'caption', bold: true, text: `prstGeom ${s.prst}` },
          ...s.paint.map((text) => ({ style: 'caption', text })),
          { style: 'caption', text: `x ${box.x}  y ${box.y}` },
          { style: 'caption', text: `cx ${box.cx}  cy ${box.cy}` },
        ]),
      ];
    }),
  ]),
};

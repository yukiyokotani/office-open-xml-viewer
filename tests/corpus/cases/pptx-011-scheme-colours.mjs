import { COLOR } from '../lib/grid.mjs';
import { field, header, pptxPackage, shape, textBox } from '../lib/pptx.mjs';

const scheme = (val, mods = '') => `<a:solidFill><a:schemeClr val="${val}">${mods}</a:schemeClr></a:solidFill>`;
// The theme's accent1 is the corpus accent; the theme's fill style 1 is a
// solid placeholder colour (phClr), so a style reference paints with the
// colour given in the reference.
const STYLE_REF = '<p:style><a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>'
  + '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style>';

const SHAPES = [
  { name: 'Style reference', fill: '', style: STYLE_REF, lines: ['p:style fillRef idx 1, accent1', 'fill from the theme fill style list', `resolves to ${COLOR.accent}`] },
  { name: 'Scheme colour', fill: scheme('tx1'), lines: ['solidFill schemeClr tx1', 'clrMap tx1 → dk1', `resolves to ${COLOR.ink}`] },
  { name: 'Shade', fill: scheme('accent1', '<a:lumMod val="50000"/>'), lines: ['schemeClr accent1', 'lumMod 50000', 'half the luminance of the accent'] },
  { name: 'Tint', fill: scheme('accent1', '<a:lumMod val="20000"/><a:lumOff val="80000"/>'), lines: ['schemeClr accent1', 'lumMod 20000, lumOff 80000', 'a light tint of the accent'] },
];

export default {
  id: 'pptx-011',
  slug: 'scheme-colours',
  format: 'pptx',
  title: 'Scheme colours',
  covers: [
    'ppt/slides/slide1.xml|p:sp>p:style',
    'ppt/slides/slide1.xml|p:style>a:fillRef',
    'ppt/slides/slide1.xml|a:fillRef@idx',
    'ppt/slides/slide1.xml|p:style>a:fontRef',
    'ppt/slides/slide1.xml|a:solidFill>a:schemeClr',
    'ppt/slides/slide1.xml|a:schemeClr>a:lumMod',
    'ppt/slides/slide1.xml|a:schemeClr>a:lumOff',
    'ppt/theme/theme1.xml|a:clrScheme>a:accent1',
    'ppt/slidemasters/slidemaster1.xml|p:sldMaster>p:clrMap',
  ],
  parts: () => pptxPackage([
    header({
      id: 'pptx-011',
      title: 'Scheme colours',
      verify: [
        'Verify: from the left, accent orange, black, dark orange and pale orange, each filling three',
        'columns by two rows; every colour is resolved through the theme.',
      ],
    }),
    ...SHAPES.flatMap((s, i) => {
      const c1 = 1 + 3 * i;
      const sp = shape(3 + 2 * i, s.name, field(c1, c1 + 2, 2, 3), 'rect', s.fill, '<a:ln><a:noFill/></a:ln>');
      return [
        s.style ? sp.replace('</p:spPr></p:sp>', `</p:spPr>${s.style}</p:sp>`) : sp,
        textBox(4 + 2 * i, `Caption ${i + 1}`, field(c1, c1 + 2, 4, 4), s.lines.map((text, j) => ({ style: 'caption', bold: j === 0, text }))),
      ];
    }),
  ]),
};

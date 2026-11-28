import { COLOR, OUTLINED_INSET } from '../lib/grid.mjs';
import { field, header, paragraph, pptxPackage, xfrm } from '../lib/pptx.mjs';
import { attrs } from '../lib/xml.mjs';

const HAIR = `<a:ln w="6350"><a:solidFill><a:srgbClr val="${COLOR.ink}"/></a:solidFill></a:ln>`;
// Each field is one outlined rectangle that holds its own text, inset by one
// gutter (OUTLINED_INSET) so the text keeps clear of the hairline.
const I = OUTLINED_INSET;
const box = (id, name, b, bodyAttrs, fit, paragraphs) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>`
  + `<p:spPr>${xfrm(b)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>${HAIR}</p:spPr>`
  + `<p:txBody><a:bodyPr${attrs({ wrap: 'square', lIns: I, tIns: I, rIns: I, bIns: I, ...bodyAttrs })}>${fit}</a:bodyPr><a:lstStyle/>${paragraphs.map(paragraph).join('')}</p:txBody></p:sp>`;

const CASES = [
  ['anchor t', { anchor: 't' }, '<a:noAutofit/>', [{ style: 'caption', text: 'anchor t: the first line at the top of the field.' }]],
  ['anchor ctr', { anchor: 'ctr' }, '<a:noAutofit/>', [{ style: 'caption', text: 'anchor ctr: the line centred in the field.' }]],
  ['anchor b', { anchor: 'b' }, '<a:noAutofit/>', [{ style: 'caption', text: 'anchor b: the last line at the foot of the field.' }]],
  ['normAutofit', { anchor: 't' }, '<a:normAutofit fontScale="50000" lnSpcReduction="0"/>', [{
    text: 'normAutofit fontScale 50 %: this 20 pt text is drawn at 10 pt, as PowerPoint stored it when the text '
      + 'outgrew the field. A renderer applies the stored scale; it does not compute its own.',
  }]],
];

export default {
  id: 'pptx-014',
  slug: 'anchoring-autofit',
  format: 'pptx',
  title: 'Anchoring and autofit',
  covers: [
    'ppt/slides/slide1.xml|a:bodyPr@anchor=ctr',
    'ppt/slides/slide1.xml|a:bodyPr@anchor=b',
    'ppt/slides/slide1.xml|a:bodyPr>a:normAutofit',
    'ppt/slides/slide1.xml|a:normAutofit@fontScale',
    'ppt/slides/slide1.xml|a:normAutofit@lnSpcReduction',
  ],
  parts: () => pptxPackage([
    header({
      id: 'pptx-014',
      title: 'Anchoring and autofit',
      verify: [
        'Verify: text 12 pt inside the top, middle and foot of the first three outlined fields;',
        'in the fourth, 20 pt text drawn at 10 pt by the stored fontScale.',
      ],
    }),
    ...CASES.flatMap(([name, bodyAttrs, fit, paragraphs], i) => {
      const f = field(1 + 3 * i, 3 + 3 * i, 2, 3);
      return [box(3 + i, name, f, bodyAttrs, fit, paragraphs)];
    }),
  ]),
};

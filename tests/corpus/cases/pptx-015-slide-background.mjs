import { field, header, pptxPackage, shape, textBox } from '../lib/pptx.mjs';

// The slide refers to the theme background style 1001 painted in bg1, and
// overrides the master colour map so that bg1 is dark 1 and tx1 is light 1:
// black background, white text, without a single colour on the slide itself.
const BG = '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>';
const OVERRIDE = '<a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" '
  + 'accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';

export default {
  id: 'pptx-015',
  slug: 'slide-background',
  format: 'pptx',
  title: 'Slide background',
  covers: [
    'ppt/slides/slide1.xml|p:cSld>p:bg',
    'ppt/slides/slide1.xml|p:bg>p:bgRef',
    'ppt/slides/slide1.xml|p:bgRef@idx',
    'ppt/slides/slide1.xml|p:clrMapOvr>a:overrideClrMapping',
    'ppt/slides/slide1.xml|a:overrideClrMapping@bg1',
    'ppt/theme/theme1.xml|a:fmtScheme>a:bgFillStyleLst',
  ],
  parts: () => pptxPackage([
    header({
      id: 'pptx-015',
      title: 'Slide background',
      verify: [
        'Verify: black background, white text and a white field in columns 1-3 of rows 2-3,',
        'a cyan field in columns 4-6; no colour is written on the slide, only theme references.',
      ],
    }),
    shape(3, 'bg2 field', field(1, 3, 2, 3), 'rect', '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>', '<a:ln><a:noFill/></a:ln>'),
    shape(4, 'accent field', field(4, 6, 2, 3), 'rect', '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>', '<a:ln><a:noFill/></a:ln>'),
    textBox(5, 'Caption', field(7, 12, 2, 2), [
      { style: 'caption', bold: true, text: 'p:bg / p:bgRef idx 1001, schemeClr bg1' },
      { style: 'caption', text: 'Background style 1 of the theme (bgFillStyleLst) painted in bg1.' },
      { style: 'caption', text: 'a:overrideClrMapping bg1 → dk1, tx1 → lt1: bg1 is black, tx1 white.' },
      { style: 'caption', text: 'Left field: schemeClr tx1 (white). Right field: accent1 (cyan).' },
    ]),
  ], { slideBg: BG, clrMapOvr: OVERRIDE }),
};

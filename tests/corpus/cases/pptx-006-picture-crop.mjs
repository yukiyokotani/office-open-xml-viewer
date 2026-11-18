import { COLOR, PPTX } from '../lib/grid.mjs';
import { png } from '../lib/png.mjs';
import { NO_LINE, field, header, pptxPackage, textBox, xfrm } from '../lib/pptx.mjs';

// The picture draws the grid it covers at 1 px = 1 pt: three columns by two
// rows, with the second field of the first row filled in the accent colour.
const PT = 12700;
const box = field(1, 3, 2, 3);
const W = box.cx / PT;
const H = box.cy / PT;
const COL = PPTX.COL / PT;
const GUT = PPTX.GUTTER / PT;
const ROW = PPTX.ROW / PT;
const fieldRect = (c, r) => [c * (COL + GUT), r * (ROW + GUT), COL, ROW];
const outline = ([x, y, w, h]) => [[x, y, w, 1, 1], [x, y + h - 1, w, 1, 1], [x, y, 1, h, 1], [x + w - 1, y, 1, h, 1]];
const rects = [];
for (let r = 0; r < 2; r += 1) for (let c = 0; c < 3; c += 1) rects.push(...outline(fieldRect(c, r)));
const [ax, ay, aw, ah] = fieldRect(1, 0);
rects.push([ax, ay, aw, ah, 2]);
const PICTURE = png(Math.round(W), Math.round(H), [COLOR.paper, COLOR.ink, COLOR.accent], rects.map((v) => v.map(Math.round)));

// srcRect crops to the accent field, in thousandths of a percent per edge.
const crop = { l: Math.round((ax / W) * 100000), t: 0, r: Math.round(((W - ax - aw) / W) * 100000), b: Math.round(((H - ah) / H) * 100000) };

const pic = (id, name, frame, srcRect) => `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${name}"/>`
  + '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
  + `<p:blipFill><a:blip r:embed="rId2"/>${srcRect ? `<a:srcRect l="${srcRect.l}" t="${srcRect.t}" r="${srcRect.r}" b="${srcRect.b}"/>` : ''}`
  + '<a:stretch><a:fillRect/></a:stretch></p:blipFill>'
  + `<p:spPr>${xfrm(frame)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${NO_LINE}</p:spPr></p:pic>`;

export default {
  id: 'pptx-006',
  slug: 'picture-crop',
  format: 'pptx',
  title: 'Picture crop',
  covers: [
    'ppt/slides/slide1.xml|p:spTree>p:pic',
    'ppt/slides/slide1.xml|p:blipFill>a:blip',
    'ppt/slides/slide1.xml|a:blip@r:embed',
    'ppt/slides/slide1.xml|p:blipFill>a:srcRect',
    'ppt/slides/slide1.xml|a:srcRect@l',
    'ppt/slides/slide1.xml|a:srcRect@b',
    'ppt/slides/slide1.xml|a:stretch>a:fillRect',
    'ppt/slides/slide1.xml|a:picLocks@noChangeAspect',
  ],
  parts: () => pptxPackage([
    header({
      id: 'pptx-006',
      title: 'Picture crop',
      verify: [
        'Verify: the left picture shows six outlined fields over columns 1-3, rows 2-3;',
        'the right picture is cropped to the orange field and fills column 5, row 2.',
      ],
    }),
    pic(3, 'Uncropped', box),
    pic(4, 'Cropped', field(5, 5, 2, 2), crop),
    textBox(5, 'Caption', field(1, 6, 4, 4), [
      { style: 'caption', text: `Picture: ${Math.round(W)} × ${Math.round(H)} px, 1 px = 1 pt, written by the corpus builder.` },
      { style: 'caption', text: `a:srcRect l ${crop.l}, t ${crop.t}, r ${crop.r}, b ${crop.b} (thousandths of a percent of the picture).` },
    ]),
  ], {
    parts: { 'ppt/media/image1.png': PICTURE },
    rels: [['rId2', 'image', '../media/image1.png']],
    defaults: { png: 'image/png' },
  }),
};

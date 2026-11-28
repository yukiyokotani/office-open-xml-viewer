import { DOCX } from '../lib/grid.mjs';
import { A_NS, PIC_NS, WP_NS, anchoredPicture, inlinePicture } from '../lib/drawing.mjs';
import { R_NS, docxPackage, header, run } from '../lib/docx.mjs';
import { gridPicture } from '../lib/png.mjs';

const at = DOCX.rowLine;
const { emu, colStart, span, lineTop, LINE } = DOCX;
// 1 px = 20 twips (1 pt) in both pictures: each draws the grid fields it covers.
const PX = 20;

// Anchored: columns 3-4 of row 2 (9 lines), wrapped by text in columns 1-2.
const anchoredBox = { cols: 2, rows: 1, field: span(3, 3) / PX, gutterX: DOCX.gutterAfter(3) / PX, fieldH: 9 * LINE / PX, gutterY: 0 };
// Inline: column 1, eight lines.
const inlineBox = { cols: 1, rows: 1, field: span(1, 1) / PX, gutterX: 0, fieldH: 8 * LINE / PX, gutterY: 0 };
const anchoredPng = gridPicture({ ...anchoredBox, field: Math.round(anchoredBox.field), gutterX: Math.round(anchoredBox.gutterX) });
const inlinePng = gridPicture({ ...inlineBox, field: Math.round(inlineBox.field) });

const ANCHOR = {
  id: 1, name: 'Grid columns 3-4', rId: 'rId3',
  x: emu(colStart(3)), y: emu(lineTop(at(2))), cx: emu(span(3, 4)), cy: emu(9 * LINE),
  distL: emu(DOCX.gutterAfter(2)),
};
const INLINE = { id: 2, name: 'Grid column 1', rId: 'rId4', cx: emu(span(1, 1)), cy: emu(8 * LINE) };

export default {
  id: 'docx-010',
  slug: 'pictures',
  format: 'docx',
  title: 'Pictures',
  covers: [
    'word/document.xml|w:drawing>wp:anchor',
    'word/document.xml|wp:anchor>wp:wrapSquare',
    'word/document.xml|wp:positionH@relativeFrom=margin',
    'word/document.xml|wp:positionH>wp:posOffset',
    'word/document.xml|wp:anchor@distL',
    'word/document.xml|w:drawing>wp:inline',
    'word/document.xml|wp:inline>wp:extent',
    'word/document.xml|pic:blipFill>a:blip',
    'word/document.xml|a:blip@r:embed',
    'word/document.xml|a:stretch>a:fillRect',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-010',
      title: 'Pictures',
      verify: [
        'Verify: the anchored picture fills columns 3-4 of row 2,',
        'text wraps one gutter from it; the inline picture fills',
        'column 1 for eight lines. Pictures draw their own fields.',
      ],
    }),
    {
      kind: 'p',
      at: at(2),
      runs: [
        anchoredPicture(ANCHOR),
        run('The picture to the right is anchored to the margin: wp:positionH and wp:positionV place its top left '
          + `corner at ${ANCHOR.x} and ${ANCHOR.y} EMU from the margin, the left edge of grid column 3 and the top `
          + `of grid row 2. Its extent, ${ANCHOR.cx} by ${ANCHOR.cy} EMU, is two grid columns with their gutter `
          + 'by nine baselines. Square wrapping (wp:wrapSquare) lets this paragraph run beside it; the distance '
          + `wp:anchor distL of ${ANCHOR.distL} EMU is one gutter, so the lines stop at the right edge of column 2. `
          + 'Below the picture the text takes the full measure again. The picture itself is a PNG written by the '
          + 'corpus builder: two outlined fields and the gutter between them, drawn at one point '
          + 'per pixel, so the picture shows the part of the grid it covers.'),
      ],
    },
    { kind: 'p', before: 1, style: 'caption', bold: true, lines: ['Inline picture: wp:inline, one grid column by eight baselines'] },
    { kind: 'p', lineAuto: true, runs: [inlinePicture(INLINE)], height: 8 },
  ], {
    ns: [R_NS, WP_NS, A_NS, PIC_NS],
    parts: { 'word/media/image1.png': anchoredPng.bytes, 'word/media/image2.png': inlinePng.bytes },
    rels: [['rId3', 'image', 'media/image1.png'], ['rId4', 'image', 'media/image2.png']],
    defaults: { png: 'image/png' },
  }),
};

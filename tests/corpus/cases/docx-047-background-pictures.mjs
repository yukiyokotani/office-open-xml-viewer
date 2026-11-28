import { COLOR, DOCX } from '../lib/grid.mjs';
import { A_NS, PIC_NS, WP_NS, inlinePicture, pictureGraphic } from '../lib/drawing.mjs';
import { R_NS, W, WML, docxPackage, header } from '../lib/docx.mjs';
import { png } from '../lib/png.mjs';
import { tiff } from '../lib/tiff.mjs';
import { XML_DECL } from '../lib/xml.mjs';

// The page background and pictures outside the usual inline PNG:
//   A. A document background colour (w:background, shown because
//      w:displayBackgroundShape is set): the page is the accent to every
//      edge.
//   B. Pictures in the header and the footer, anchored to the page behind
//      the text: a black band across the top and across the foot of the
//      page.
//   C. A picture written as legacy VML (w:pict, v:shape with v:imagedata).
//   D. A picture stored as TIFF (uncompressed RGB).
const { emu, PAGE_W, LINE, span } = DOCX;
const band = png(PAGE_W / 20, 8, [COLOR.ink], []);
const BOX_W = span(1, 1);
const pngField = (png(64, 64, [COLOR.paper, COLOR.ink], [[0, 0, 64, 4, 1], [0, 60, 64, 4, 1], [0, 0, 4, 64, 1], [60, 0, 4, 64, 1], [16, 16, 32, 32, 1]]));
const tiffField = tiff(64, 64, [COLOR.paper, COLOR.ink], [[0, 0, 64, 4, 1], [0, 60, 64, 4, 1], [0, 0, 4, 64, 1], [60, 0, 4, 64, 1], [24, 8, 16, 48, 1]]);

const anchoredBand = (id, rId, top) => '<w:r><w:drawing>'
  + '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">'
  + '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>'
  + `<wp:positionV relativeFrom="page"><wp:align>${top ? 'top' : 'bottom'}</wp:align></wp:positionV>`
  + `<wp:extent cx="${emu(PAGE_W)}" cy="${emu(2 * LINE)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>`
  + `<wp:docPr id="${id}" name="Band ${id}"/>${pictureGraphic(id, `Band ${id}`, rId, emu(PAGE_W), emu(2 * LINE))}`
  + '</wp:anchor></w:drawing></w:r>';
const ns = `${W} ${R_NS} ${WP_NS} ${A_NS} ${PIC_NS}`;
const HEADER = XML_DECL + `<w:hdr ${ns}><w:p>${anchoredBand(10, 'rId1', true)}</w:p></w:hdr>`;
const FOOTER = XML_DECL + `<w:ftr ${ns}><w:p>${anchoredBand(11, 'rId1', false)}</w:p></w:ftr>`;
const partRels = (target) => XML_DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
  + `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${target}"/></Relationships>`;

const pt = (twips) => twips / 20;
const VML = `<w:r><w:pict><v:shape id="vml1" type="#_x0000_t75" style="width:${pt(BOX_W)}pt;height:${pt(BOX_W)}pt">`
  + '<v:imagedata r:id="rId6" o:title="VML picture"/></v:shape></w:pict></w:r>';
const caption = (text, extra = {}) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1, ...extra });
const picPara = (run) => ({ kind: 'raw', height: NaN, xml: `<w:p><w:pPr><w:spacing w:line="${LINE}" w:lineRule="auto"/></w:pPr>${run}</w:p>` });

export default {
  id: 'docx-047',
  slug: 'background-pictures',
  format: 'docx',
  title: 'Page background and pictures',
  covers: [
    'word/document.xml|w:document>w:background',
    'word/settings.xml|w:settings>w:displayBackgroundShape',
    'word/header1.xml|w:drawing>wp:anchor',
    'word/footer1.xml|w:drawing>wp:anchor',
    'word/document.xml|w:pict>v:shape',
    'word/document.xml|v:shape>v:imagedata',
    'word/media/|*.tiff',
  ],
  parts: () => {
    const pkg = docxPackage([
      ...header({
        id: 'docx-047',
        title: 'Page background and pictures',
        verify: [
          'Verify: the page is the accent to every edge, with a',
          'black band at top and foot; C, drawn from VML, and D,',
          'a TIFF, show a framed square and a framed bar.',
        ],
      }),
      caption('C. Legacy VML picture (w:pict, v:imagedata)', { at: DOCX.rowLine(2), before: 0 }),
      picPara(VML),
      caption('D. TIFF picture'),
      picPara(inlinePicture({ id: 2, name: 'TIFF picture', rId: 'rId7', cx: emu(BOX_W), cy: emu(BOX_W) })),
    ], {
      firstSection: {
        refs: '<w:headerReference w:type="default" r:id="rId3"/><w:footerReference w:type="default" r:id="rId4"/>',
        headerDist: DOCX.HEADER_DIST,
        footerDist: DOCX.FOOTER_DIST,
      },
      settings: '<w:displayBackgroundShape/>',
      parts: {
        'word/header1.xml': HEADER,
        'word/footer1.xml': FOOTER,
        'word/_rels/header1.xml.rels': partRels('media/image1.png'),
        'word/_rels/footer1.xml.rels': partRels('media/image1.png'),
        'word/media/image1.png': band,
        'word/media/image2.png': pngField,
        'word/media/image3.tiff': tiffField,
      },
      rels: [['rId3', 'header', 'header1.xml'], ['rId4', 'footer', 'footer1.xml'], ['rId6', 'image', 'media/image2.png'], ['rId7', 'image', 'media/image3.tiff']],
      overrides: { 'word/header1.xml': `${WML}.header+xml`, 'word/footer1.xml': `${WML}.footer+xml` },
      defaults: { png: 'image/png', tiff: 'image/tiff' },
      ns: [R_NS, WP_NS, A_NS, PIC_NS, 'xmlns:v="urn:schemas-microsoft-com:vml"', 'xmlns:o="urn:schemas-microsoft-com:office:office"'],
    });
    // The background is the first child of w:document, before w:body.
    pkg['word/document.xml'] = pkg['word/document.xml'].replace('<w:body>', `<w:background w:color="${COLOR.accent}"/><w:body>`);
    return pkg;
  },
};

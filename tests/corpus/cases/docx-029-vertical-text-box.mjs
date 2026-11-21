import { DOCX, FONT } from '../lib/grid.mjs';
import { A_NS } from '../lib/drawing.mjs';
import { docxPackage, header, run } from '../lib/docx.mjs';

// Vertical Japanese inside a text box on a horizontal page, as notices and
// newsletters set it: wps:bodyPr vert="eaVert". Punctuation and brackets take
// their vertical forms; two-digit date numbers stand upright side by side
// (w:eastAsianLayout vert, vertCompress); an e-mail address in half-width
// letters lies on its side.
const { emu, colStart, span, LINE } = DOCX;
const X = colStart(3);
const Y = DOCX.lineTop(DOCX.rowLine(2));
const W = span(3, 4);
const H = 27 * LINE;
const JA = `<w:rFonts w:ascii="${FONT.eastAsia}" w:eastAsia="${FONT.eastAsia}" w:hAnsi="${FONT.eastAsia}"/>`;
const ja = (text) => run(text, `${JA}<w:sz w:val="24"/>`);
const tcy = (digits, id) => run(digits, `${JA}<w:sz w:val="24"/><w:eastAsianLayout w:id="${id}" w:vert="1" w:vertCompress="1"/>`);
const latin = (text) => run(text, '<w:sz w:val="24"/>');
const para = (...runs) => `<w:p><w:pPr><w:spacing w:line="${2 * LINE}" w:lineRule="exact"/></w:pPr>${runs.join('')}</w:p>`;
const CONTENT = '<w:txbxContent>'
  + para(ja('子ども会のお知らせ（第五号）'))
  + para(ja('運動会は'), tcy('10', 1), ja('月'), tcy('12', 2), ja('日に開かれる。雨の場合は「延期」とし、翌週に行う。'))
  + para(ja('持ち物は、水筒・帽子・タオルなど……。集合はグラウンドの東側、ちょっと早めに。'))
  + para(ja('問い合わせは'), latin('info@example.org'), ja('まで。'))
  + '</w:txbxContent>';

const BOX = '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>'
  + '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
  + `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>${emu(X)}</wp:posOffset></wp:positionH>`
  + `<wp:positionV relativeFrom="margin"><wp:posOffset>${emu(Y)}</wp:posOffset></wp:positionV>`
  + `<wp:extent cx="${emu(W)}" cy="${emu(H)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>`
  + '<wp:docPr id="1" name="Vertical text box"/><wp:cNvGraphicFramePr/>'
  + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>'
  + `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(W)}" cy="${emu(H)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>`
  + `<wps:txbx>${CONTENT}</wps:txbx>`
  + '<wps:bodyPr rot="0" vert="eaVert" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
  + '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback/></mc:AlternateContent></w:r>';

export default {
  id: 'docx-029',
  slug: 'vertical-text-box',
  format: 'docx',
  title: 'Vertical text box',
  covers: [
    'word/document.xml|wps:bodyPr@vert=eaVert',
    'word/document.xml|wps:wsp>wps:txbx',
    'word/document.xml|w:eastAsianLayout@w:vertCompress',
    'word/document.xml|mc:AlternateContent>mc:Choice',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-029',
      title: 'Vertical text box',
      verify: [
        'Verify: columns 3-4 hold four vertical lines from right to left;',
        '「」（）、。…・ in vertical forms, 10 and 12 upright side by',
        'side, the e-mail address turned 90° clockwise.',
      ],
    }),
    {
      kind: 'p',
      at: DOCX.rowLine(2),
      cols: [1, 2],
      runs: [BOX, run('The text box to the right is anchored to the margin over grid columns 3 and 4, rows 2 to 4. Its body '
        + 'is set vertically (wps:bodyPr vert="eaVert"): lines run top to bottom and follow each other from right to left. '
        + 'The page itself stays horizontal, so this caption runs across the page as usual.')],
    },
  ], {
    ns: [A_NS,
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'],
  }),
};

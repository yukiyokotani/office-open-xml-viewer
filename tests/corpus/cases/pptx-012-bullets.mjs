import { COLOR, FONT, PPTX } from '../lib/grid.mjs';
import { bodyPr, field, header, pptxPackage, textBox, xfrm } from '../lib/pptx.mjs';

// Bullets hang one column module: the bullet on a column edge, the text on the
// next. Level 1 spaces lines with spcPts, level 2 with spcPct 120 % (24 pt for
// 20 pt type); both land on the 12 pt baseline.
const T = PPTX.TYPE.text;
const lvl = (n, spacing, bullet) => `<a:lvl${n}pPr marL="${n * PPTX.STEP}" indent="${-PPTX.STEP}">${spacing}`
  + `<a:spcBef><a:spcPts val="0"/></a:spcBef>${bullet}<a:defRPr sz="${T.sz}"/></a:lvl${n}pPr>`;
const DASH = `<a:buFont typeface="${FONT.latin}"/><a:buChar char="–"/>`;
const TX_STYLES = '<p:txStyles>'
  + `<p:titleStyle><a:lvl1pPr><a:lnSpc><a:spcPts val="${PPTX.TYPE.title.leading}"/></a:lnSpc><a:buNone/><a:defRPr sz="${PPTX.TYPE.title.sz}" b="1"/></a:lvl1pPr></p:titleStyle>`
  + `<p:bodyStyle>${lvl(1, `<a:lnSpc><a:spcPts val="${T.leading}"/></a:lnSpc>`, `<a:buClr><a:srgbClr val="${COLOR.accent}"/></a:buClr><a:buSzPct val="100000"/>${DASH}`)}`
  + `${lvl(2, '<a:lnSpc><a:spcPct val="120000"/></a:lnSpc>', `<a:buClrTx/><a:buSzTx/>${DASH}`)}</p:bodyStyle>`
  + `<p:otherStyle><a:lvl1pPr><a:defRPr sz="${T.sz}"/></a:lvl1pPr></p:otherStyle></p:txStyles>`;

const BODY = field(1, 6, 2, 3);
const masterBody = '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>'
  + `<p:spPr>${xfrm(BODY)}</p:spPr><p:txBody>${bodyPr()}<a:lstStyle/><a:p><a:endParaRPr lang="en-GB"/></a:p></p:txBody></p:sp>`;
const layoutBody = '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>'
  + '<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-GB"/></a:p></p:txBody></p:sp>';

const item = (lvlN, text) => `<a:p>${lvlN ? `<a:pPr lvl="${lvlN}"/>` : ''}<a:r><a:rPr lang="en-GB"/><a:t>${text}</a:t></a:r></a:p>`;
const slideBody = '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>'
  + '<p:txBody><a:bodyPr/><a:lstStyle/>'
  + item(0, 'Level 1: orange dash, buClr accent')
  + item(1, 'Level 2: dash in the text colour')
  + item(1, 'Line spacing spcPct 120 %')
  + item(0, 'Bullets from the master body style')
  + '</p:txBody></p:sp>';

const numbered = (n, text) => ({
  xml: `<a:p><a:pPr marL="${PPTX.STEP}" indent="${-PPTX.STEP}"><a:lnSpc><a:spcPts val="${T.leading}"/></a:lnSpc>`
    + `<a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/></a:pPr>`
    + `<a:r><a:rPr lang="en-GB" sz="${T.sz}"/><a:t>${text}</a:t></a:r></a:p>`,
});

export default {
  id: 'pptx-012',
  slug: 'bullets',
  format: 'pptx',
  title: 'Bullets and numbering',
  covers: [
    'ppt/slidemasters/slidemaster1.xml|a:lvl1pPr>a:buChar',
    'ppt/slidemasters/slidemaster1.xml|a:lvl1pPr>a:buFont',
    'ppt/slidemasters/slidemaster1.xml|a:lvl1pPr>a:buClr',
    'ppt/slidemasters/slidemaster1.xml|a:lvl1pPr>a:buSzPct',
    'ppt/slidemasters/slidemaster1.xml|a:lvl2pPr>a:buClrTx',
    'ppt/slidemasters/slidemaster1.xml|a:lvl2pPr>a:buSzTx',
    'ppt/slidemasters/slidemaster1.xml|a:lnSpc>a:spcPct',
    'ppt/slidemasters/slidemaster1.xml|a:lvl1pPr>a:spcBef',
    'ppt/slidemasters/slidemaster1.xml|a:lvl1pPr@indent',
    'ppt/slides/slide1.xml|a:pPr>a:buAutoNum',
    'ppt/slides/slide1.xml|a:buAutoNum@type=arabicPeriod',
  ],
  parts: () => pptxPackage([
    header({
      id: 'pptx-012',
      title: 'Bullets',
      verify: [
        'Verify: left, master bullets on columns 1 and 2 with text one column in, level 1 dash orange;',
        'right, numbers 1-3 on column 7 with text on column 8. All lines on the 24 pt pitch.',
      ],
    }),
    slideBody,
    textBox(4, 'Numbered', field(7, 12, 2, 3), [
      numbered(1, 'buAutoNum arabicPeriod'),
      numbered(2, 'Numbers count the paragraphs'),
      numbered(3, 'marL one module, indent minus one'),
    ]),
    textBox(5, 'Caption', field(1, 6, 4, 4), [
      { style: 'caption', text: `Master body style: level n has marL n × ${PPTX.STEP} and indent −${PPTX.STEP} EMU, so the bullet hangs one module.` },
      { style: 'caption', text: 'Level 1: lnSpc spcPts 2400, buClr accent, buSzPct 100 %. Level 2: lnSpc spcPct 120 %, buClrTx, buSzTx.' },
    ]),
  ], {
    master: { shapes: masterBody, txStyles: TX_STYLES },
    layout: { type: 'obj', name: 'Body', shapes: layoutBody },
  }),
};

import { FONT, PPTX } from '../lib/grid.mjs';
import { bodyPr, field, header, pptxPackage, xfrm } from '../lib/pptx.mjs';

// The master body style defines all nine levels; the slide gives only text and
// a level per paragraph. Each level hangs its bullet one column module left of
// its text, so level n puts the bullet on column n and the text on column n+1.
// spcBef of one baseline keeps every paragraph on three 12 pt baselines.
const LEVELS = [1, 2, 3, 4, 5, 6, 7, 8, 9];
const bullet = (n) => (n % 2 ? '•' : '–');
const lvl = (n, type, body) => `<a:lvl${n}pPr marL="${body ? n * PPTX.STEP : 0}" indent="${body ? -PPTX.STEP : 0}" algn="l"`
  + ` defTabSz="${PPTX.STEP}" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">`
  + `<a:lnSpc><a:spcPts val="${type.leading}"/></a:lnSpc>`
  + `<a:spcBef><a:spcPts val="${body ? 1200 : 0}"/></a:spcBef>`
  + (body ? `<a:buFont typeface="${FONT.latin}"/><a:buChar char="${bullet(n)}"/>` : '<a:buNone/>')
  + `<a:defRPr sz="${type.sz}" kern="0"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>`
  + '<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>'
  + `</a:lvl${n}pPr>`;
const TX_STYLES = '<p:txStyles>'
  + `<p:titleStyle>${lvl(1, PPTX.TYPE.title, false)}</p:titleStyle>`
  + `<p:bodyStyle>${LEVELS.map((n) => lvl(n, PPTX.TYPE.text, true)).join('')}</p:bodyStyle>`
  + `<p:otherStyle>${LEVELS.map((n) => lvl(n, PPTX.TYPE.text, false)).join('')}</p:otherStyle>`
  + '</p:txStyles>';

const BODY = field(1, 12, 2, 4);
const ph = (id, name, phXml, box) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/>`
  + `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${phXml}</p:nvPr></p:nvSpPr>`
  + `<p:spPr>${box ? xfrm(box) : ''}</p:spPr>`
  + `<p:txBody>${box ? bodyPr() : '<a:bodyPr/>'}<a:lstStyle/><a:p><a:endParaRPr lang="en-GB"/></a:p></p:txBody></p:sp>`;

const slideBody = `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>`
  + '<p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>'
  + LEVELS.map((n) => `<a:p>${n > 1 ? `<a:pPr lvl="${n - 1}"/>` : ''}`
    + `<a:r><a:rPr lang="en-GB"/><a:t>Level ${n}, column ${n + 1}</a:t></a:r></a:p>`).join('')
  + '</p:txBody></p:sp>';

export default {
  id: 'pptx-018',
  slug: 'master-body-levels',
  format: 'pptx',
  title: 'Master body levels',
  covers: [
    'ppt/slidemasters/slidemaster1.xml|p:bodyStyle>a:lvl3pPr',
    'ppt/slidemasters/slidemaster1.xml|p:bodyStyle>a:lvl9pPr',
    'ppt/slidemasters/slidemaster1.xml|a:lvl3pPr@marL',
    'ppt/slidemasters/slidemaster1.xml|a:lvl3pPr@indent',
    'ppt/slidemasters/slidemaster1.xml|a:lvl3pPr@algn',
    'ppt/slidemasters/slidemaster1.xml|a:lvl3pPr>a:spcBef',
    'ppt/slidemasters/slidemaster1.xml|a:lvl3pPr>a:buChar',
    'ppt/slidemasters/slidemaster1.xml|a:lvl3pPr>a:buFont',
    'ppt/slidemasters/slidemaster1.xml|a:lvl3pPr>a:defRPr',
    'ppt/slidemasters/slidemaster1.xml|a:defRPr>a:solidFill',
    'ppt/slidemasters/slidemaster1.xml|a:defRPr>a:ea',
    'ppt/slidemasters/slidemaster1.xml|a:defRPr>a:cs',
    'ppt/slidemasters/slidemaster1.xml|p:otherStyle>a:lvl2pPr',
    'ppt/slidemasters/slidemaster1.xml|a:lvl1pPr@defTabSz',
    'ppt/slidemasters/slidemaster1.xml|a:lvl1pPr@eaLnBrk',
    'ppt/slidemasters/slidemaster1.xml|a:lvl1pPr@hangingPunct',
    'ppt/slides/slide1.xml|a:pPr@lvl',
  ],
  parts: () => pptxPackage([
    header({
      id: 'pptx-018',
      title: 'Master body levels',
      verify: [
        'Verify: nine 20 pt paragraphs on rows 2-4; level n hangs its bullet on column n',
        'and starts its text on column n+1. Bullets alternate • and –; one blank baseline between.',
      ],
    }),
    slideBody,
  ], {
    master: {
      shapes: ph(3, 'Body', '<p:ph type="body" idx="1"/>', BODY),
      txStyles: TX_STYLES,
    },
    layout: {
      type: 'obj',
      name: 'Body',
      shapes: ph(3, 'Body', '<p:ph idx="1"/>'),
    },
  }),
};

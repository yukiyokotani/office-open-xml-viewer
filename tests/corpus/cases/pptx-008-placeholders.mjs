import { PPTX } from '../lib/grid.mjs';
import { bodyPr, field, header, pptxPackage, textBox, xfrm } from '../lib/pptx.mjs';

// The master places the title and body placeholders and styles their text;
// the layout moves the body; the slide gives only text. Position, size and
// type all come down the chain master → layout → slide.
const lvl = (n, type, extra = '') => `<a:lvl${n}pPr marL="${(n - 1) * PPTX.STEP}" indent="0">`
  + `<a:lnSpc><a:spcPts val="${type.leading}"/></a:lnSpc><a:buNone/>`
  + `<a:defRPr sz="${type.sz}"${extra}><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl${n}pPr>`;
const TX_STYLES = '<p:txStyles>'
  + `<p:titleStyle>${lvl(1, PPTX.TYPE.title, ' b="1"')}</p:titleStyle>`
  + `<p:bodyStyle>${lvl(1, PPTX.TYPE.text)}${lvl(2, PPTX.TYPE.text)}</p:bodyStyle>`
  + `<p:otherStyle>${lvl(1, PPTX.TYPE.text)}</p:otherStyle>`
  + '</p:txStyles>';

const ph = (id, name, phXml, box) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/>`
  + `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${phXml}</p:nvPr></p:nvSpPr>`
  + `<p:spPr>${box ? xfrm(box) : ''}</p:spPr>`
  + `<p:txBody>${box ? bodyPr() : '<a:bodyPr/>'}<a:lstStyle/><a:p><a:endParaRPr lang="en-GB"/></a:p></p:txBody></p:sp>`;

const MASTER_TITLE = field(1, 12, 2, 2);
const MASTER_BODY = field(1, 12, 3, 4);
const LAYOUT_BODY = field(1, 6, 3, 4);

const slidePh = (id, name, phXml, paragraphs) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/>`
  + `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${phXml}</p:nvPr></p:nvSpPr><p:spPr/>`
  + `<p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs.map(([lvlN, text]) => `<a:p>${lvlN ? `<a:pPr lvl="${lvlN}"/>` : ''}`
    + `<a:r><a:rPr lang="en-GB"/><a:t>${text}</a:t></a:r></a:p>`).join('')}</p:txBody></p:sp>`;

export default {
  id: 'pptx-008',
  slug: 'placeholders',
  format: 'pptx',
  title: 'Placeholders',
  covers: [
    'ppt/slides/slide1.xml|p:nvPr>p:ph',
    'ppt/slides/slide1.xml|p:ph@type=title',
    'ppt/slides/slide1.xml|p:ph@idx',
    'ppt/slidelayouts/slidelayout1.xml|p:nvPr>p:ph',
    'ppt/slidemasters/slidemaster1.xml|p:sldMaster>p:txStyles',
    'ppt/slidemasters/slidemaster1.xml|p:txStyles>p:titleStyle',
    'ppt/slidemasters/slidemaster1.xml|p:txStyles>p:bodyStyle',
    'ppt/slidemasters/slidemaster1.xml|a:lvl1pPr>a:buNone',
  ],
  parts: () => pptxPackage([
    header({
      id: 'pptx-008',
      title: 'Placeholders',
      verify: [
        'Verify: the title fills row 2 and the body columns 1-6 of rows 3-4, though the slide',
        'gives neither position nor size; body level 1 starts on column 2.',
      ],
    }),
    slidePh(3, 'Title', '<p:ph type="title"/>', [[0, 'Title placeholder, 40/48 pt from the master']]),
    slidePh(4, 'Body', '<p:ph idx="1"/>', [
      [0, 'Body placeholder idx 1. The master puts it across all twelve columns;'],
      [0, 'the layout narrows it to columns 1-6.'],
      [1, 'Level 1: marL one column module, from the master body style.'],
    ]),
    textBox(5, 'Caption', field(7, 12, 3, 3), [
      { style: 'caption', text: 'Inheritance: slide → layout → master. The slide placeholders have an empty p:spPr and a:bodyPr.' },
      { style: 'caption', text: 'Master: title row 2, body rows 3-4 across 12 columns; p:txStyles sets sizes, spacing and indents.' },
      { style: 'caption', text: 'Layout: overrides the body xfrm only, to columns 1-6.' },
    ]),
  ], {
    master: {
      shapes: ph(2, 'Title', '<p:ph type="title"/>', MASTER_TITLE) + ph(3, 'Body', '<p:ph type="body" idx="1"/>', MASTER_BODY),
      txStyles: TX_STYLES,
    },
    layout: {
      type: 'obj',
      name: 'Title and body',
      shapes: ph(2, 'Title', '<p:ph type="title"/>') + ph(3, 'Body', '<p:ph idx="1"/>', LAYOUT_BODY),
    },
  }),
};

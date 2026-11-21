import { COLOR, DOCX, FONT } from '../lib/grid.mjs';
import { docxPackage, header } from '../lib/docx.mjs';
import { esc } from '../lib/xml.mjs';

// Decoration drawn behind and under runs, where its extent is the feature:
//   A. Highlight (w:highlight yellow) across a manual line break and over
//      runs of two sizes: the band follows each line and each run's height.
//   B. Pattern shading on runs (w:shd pct15, pct50, solid accent) on an
//      exact 12 pt line and on an auto (single) line: the band fills the
//      line box of exact lines, and the font's line height otherwise.
//   C. A heading in white on a black paragraph fill (w:pPr/w:shd) at 24 pt,
//      exact 24 pt and auto line height: the fill height follows the line.
//   D. Underlined trailing spaces at a manual line break (one run, and the
//      spaces split into their own run): the underline ends with the text
//      and is not carried to the start of the next line.
//   E. Run effects: small caps, all caps, strike, double strike, outline,
//      shadow, emboss, imprint.
//   F. A paragraph with shading and a box border with w:space 0: the fill
//      and the border share their edges.
const { LINE } = DOCX;
const t = (text) => `<w:t xml:space="preserve">${esc(text)}</w:t>`;
const r = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}${t(text)}</w:r>`;
const BR = '<w:r><w:br/></w:r>';
const JA = `<w:rFonts w:ascii="${FONT.eastAsia}" w:eastAsia="${FONT.eastAsia}" w:hAnsi="${FONT.eastAsia}"/>`;
const HL = '<w:highlight w:val="yellow"/>';
const SZ12 = '<w:sz w:val="24"/>';
const p = (runs, pPr = '', height = 1) => ({ kind: 'raw', height, xml: `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>` });
const AUTO = `<w:spacing w:line="${LINE}" w:lineRule="auto"/>`;
const caption = (text, extra = {}) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1, ...extra });
const shd = (val, fill = 'auto', color = COLOR.ink) => `<w:shd w:val="${val}" w:color="${color}" w:fill="${fill}"/>`;
const BOX = '<w:pBdr>' + ['top', 'left', 'bottom', 'right']
  .map((n) => `<w:${n} w:val="single" w:sz="8" w:space="0" w:color="${COLOR.ink}"/>`).join('') + '</w:pBdr>';
const HALF = `<w:ind w:right="${DOCX.LIVE_W - DOCX.colEnd(2)}"/>`;

export default {
  id: 'docx-037',
  slug: 'run-decoration',
  format: 'docx',
  title: 'Run highlight, shading and effects',
  covers: [
    'word/document.xml|w:rPr>w:highlight',
    'word/document.xml|w:highlight@w:val=yellow',
    'word/document.xml|w:rPr>w:shd',
    'word/document.xml|w:shd@w:val',
    'word/document.xml|w:shd@w:val=solid',
    'word/document.xml|w:pPr>w:shd',
    'word/document.xml|w:rPr>w:u',
    'word/document.xml|w:rPr>w:smallCaps',
    'word/document.xml|w:rPr>w:caps',
    'word/document.xml|w:rPr>w:strike',
    'word/document.xml|w:rPr>w:dstrike',
    'word/document.xml|w:rPr>w:outline',
    'word/document.xml|w:rPr>w:shadow',
    'word/document.xml|w:rPr>w:emboss',
    'word/document.xml|w:rPr>w:imprint',
    'word/document.xml|w:rPr>w:color',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-037',
      title: 'Run highlight, shading and effects',
      verify: [
        'Verify: highlight and shading follow each line and run size;',
        'the black fill follows its line height; underlines stop at',
        'the break; F fill and border edges coincide.',
      ],
    }),
    caption('A. Highlight across a line break and two run sizes', { at: DOCX.rowLine(2), before: 0 }),
    p(r('Highlighted text runs to the end of this line', HL) + BR
      + r('and continues on the next line in ', HL) + r('12 pt', HL + SZ12) + r(' and 9 pt.', HL)
      + BR + r('日本語の文にも蛍光ペンを引く。', JA + HL), '', 3),
    caption('B. Run shading pct15, pct50, solid: exact line, then auto line'),
    p(r('pct15 pattern', shd('pct15')) + r('   ') + r('pct50 pattern', shd('pct50')) + r('   ')
      + r('solid accent', shd('solid', 'auto', COLOR.accent)) + r('   exact 12 pt line')),
    p(r('pct15 pattern', shd('pct15')) + r('   ') + r('pct50 pattern', shd('pct50')) + r('   ')
      + r('solid accent', shd('solid', 'auto', COLOR.accent)) + r('   auto line height'), AUTO, NaN),
    caption('C. White heading on a black paragraph fill, exact and auto'),
    p(r('Heading on exact 24 pt', '<w:b/><w:color w:val="FFFFFF"/><w:sz w:val="48"/>'),
      `${shd('clear', COLOR.ink)}<w:spacing w:line="${2 * LINE}" w:lineRule="exact"/>`, 2),
    { kind: 'p', lines: [''], height: 1 },
    p(r('Heading on auto line height', '<w:b/><w:color w:val="FFFFFF"/><w:sz w:val="48"/>'),
      `${shd('clear', COLOR.ink)}${AUTO}`, NaN),
    caption('D. Underlined trailing spaces before a line break'),
    p(r('One run ends with three spaces   ', '<w:u w:val="single"/>') + BR
      + r('Next line: no underline before this text.') + BR
      + r('Spaces in their own run', '<w:u w:val="single"/>') + r('   ', '<w:u w:val="single"/>') + BR
      + r('Next line: no underline before this text either.'), '', 4),
    caption('E. Run effects'),
    p(r('Small caps', '<w:smallCaps/>') + r(', ') + r('all caps', '<w:caps/>') + r(', ')
      + r('strike', '<w:strike/>') + r(', ') + r('double strike', '<w:dstrike/>') + r(', ')
      + r('outline', '<w:outline/>') + r(', ') + r('shadow', '<w:shadow/>') + r(', ')
      + r('emboss', '<w:emboss/>') + r(', ') + r('imprint', '<w:imprint/>') + r('.')),
    caption('F. Paragraph shading inside a box border, w:space 0'),
    p(r('The pct15 fill and the 1 pt border share their edges.'), `${BOX}${shd('pct15')}${HALF}`, NaN),
  ]),
};

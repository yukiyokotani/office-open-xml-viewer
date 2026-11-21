import { DOCX } from '../lib/grid.mjs';
import { WML, docxPackage, header } from '../lib/docx.mjs';
import { XML_DECL } from '../lib/xml.mjs';

const at = DOCX.rowLine;
const { colStart: start } = DOCX;

// Numbers sit on a column edge, text on the next: the hanging indent is one
// column module.
const lvl = (ilvl, numFmt, lvlText) => `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/>`
  + `<w:numFmt w:val="${numFmt}"/><w:lvlText w:val="${lvlText}"/><w:lvlJc w:val="left"/>`
  + `<w:pPr><w:ind w:left="${start(ilvl + 2)}" w:hanging="${start(ilvl + 2) - start(ilvl + 1)}"/></w:pPr></w:lvl>`;

const NUMBERING = XML_DECL
  + '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
  + `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>${lvl(0, 'decimal', '%1.')}${lvl(1, 'decimal', '%1.%2')}</w:abstractNum>`
  + `<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="singleLevel"/>${lvl(0, 'bullet', '–')}</w:abstractNum>`
  + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
  + '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
  + '<w:num w:numId="3"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>'
  + '</w:numbering>';

const item = (numId, ilvl, text) => ({
  kind: 'p', lines: [text], numPr: `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`,
});
const caption = (text, row) => ({ kind: 'p', at: row ? at(row) : undefined, before: row ? 0 : 1, style: 'caption', bold: true, lines: [text] });

export default {
  id: 'docx-008',
  slug: 'numbering',
  format: 'docx',
  title: 'Numbering',
  covers: [
    'word/numbering.xml|w:abstractNum>w:lvl',
    'word/numbering.xml|w:lvl>w:numFmt',
    'word/numbering.xml|w:numFmt@w:val=decimal',
    'word/numbering.xml|w:numFmt@w:val=bullet',
    'word/numbering.xml|w:lvl>w:lvlText',
    'word/numbering.xml|w:lvl>w:start',
    'word/numbering.xml|w:abstractNum>w:multiLevelType',
    'word/numbering.xml|w:num>w:abstractNumId',
    'word/numbering.xml|w:lvlOverride>w:startOverride',
    'word/document.xml|w:pPr>w:numPr',
    'word/document.xml|w:numPr>w:ilvl',
    'word/document.xml|w:numPr>w:numId',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-008',
      title: 'Numbering',
      verify: [
        'Verify: numbers and bullets on a column edge, text',
        'on the next; counts 1, 2, 2.1, 2.2, 3 and 5, 6.',
      ],
    }),
    caption('Decimal list in two levels: w:abstractNum 0, used by w:num 1', 2),
    item(1, 0, `Level 0: numFmt decimal, lvlText %1., number at 0, text at ${start(2)}.`),
    item(1, 0, 'The count continues from item to item within one w:num.'),
    item(1, 1, `Level 1: lvlText %1.%2, number at ${start(2)}, text at ${start(3)}.`),
    item(1, 1, 'A level-1 count restarts under each new level-0 item.'),
    item(1, 0, 'Back at level 0, the count resumes at 3.'),
    caption('Bullet list: w:abstractNum 1, numFmt bullet, lvlText en dash', 3),
    item(2, 0, `The bullet sits at 0 and the text at ${start(2)}, as in level 0 above.`),
    item(2, 0, 'A bullet is a numbering level whose text does not count.'),
    caption('Restart: w:num 3 reuses w:abstractNum 0 with w:startOverride 5', 4),
    item(3, 0, 'This item is numbered 5 by a level override.'),
    item(3, 0, 'The next item follows at 6.'),
  ], {
    parts: { 'word/numbering.xml': NUMBERING },
    rels: [['rId3', 'numbering', 'numbering.xml']],
    overrides: { 'word/numbering.xml': `${WML}.numbering+xml` },
  }),
};

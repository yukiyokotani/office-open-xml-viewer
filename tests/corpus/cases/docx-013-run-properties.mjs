import { COLOR, DOCX, FONT } from '../lib/grid.mjs';
import { docxPackage, header, run } from '../lib/docx.mjs';

const at = DOCX.rowLine;
// One property per line: the element name in column 1, the sample from column 2.
const TAB = [['left', DOCX.colStart(2)]];
const PROPS = [
  ['w:b', '<w:b/>', 'Bold: the second weight of the family.'],
  ['w:i', '<w:i/>', 'Italic: a slanted form, synthesised where the font has none.'],
  ['w:u single', '<w:u w:val="single"/>', 'Underline: a single rule below the baseline.'],
  ['w:strike', '<w:strike/>', 'Strikethrough: a single rule through the x-height.'],
  ['w:caps', '<w:caps/>', 'Capitals: lower-case letters shown as capitals.'],
  ['w:smallCaps', '<w:smallCaps/>', 'Small capitals: lower-case letters as reduced capitals.'],
  ['w:spacing 20', '<w:spacing w:val="20"/>', 'Letter spacing: 1 pt added after each character.'],
  [`w:color ${COLOR.accent}`, `<w:color w:val="${COLOR.accent}"/>`, 'Colour: the accent cyan of the corpus.'],
  ['w:vanish', '<w:vanish/>', 'Hidden text'],
  ['w:rFonts eastAsia', `<w:rFonts w:eastAsia="${FONT.eastAsia}"/>`, '和文はNoto Sans JP、欧文はRobotoで組む。'],
];

export default {
  id: 'docx-013',
  slug: 'run-properties',
  format: 'docx',
  title: 'Run properties',
  covers: [
    'word/document.xml|w:rPr>w:i',
    'word/document.xml|w:rPr>w:u',
    'word/document.xml|w:u@w:val=single',
    'word/document.xml|w:rPr>w:strike',
    'word/document.xml|w:rPr>w:caps',
    'word/document.xml|w:rPr>w:smallCaps',
    'word/document.xml|w:rPr>w:spacing',
    'word/document.xml|w:rPr>w:color',
    'word/document.xml|w:rPr>w:vanish',
    'word/document.xml|w:rPr>w:vertAlign',
    'word/document.xml|w:vertAlign@w:val=superscript',
    'word/document.xml|w:vertAlign@w:val=subscript',
    'word/document.xml|w:rPr>w:rFonts',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-013',
      title: 'Run properties',
      verify: [
        'Verify: each sample shows only the property named beside it;',
        'nothing appears between the brackets of the hidden line.',
      ],
    }),
    { kind: 'p', at: at(2), style: 'caption', bold: true, lines: ['Element\tSample'], tabs: TAB },
    ...PROPS.map(([name, rPr, sample]) => ({
      kind: 'p',
      tabs: TAB,
      height: 1,
      runs: name === 'w:vanish'
        ? [run(`${name}\t[`), run(sample, rPr), run(']')]
        : [run(`${name}\t`), run(sample, rPr)],
    })),
    {
      kind: 'p', tabs: TAB, height: 1,
      runs: [run('w:vertAlign\tSuperscript: 10'), run('3', '<w:vertAlign w:val="superscript"/>'),
        run(' mm is one metre. Subscript: log'), run('10', '<w:vertAlign w:val="subscript"/>'), run(' 1000 = 3.')],
    },
  ]),
};

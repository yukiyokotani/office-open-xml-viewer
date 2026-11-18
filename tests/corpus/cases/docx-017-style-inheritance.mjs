import { DOCX, FONT } from '../lib/grid.mjs';
import { docxPackage, header, run } from '../lib/docx.mjs';

const at = DOCX.rowLine;
const INDENT = DOCX.colStart(2);
const CAPTION_SZ = DOCX.TYPE.caption.sz;

// Normal → Body (adds the indent) → Body Note (adds the caption size); Label is
// a character style. Every paragraph below uses Body or Body Note.
const STYLES = '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
  + `<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/><w:basedOn w:val="Normal"/><w:next w:val="Body"/><w:uiPriority w:val="1"/><w:qFormat/><w:pPr><w:ind w:left="${INDENT}"/></w:pPr></w:style>`
  + `<w:style w:type="paragraph" w:styleId="BodyNote"><w:name w:val="Body Note"/><w:basedOn w:val="Body"/><w:next w:val="Body"/><w:uiPriority w:val="2"/><w:qFormat/><w:rPr><w:sz w:val="${CAPTION_SZ}"/></w:rPr></w:style>`
  + '<w:style w:type="character" w:styleId="Label"><w:name w:val="Label"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="3"/><w:qFormat/><w:rPr><w:b/><w:bCs/></w:rPr></w:style>';

const body = (runs, extra = {}) => ({ kind: 'p', pStyle: 'Body', height: 1, runs, ...extra });

export default {
  id: 'docx-017',
  slug: 'style-inheritance',
  format: 'docx',
  title: 'Style inheritance',
  covers: [
    'word/styles.xml|w:style>w:basedOn',
    'word/styles.xml|w:style>w:next',
    'word/styles.xml|w:style>w:qFormat',
    'word/styles.xml|w:style>w:semiHidden',
    'word/styles.xml|w:style>w:unhideWhenUsed',
    'word/styles.xml|w:style@w:type=character',
    'word/styles.xml|w:style@w:default',
    'word/document.xml|w:rPr>w:rStyle',
    'word/document.xml|w:rPr>w:sz',
    'word/document.xml|w:rPr>w:szCs',
    'word/document.xml|w:rFonts@w:ascii',
    'word/document.xml|w:rFonts@w:cs',
    'word/document.xml|w:rFonts@w:hint',
    'word/document.xml|w:rPr>w:lang',
    'word/document.xml|w:rPr>w:bCs',
    'word/document.xml|w:pPr>w:rPr',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-017',
      title: 'Style inheritance',
      verify: [
        'Verify: every line starts on column 2 (from Body); the Note',
        'line and the w:sz run are 8 pt; Label is bold; nothing else is.',
      ],
    }),
    { kind: 'p', at: at(2), style: 'caption', bold: true, lines: ['Normal → Body → Body Note; character style Label'] },
    body([run(`Body: basedOn Normal, adds w:ind left ${INDENT}, so its text starts on column 2.`)]),
    { kind: 'p', pStyle: 'BodyNote', lines: [`Body Note: basedOn Body, adds w:sz ${CAPTION_SZ}; it inherits the indent and is set at 8 pt.`] },
    body([run('Label', '<w:rStyle w:val="Label"/>'), run(': a character style, bold, applied to one run of a Body paragraph.')]),
    body([run('Direct formatting wins: '), run(`w:sz ${CAPTION_SZ} and w:szCs ${CAPTION_SZ} on this run`, `<w:sz w:val="${CAPTION_SZ}"/><w:szCs w:val="${CAPTION_SZ}"/>`), run(' override the style size.')]),
    body([run(`w:rFonts ascii, hAnsi and cs name ${FONT.latin} on the run, w:hint eastAsia: 和文`, `<w:rFonts w:ascii="${FONT.latin}" w:hAnsi="${FONT.latin}" w:cs="${FONT.latin}" w:hint="eastAsia"/>`)]),
    body([run('w:lang de-DE on this run: Maßstab und Satzspiegel.', '<w:lang w:val="de-DE"/>')]),
    body([run('w:bCs marks complex-script bold; it does not change this Latin text.', '<w:bCs/>')]),
    body([run('The paragraph mark of this line is bold (w:pPr/w:rPr); its text is not.')], { markRPr: '<w:b/>' }),
  ], { styles: STYLES }),
};

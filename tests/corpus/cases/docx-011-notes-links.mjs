import { DOCX } from '../lib/grid.mjs';
import { R_NS, W, WML, docxPackage, header, run } from '../lib/docx.mjs';
import { XML_DECL } from '../lib/xml.mjs';

const at = DOCX.rowLine;
const ECMA_376 = 'https://ecma-international.org/publications-and-standards/standards/ecma-376/';

// Link: underlined only. Note reference: superscript. Note text: caption size.
const STYLES = '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:u w:val="single"/></w:rPr></w:style>'
  + '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>'
  + `<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="${DOCX.TYPE.caption.sz}"/></w:rPr></w:style>`;
const NOTE_REF = '<w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>';

const FOOTNOTES = XML_DECL + `<w:footnotes ${W}>`
  + '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
  + '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
  + `<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r>${NOTE_REF}<w:footnoteRef/></w:r>`
  + '<w:r><w:t xml:space="preserve"> ISO/IEC 29500, the international standard for the same formats, was first published in 2008.</w:t></w:r></w:p></w:footnote>'
  + '</w:footnotes>';

const BR = '<w:r><w:br/></w:r>';
const link = (attrsXml, text) => `<w:hyperlink ${attrsXml}><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>${text}</w:t></w:r></w:hyperlink>`;

export default {
  id: 'docx-011',
  slug: 'notes-links',
  format: 'docx',
  title: 'Footnotes and links',
  covers: [
    'word/document.xml|w:r>w:footnoteReference',
    'word/footnotes.xml|w:footnotes>w:footnote',
    'word/footnotes.xml|w:footnote@w:type=separator',
    'word/footnotes.xml|w:r>w:footnoteRef',
    'word/footnotes.xml|w:r>w:separator',
    'word/settings.xml|w:footnotePr>w:footnote',
    'word/document.xml|w:p>w:hyperlink',
    'word/document.xml|w:hyperlink@r:id',
    'word/document.xml|w:hyperlink@w:anchor',
    'word/document.xml|w:p>w:bookmarkStart',
    'word/document.xml|w:p>w:bookmarkEnd',
    'word/styles.xml|w:rPr>w:vertAlign',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-011',
      title: 'Footnotes and links',
      verify: [
        'Verify: note 1 at the foot of the page; the external and',
        'internal links underlined; the bookmark target on line 21.',
      ],
    }),
    {
      kind: 'p',
      at: at(2),
      cols: [1, 2],
      height: 4,
      runs: [
        run('The file formats in this corpus are defined by'), BR,
        link('r:id="rId4"', 'ECMA-376'),
        `<w:r>${NOTE_REF}<w:footnoteReference w:id="1"/></w:r>`,
        run(', first adopted in December 2006.'), BR,
        run('The link is a w:hyperlink with an external'), BR,
        run('relationship.'),
      ],
    },
    {
      kind: 'p',
      before: 1,
      cols: [1, 2],
      height: 2,
      runs: [run('An internal link points to a bookmark instead:'), BR, link('w:anchor="grid"', 'go to the grid note'), run('.')],
    },
    {
      kind: 'p',
      at: at(3),
      cols: [1, 2],
      height: 2,
      runs: [
        '<w:bookmarkStart w:id="0" w:name="grid"/>',
        run('Grid note: this paragraph is the bookmark named grid.'), BR,
        run('It starts on line 21 of 60.'),
        '<w:bookmarkEnd w:id="0"/>',
      ],
    },
  ], {
    ns: [R_NS],
    styles: STYLES,
    settings: '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>',
    parts: { 'word/footnotes.xml': FOOTNOTES },
    rels: [['rId3', 'footnotes', 'footnotes.xml']],
    overrides: { 'word/footnotes.xml': `${WML}.footnotes+xml` },
    externalRels: [['rId4', 'hyperlink', ECMA_376]],
  }),
};

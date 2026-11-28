import { DOCX } from '../lib/grid.mjs';
import { R_NS, WML, W, docxPackage, header, run } from '../lib/docx.mjs';
import { XML_DECL } from '../lib/xml.mjs';

const at = DOCX.rowLine;
const RIGHT_TAB = `<w:tabs><w:tab w:val="right" w:pos="${DOCX.LIVE_W}"/></w:tabs>`;
const captionPara = (content) => `<w:p><w:pPr><w:pStyle w:val="Caption"/>${RIGHT_TAB}</w:pPr>${content}</w:p>`;

// PAGE as a simple field, NUMPAGES as a complex field: both forms of w:fld.
const FIELD_PAGE = '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>';
const FIELD_NUMPAGES = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>'
  + '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';

const HEADER = XML_DECL + `<w:hdr ${W}>${captionPara(
  `${run('Running head: w:hdr, two baselines above the live area')}${run('\tdocx-009')}`,
)}</w:hdr>`;
const FOOTER = XML_DECL + `<w:ftr ${W}>${captionPara(
  `${run('Running foot: w:ftr, two baselines below the live area')}${run('\tPage ')}${FIELD_PAGE}${run(' of ')}${FIELD_NUMPAGES}`,
)}</w:ftr>`;

export default {
  id: 'docx-009',
  slug: 'headers-footers',
  format: 'docx',
  title: 'Headers and footers',
  covers: [
    'word/document.xml|w:sectPr>w:headerReference',
    'word/document.xml|w:sectPr>w:footerReference',
    'word/document.xml|w:headerReference@w:type=default',
    'word/document.xml|w:pgMar@w:header',
    'word/document.xml|w:br@w:type=page',
    'word/header1.xml|w:hdr>w:p',
    'word/footer1.xml|w:ftr>w:p',
    'word/footer1.xml|w:p>w:fldSimple',
    'word/footer1.xml|w:fldSimple@w:instr',
    'word/footer1.xml|w:r>w:fldChar',
    'word/footer1.xml|w:fldChar@w:fldCharType=begin',
    'word/footer1.xml|w:r>w:instrText',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-009',
      title: 'Headers and footers',
      verify: [
        'Verify: the head and foot repeat on both pages, one',
        'empty line outside the live area; the foot counts 1 of 2, 2 of 2.',
      ],
    }),
    {
      kind: 'p',
      at: at(2),
      cols: [1, 2],
      lines: [
        'Header and footer distances come from the grid:',
        `w:header ${DOCX.HEADER_DIST} and w:footer ${DOCX.FOOTER_DIST} twips put`,
        'each line two baselines outside the live area.',
        'The page number is a PAGE field, the page count',
        'a NUMPAGES field. A page break follows.',
      ],
    },
    { kind: 'p', pageBreakBefore: true, cols: [1, 2], lines: ['Page 2 begins on grid line 1 after w:br w:type page.'] },
  ], {
    ns: [R_NS],
    firstSection: {
      refs: '<w:headerReference w:type="default" r:id="rId3"/><w:footerReference w:type="default" r:id="rId4"/>',
      headerDist: DOCX.HEADER_DIST,
      footerDist: DOCX.FOOTER_DIST,
    },
    parts: { 'word/header1.xml': HEADER, 'word/footer1.xml': FOOTER },
    rels: [['rId3', 'header', 'header1.xml'], ['rId4', 'footer', 'footer1.xml']],
    overrides: { 'word/header1.xml': `${WML}.header+xml`, 'word/footer1.xml': `${WML}.footer+xml` },
  }),
};

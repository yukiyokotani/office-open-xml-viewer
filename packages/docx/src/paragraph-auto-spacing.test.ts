import { readFile } from 'node:fs/promises';
import { beforeAll, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { normalizeDocxDocumentModel, bodyLayoutAcquisitionInput } from './parser-model.js';
import type { BodyElement, DocxDocumentModel, DocParagraph } from './types.js';

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const O = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const AUTO_STYLE = '<w:style w:type="paragraph" w:styleId="Auto"><w:pPr><w:spacing w:before="100" w:after="100" w:beforeAutospacing="1" w:afterAutospacing="1"/></w:pPr></w:style>';
const p = (props = '', text = 'text') => `<w:p><w:pPr><w:pStyle w:val="Auto"/>${props}</w:pPr><w:r><w:rPr><w:sz w:val="48"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
const cell = (content: string) => `<w:tc><w:tcPr><w:vAlign w:val="center"/></w:tcPr>${content}</w:tc>`;
const table = (...cells: string[]) => `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid>${cells.map(() => '<w:gridCol w:w="2500"/>').join('')}</w:tblGrid><w:tr>${cells.map(cell).join('')}</w:tr></w:tbl>`;
const numPr = (numId: number, level = 0) => `<w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${numId}"/></w:numPr>`;
const item = (numId: number, level: number, text: string) => p(numPr(numId, level), text);

function parse(
  body: string,
  options: Readonly<{ fixed?: boolean; parts?: ReadonlyMap<string, readonly [type: string, xml: string]>; sectPr?: string }> = {},
): DocxDocumentModel {
  const files = new Map<string, string>([
    ['word/document.xml', `<w:document xmlns:w="${W}" xmlns:r="${O}" xmlns:v="urn:schemas-microsoft-com:vml"><w:body>${body}<w:sectPr>${options.sectPr ?? ''}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`],
    ['word/styles.xml', `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>${AUTO_STYLE}</w:styles>`],
    ['word/settings.xml', `<w:settings xmlns:w="${W}"><w:compat>${options.fixed ? '<w:doNotUseHTMLParagraphAutoSpacing/>' : ''}</w:compat></w:settings>`],
  ]);
  const relationships: [id: string, type: string, target: string][] = [['styles', 'styles', 'styles.xml'], ['settings', 'settings', 'settings.xml']];
  for (const [id, [type, xml]] of options.parts ?? []) {
    files.set(`word/${id}.xml`, xml);
    relationships.push([id, type, `${id}.xml`]);
  }
  files.set('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.map(([id, type, target]) => `<Relationship Id="${id}" Type="${O}/${type}" Target="${target}"/>`).join('')}</Relationships>`);
  files.set('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${O}/officeDocument" Target="word/document.xml"/></Relationships>`);
  files.set('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  const archive = new DocxArchive(storeZip(new Map([...files].map(([name, xml]) => [name, new TextEncoder().encode(xml)]))));
  try { return JSON.parse(new TextDecoder().decode(archive.parse())) as DocxDocumentModel; }
  finally { archive.free(); }
}

/** Cell, text-box, header, footer and footnote stories that each hold one `p()`. */
function containersDocument(fixed: boolean): DocxDocumentModel {
  const box = `<w:p><w:r><w:pict><v:shape id="box" type="#_x0000_t202" style="width:100pt;height:100pt"><v:textbox inset="0,0,0,0"><w:txbxContent>${p()}</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>`;
  return parse(table(p()) + box + p(), {
    fixed,
    sectPr: '<w:headerReference w:type="default" r:id="header"/><w:footerReference w:type="default" r:id="footer"/>',
    parts: new Map([
      ['header', ['header', `<w:hdr xmlns:w="${W}">${p()}</w:hdr>`]],
      ['footer', ['footer', `<w:ftr xmlns:w="${W}">${p()}</w:ftr>`]],
      ['footnotes', ['footnotes', `<w:footnotes xmlns:w="${W}"><w:footnote w:id="1">${p()}</w:footnote></w:footnotes>`]],
    ]),
  });
}

const margins = (content: readonly BodyElement[] | undefined) => (content ?? [])
  .filter((e): e is DocParagraph & { type: 'paragraph' } => e.type === 'paragraph')
  .map(paragraph => [paragraph.spaceBefore, paragraph.spaceAfter]);
const cellMargins = (content: readonly BodyElement[]) => {
  const tableElement = content.find(e => e.type === 'table');
  if (tableElement?.type !== 'table') throw new Error('expected a table');
  return tableElement.rows[0]!.cells.map(c => margins(c.content));
};

it('uses a 14pt automatic margin at every font size while explicit false retains authored spacing', () => {
  // A leading table keeps the document-start edge away from the probes.
  const raw = parse(table(p()) + p() + p('<w:spacing w:beforeAutospacing="0" w:afterAutospacing="0"/>')
    + p('<w:spacing w:before="800" w:beforeLines="500"/>'));
  const normalized = normalizeDocxDocumentModel(raw);
  expect(margins(normalized.body)).toEqual([[14, 14], [5, 5], [14, 14]]);
  // Projection is immutable and the pagination sequence uses the same resolved pair.
  expect((raw.body[1] as DocParagraph).spaceBefore).toBe(5);
  expect(normalizeDocxDocumentModel(normalized).body).toEqual(normalized.body);
  const sequence = bodyLayoutAcquisitionInput(normalized).sequence;
  expect(sequence.flatMap(e => e.kind === 'body-block' && e.block.kind === 'paragraph'
    ? [[e.block.spaceBeforePt, e.block.spaceAfterPt]] : [])).toEqual([[14, 14], [5, 5], [14, 14]]);
});

it('retains stored margins and drops no automatic edge under Word fixed automatic-spacing compatibility', () => {
  const normalized = normalizeDocxDocumentModel(parse(
    p() + table(p() + p()) + item(1, 0, 'a') + item(1, 0, 'b') + p('<w:spacing w:before="800" w:beforeLines="500"/>'),
    { fixed: true },
  ));
  expect(margins(normalized.body)).toEqual([[5, 5], [5, 5], [5, 5], [40, 5]]);
  expect(cellMargins(normalized.body)).toEqual([[[5, 5], [5, 5]]]);
});

it('resolves Word HTML automatic spacing at cell, document-start and same-list boundaries', () => {
  const lvl = (ilvl: number, text: string) => `<w:lvl w:ilvl="${ilvl}"><w:numFmt w:val="bullet"/><w:lvlText w:val="${text}"/></w:lvl>`;
  const numbering = `<w:numbering xmlns:w="${W}">${[1, 2].map(id => `<w:abstractNum w:abstractNumId="${id}">${lvl(0, '•')}${lvl(1, 'o')}</w:abstractNum>`).join('')}<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num></w:numbering>`;
  const stored = p(`${numPr(1)}<w:spacing w:before="40" w:after="60" w:beforeAutospacing="0" w:afterAutospacing="0"/>`, 'stored');
  const normalized = normalizeDocxDocumentModel(parse(
    p('', 'first')
    + table(p('', 'only'), p('', 'top') + p('', 'middle') + p('', 'bottom'), item(1, 0, 'cell 1') + item(1, 0, 'cell 2'))
    + p('', 'lead') + item(1, 0, 'a') + item(1, 0, 'b') + item(1, 1, 'b.i') + item(1, 0, 'c') + stored
    + p('', 'between') + item(2, 0, 'other list') + item(1, 0, 'back'),
    { parts: new Map([['numbering', ['numbering', numbering]]]) },
  ));
  expect(margins(normalized.body)).toEqual([
    [0, 14], // document body start: no automatic before
    [14, 14], // lead
    [14, 0], [0, 0], [0, 0], [0, 0], // a, b, b.i (level 1), c: same numId
    [2, 3], // stored values of a same-list item stay
    [14, 14], // between: unnumbered
    [14, 14], [14, 14], // another numId separates the lists
  ]);
  // Cell edges drop the first automatic before and the last automatic after.
  expect(cellMargins(normalized.body)).toEqual([
    [[0, 0]],
    [[0, 14], [14, 14], [14, 0]],
    [[0, 0], [0, 0]],
  ]);
});

it.each([false, true])('propagates document fixed=%s into inherited cell and story margins', (fixed) => {
  const raw = containersDocument(fixed);
  const normalized = normalizeDocxDocumentModel(raw);
  // A cell's single paragraph sits on both cell edges; header, footer and
  // note story edges are not observed box edges.
  expect(cellMargins(normalized.body)).toEqual([[fixed ? [5, 5] : [0, 0]]]);
  for (const content of [normalized.headers.default?.body, normalized.footers.default?.body, normalized.footnotes?.[0]?.content]) {
    expect(margins(content)).toEqual([fixed ? [5, 5] : [14, 14]]);
  }
  expect((raw.headers.default?.body[0] as DocParagraph).spaceBefore).toBe(5);
});

it('keeps parsed compatibility text-box margins independent of the first inline font', () => {
  const raw = containersDocument(false);
  const normalized = normalizeDocxDocumentModel(raw);
  const shapes = normalized.body.flatMap(e => e.type === 'paragraph' ? e.runs.filter(run => run.type === 'shape') : []);
  expect(shapes).toHaveLength(1);
  const block = shapes[0]?.type === 'shape' ? shapes[0].textBlocks?.[0] : undefined;
  expect(block?.fontSizePt).toBe(24);
  expect([block?.spaceBefore, block?.spaceAfter]).toEqual([14, 14]);
  expect(normalizeDocxDocumentModel(normalized).body).toEqual(normalized.body);
});

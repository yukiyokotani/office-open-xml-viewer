import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { normalizeDocxDocumentModel, bodyLayoutAcquisitionInput } from './parser-model.js';
import { layoutDocument } from './document-layout.js';
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
  const numbering = `<w:numbering xmlns:w="${W}">${[1, 2].map(id => `<w:abstractNum w:abstractNumId="${id}">${lvl(0, '•')}${lvl(1, 'o')}</w:abstractNum>`).join('')}<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num><w:num w:numId="3"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
  const stored = p(`${numPr(1)}<w:spacing w:before="40" w:after="60" w:beforeAutospacing="0" w:afterAutospacing="0"/>`, 'stored');
  const normalized = normalizeDocxDocumentModel(parse(
    p('', 'first')
    + table(p('', 'only'), p('', 'top') + p('', 'middle') + p('', 'bottom'), item(1, 0, 'cell 1') + item(1, 0, 'cell 2'))
    + p('', 'lead') + item(1, 0, 'a') + item(1, 0, 'b') + item(1, 1, 'b.i') + item(1, 0, 'c') + stored
    + p('', 'between') + item(2, 0, 'other list') + item(1, 0, 'back') + item(3, 0, 'shared abstract'),
    { parts: new Map([['numbering', ['numbering', numbering]]]) },
  ));
  expect(margins(normalized.body)).toEqual([
    [0, 14], // document body start: no automatic before
    [14, 14], // lead
    [14, 0], [0, 0], [0, 0], [0, 0], // a, b, b.i (level 1), c: same numId
    [2, 3], // stored values of a same-list item stay
    [14, 14], // between: unnumbered
    [14, 14], [14, 0], // another abstract definition separates the lists
    [0, 14], // a numId sharing the abstract definition continues the list
  ]);
  // Cell edges drop the first automatic before and the last automatic after.
  expect(cellMargins(normalized.body)).toEqual([
    [[0, 0]],
    [[0, 14], [14, 14], [14, 0]],
    [[0, 0], [0, 0]],
  ]);
});

it('takes body neighbours from flow participants, skipping vanished paragraphs layout swallows', () => {
  const lvl = '<w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl>';
  const numbering = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="1">${lvl}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
  const hidden = (props = '') => `<w:p><w:pPr>${props}<w:rPr><w:vanish/></w:rPr></w:pPr></w:p>`;
  const frame = '<w:framePr w:w="2000" w:hAnchor="page" w:vAnchor="page" w:x="1440" w:y="1440"/>';
  const byText = (raw: DocxDocumentModel) => Object.fromEntries(normalizeDocxDocumentModel(raw).body
    .flatMap(e => e.type === 'paragraph' && e.runs.length > 0
      ? [[e.runs.map(run => run.type === 'text' ? run.text : '').join(''), [e.spaceBefore, e.spaceAfter]]] : []));
  const parts = new Map([['numbering', ['numbering', numbering] as const]]);
  expect(byText(parse(
    // A swallowed paragraph before the first block does not move the body start.
    hidden() + p('', 'first')
    // An unnumbered swallowed paragraph does not separate same-list items.
    + item(1, 0, 'a') + hidden() + item(1, 0, 'b') + p('', 'gap')
    // A numbered swallowed paragraph does not join its visible neighbours.
    + item(1, 0, 'c') + hidden(`<w:pStyle w:val="Auto"/>${numPr(1)}`) + p('', 'plain')
    // A swallowed paragraph whose page break applies still separates the items.
    + item(1, 0, 'd') + hidden('<w:pageBreakBefore/>') + p(`${numPr(1)}${frame}`, 'framed'),
    { parts },
  ))).toEqual({
    first: [0, 14],
    a: [14, 0], b: [0, 14], gap: [14, 14],
    c: [14, 14], plain: [14, 14],
    d: [14, 14], framed: [14, 14],
  });
  // Without the applying break the framed item joins the list.
  expect(byText(parse(p('', 'lead') + item(1, 0, 'd') + hidden() + p(`${numPr(1)}${frame}`, 'framed'), { parts })))
    .toEqual({ lead: [0, 14], d: [14, 0], framed: [0, 14] });
});

it.each([false, true])('propagates document fixed=%s into inherited cell and story margins', (fixed) => {
  const raw = containersDocument(fixed);
  const normalized = normalizeDocxDocumentModel(raw);
  // A cell's single paragraph sits on both cell edges and a header's first
  // paragraph on the header start; the footer end and note edges keep theirs.
  expect(cellMargins(normalized.body)).toEqual([[fixed ? [5, 5] : [0, 0]]]);
  expect(margins(normalized.headers.default?.body)).toEqual([fixed ? [5, 5] : [0, 14]]);
  for (const content of [normalized.footers.default?.body, normalized.footnotes?.[0]?.content]) {
    expect(margins(content)).toEqual([fixed ? [5, 5] : [14, 14]]);
  }
  expect((raw.headers.default?.body[0] as DocParagraph).spaceBefore).toBe(5);
});

it('drops the automatic before of the first text-box paragraph independent of the first inline font', () => {
  const raw = containersDocument(false);
  const normalized = normalizeDocxDocumentModel(raw);
  const shapes = normalized.body.flatMap(e => e.type === 'paragraph' ? e.runs.filter(run => run.type === 'shape') : []);
  expect(shapes).toHaveLength(1);
  const block = shapes[0]?.type === 'shape' ? shapes[0].textBlocks?.[0] : undefined;
  expect(block?.fontSizePt).toBe(24);
  expect([block?.spaceBefore, block?.spaceAfter]).toEqual([0, 14]);
  expect(normalizeDocxDocumentModel(normalized).body).toEqual(normalized.body);
});

(globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas ??= class {
  getContext() {
    let font = '10px serif';
    return {
      get font() { return font; }, set font(v: string) { font = v; },
      letterSpacing: '0px',
      measureText: (s: string) => {
        const size = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
        return { width: [...s].length * size * 0.5, fontBoundingBoxAscent: size * 0.8, fontBoundingBoxDescent: size * 0.2,
          actualBoundingBoxAscent: size * 0.8, actualBoundingBoxDescent: size * 0.2 } as TextMetrics;
      },
    };
  }
};

/** First baseline of each probe text on the page (a table cell places each
 * child at its block offset inside the cell). */
function pageBaselines(raw: DocxDocumentModel): Readonly<Record<string, number>> {
  const found: Record<string, number> = {};
  const visit = (node: unknown, deltaPt = 0): void => {
    if (node === null || typeof node !== 'object') return;
    if (Array.isArray(node)) { node.forEach(child => visit(child, deltaPt)); return; }
    const record = node as Record<string, unknown>;
    if (record.kind === 'table-cell' && Array.isArray(record.blocks)) {
      const cellTopPt = (record.flowBounds as { yPt: number }).yPt;
      for (const block of record.blocks as Array<{ offsetPt: number; layout: { flowBounds: { yPt: number } } }>) {
        visit(block.layout, deltaPt + cellTopPt + block.offsetPt - block.layout.flowBounds.yPt);
      }
      return;
    }
    if (record.kind === 'paragraph' && Array.isArray(record.lines)) {
      for (const line of record.lines as Array<{ baselinePt: number; placements?: Array<{ kind: string; text?: string }> }>) {
        for (const placement of line.placements ?? []) {
          if (placement.kind === 'text' && placement.text && !(placement.text in found)) found[placement.text] = line.baselinePt + deltaPt;
        }
      }
    }
    for (const value of Object.values(record)) visit(value, deltaPt);
  };
  const layout = layoutDocument(raw);
  expect(layout.pages).toHaveLength(1);
  visit(layout.pages[0]);
  return found;
}

describe('Word fixed automatic-spacing compatibility: additive paragraph gaps', () => {
  // Mirrors the Windows Word controls of WORD_FIXED_PARAGRAPH_SPACING_ADDITIVE_GAP:
  // exact 12pt lines and stored before/after 3/8pt, so each baseline step is
  // 12pt plus the gap.
  const lines = 'w:line="240" w:lineRule="exact"';
  const probe = (text: string, spacing: string, props = '') =>
    `<w:p><w:pPr>${props}<w:spacing ${spacing} ${lines}/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const plain = (text: string, props = '') => probe(text, 'w:before="0" w:after="0"', props);
  const auto = (text: string, props = '') => probe(text, 'w:before="60" w:after="160" w:beforeAutospacing="1" w:afterAutospacing="1"', props);
  const stored = (text: string, props = '') => probe(text, 'w:before="60" w:after="160"', props);
  const frame = '<w:framePr w:w="2000" w:hAnchor="page" w:vAnchor="page" w:x="5000" w:y="9000"/>';
  const box = (content: string) => `<w:p><w:r><w:pict><v:shape id="box" type="#_x0000_t202" style="position:absolute;margin-left:300pt;margin-top:500pt;width:100pt;height:100pt"><v:textbox inset="0,0,0,0"><w:txbxContent>${content}</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>`;

  const noFlags = 'w:beforeAutospacing="0" w:afterAutospacing="0"';
  const contextualPair = probe('KA', `w:before="0" w:after="60" ${noFlags}`, '<w:pStyle w:val="Auto"/><w:contextualSpacing/>')
    + probe('KB', `w:before="160" w:after="0" ${noFlags}`, '<w:pStyle w:val="Auto"/>');

  function baselines(fixed: boolean): Readonly<Record<string, number>> {
    return pageBaselines(parse(
      plain('R0') + auto('A1') + auto('A2') + stored('S3') + stored('S4') + plain('R5')
      + contextualPair + plain('R6')
      + table(auto('CA') + auto('CB') + stored('CC'))
      + box(auto('TA') + auto('TB'))
      + auto('FA', frame) + auto('FB', frame) + plain('END'),
      {
        fixed,
        sectPr: '<w:headerReference w:type="default" r:id="header"/>',
        parts: new Map([['header', ['header', `<w:hdr xmlns:w="${W}">${auto('HA')}${auto('HB')}</w:hdr>`]]]),
      },
    ));
  }

  const steps = (found: Readonly<Record<string, number>>, ...texts: string[]) =>
    texts.slice(1).map((text, index) => Math.round((found[text]! - found[texts[index]!]!) * 100) / 100);

  it('adds the space after and the next space before in every flow', () => {
    const fixed = baselines(true);
    // Body: automatic and stored neighbours both add 8 + 3; a zero side adds nothing.
    expect(steps(fixed, 'R0', 'A1', 'A2', 'S3', 'S4', 'R5')).toEqual([15, 23, 23, 23, 20]);
    // An earlier-only contextualSpacing toggle drops the 3pt after and keeps the whole 8pt before.
    expect(steps(fixed, 'KA', 'KB')).toEqual([20]);
    for (const container of [['CA', 'CB', 'CC'], ['TA', 'TB'], ['HA', 'HB'], ['FA', 'FB']]) {
      expect(steps(fixed, ...container), container.join()).toEqual(container.slice(1).map(() => 23));
    }
  });

  it('keeps the §17.3.1.33 maximum without the setting', () => {
    const html = baselines(false);
    expect(steps(html, 'R0', 'A1', 'A2', 'S3', 'S4', 'R5')).toEqual([26, 26, 26, 20, 20]);
    // The maximum fold leaves only what the 8pt before exceeds the dropped 3pt after.
    expect(steps(html, 'KA', 'KB')).toEqual([17]);
    for (const container of [['CA', 'CB', 'CC'], ['TA', 'TB'], ['HA', 'HB'], ['FA', 'FB']]) {
      expect(steps(html, ...container), container.join()).toEqual(container.slice(1).map(() => 26));
    }
  });
});

describe('Word hidden cell paragraphs (WORD_HIDDEN_CELL_PARAGRAPH_JOINS_NEXT)', () => {
  const lines = 'w:line="240" w:lineRule="exact"';
  const probe = (text: string, spacing: string, mark = '') =>
    `<w:p><w:pPr><w:spacing ${spacing} ${lines}/>${mark}</w:pPr>${text ? `<w:r><w:t>${text}</w:t></w:r>` : ''}</w:p>`;
  const plain = (text: string) => probe(text, 'w:before="0" w:after="0"');
  const auto = (text: string) => probe(text, 'w:before="60" w:after="160" w:beforeAutospacing="1" w:afterAutospacing="1"');
  const vanished = (spacing = 'w:before="100" w:after="220"') => probe('', spacing, '<w:rPr><w:vanish/></w:rPr>');
  const tc = (content: string) => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
  const grid = '<w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>';
  const nested = (text: string) => `<w:tbl>${grid}<w:tr>${tc(plain(text))}</w:tr></w:tbl>`;
  // Each variant row sits between plain one-line reference rows.
  const rows = [
    vanished() + vanished() + auto('X1'), // leading: dropped with its own spacing
    auto('X3a') + vanished() + auto('X3b'), // between: dropped, the neighbours meet
    vanished() + nested('X5n') + plain('X5z'), // before a nested table: dropped
    auto('X7a') + vanished('w:before="0" w:after="0"'), // the cell's last paragraph: kept
  ];
  const body = `<w:tbl>${grid}<w:tr>${tc(plain('XR0'))}</w:tr>${rows.map((content, index) =>
    `<w:tr>${tc(content)}</w:tr><w:tr>${tc(plain(`XR${index + 1}`))}</w:tr>`).join('')}</w:tbl>${plain('END')}`;
  const order = ['XR0', 'X1', 'XR1', 'X3a', 'X3b', 'XR2', 'X5n', 'X5z', 'XR3', 'X7a', 'XR4'];

  it.each([
    [false, [12, 12, 12, 26, 12, 12, 12, 12, 12, 38]],
    [true, [15, 20, 15, 23, 20, 12, 12, 12, 15, 32]],
  ])('lays out no block for a non-final vanished inkless cell paragraph (fixed=%s)', (fixed, expected) => {
    const found = pageBaselines(parse(body, { fixed }));
    expect(order.slice(1).map((text, index) => Math.round((found[text]! - found[order[index]!]!) * 100) / 100))
      .toEqual(expected);
  });
});

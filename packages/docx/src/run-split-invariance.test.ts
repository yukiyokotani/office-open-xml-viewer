import { compatProfile } from './test-support/compat-environment.test-support.js';
import { createFontResolver } from './layout/font-service.js';
import { createTextLayoutService } from './layout/text.js';
import { readFile } from 'node:fs/promises';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { measureParagraphIntrinsicWidths } from './layout/intrinsic-width.js';
import { acquireShapeTextBoxLayout } from './layout/paragraph.js';
import type { ParagraphLayoutContext } from './layout-context.js';
import type { ShapeRun } from './types.js';
import { beforeAll, describe, expect, it } from 'vitest';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import { textRunsForPage } from './text-run-projection.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, DocxTextRun, SectionProps } from './types.js';
import type { DocumentLayout, ParagraphLayout } from './layout/types.js';

const borders = { top: null, right: null, bottom: null, left: null, insideH: null, insideV: null };
function run(text: string, extra: Partial<DocxTextRun> = {}) {
  return { type: 'text' as const, text, fontSize: 18, fontFamily: 'Arial', kerning: 8,
    bold: false, italic: false, underline: false, strikethrough: false, color: null,
    isLink: false, background: null, vertAlign: null, hyperlink: null, ...extra };
}
function paragraph(parts: string[], extra: Partial<DocxTextRun>, alignment: DocParagraph['alignment']): DocParagraph {
  return { runs: parts.map(text => run(text, extra)), alignment, indentLeft: 0, indentRight: 0,
    indentFirst: 0, spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null,
    tabStops: [], defaultFontSize: extra.fontSize ?? 18, defaultFontFamily: extra.fontFamily ?? 'Arial',
    widowControl: false };
}
function model(p: DocParagraph, width: number, container: 'paragraph' | 'fixed' | 'autofit'): DocxDocumentModel {
  const body: BodyElement[] = [{ type: 'paragraph', ...p }];
  if (container !== 'paragraph') body.splice(0, 1, { type: 'table', layout: container,
    colWidths: [width], widthPt: width, borders, cellMarginTop: 0, cellMarginRight: 0,
    cellMarginBottom: 0, cellMarginLeft: 0, jc: 'left', rows: [{ cells: [{
      content: [{ type: 'paragraph', ...p }], colSpan: 1, vMerge: null, borders,
      background: null, vAlign: 'top', widthPt: width, widthPct: undefined,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    }], rowHeight: null, rowHeightRule: 'auto', isHeader: false }] });
  return { body, settings: { compatibilityMode: 15 }, section: { pageWidth: width,
    pageHeight: 1000, marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false } as SectionProps,
    headers: { default: null, first: null, even: null }, footers: { default: null, first: null, even: null },
    fontFamilyClasses: {}, footnotes: [] } as unknown as DocxDocumentModel;
}
function context(): CanvasRenderingContext2D {
  // Independent pair metrics, including a pair spanning an ordinary space.
  // The Arial T/i scalars and T-space pair reproduce the reviewer's 10.8pt cell.
  const ctx = { font: '18px Arial', fontKerning: 'auto', letterSpacing: '0px',
    measureText(text: string) {
      const size = Number(/([\d.]+)px/u.exec(ctx.font)?.[1] ?? 18);
      const scalar = (c: string) => c === 'T' ? 10.9951171875 : c === ' ' ? 5.0009765625
        : c === 'i' ? 3.9990234375 : 10;
      const pair = ctx.fontKerning === 'normal'
        ? (text.match(/T /gu)?.length ?? 0) * .3251953125 + (text.match(/AV/gu)?.length ?? 0) * 2 : 0;
      const width = ([...text].reduce((sum, c) => sum + scalar(c), 0) - pair) * size / 18;
      return { width, fontBoundingBoxAscent: size * .8, fontBoundingBoxDescent: size * .2,
        actualBoundingBoxLeft: 0, actualBoundingBoxRight: width,
        actualBoundingBoxAscent: size * .8, actualBoundingBoxDescent: size * .2 } as TextMetrics;
    } };
  return ctx as unknown as CanvasRenderingContext2D;
}
function acquire(parts: string[], width: number, container: 'paragraph' | 'fixed' | 'autofit',
  extra: Partial<DocxTextRun> = {}, alignment: DocParagraph['alignment'] = 'right') {
  const doc = model(paragraph(parts, extra, alignment), width, container);
  return layoutDocument(doc, createLayoutServices(doc, { measureContext: context() }));
}
function intrinsic(runs: DocParagraph['runs']) {
  const p = { ...paragraph([], {}, 'right'), runs };
  const doc = model(p, 100, 'paragraph');
  const measure = context();
  return measureParagraphIntrinsicWidths(p, contextPt, 100,
    { context: measure, fontFamilyClasses: {} }, {
      pageIndex: 0, totalPages: 1, pageWritingMode: 'horizontal-tb',
      documentHasEastAsianText: false, compat: compatProfile({ compatibilityMode: 15 }),
      layoutServices: createLayoutServices(doc, { measureContext: measure }),
    });
}
function paragraphs(layout: DocumentLayout): ParagraphLayout[] {
  const result: ParagraphLayout[] = [];
  const seen = new WeakSet<object>();
  const visit = (value: unknown): void => {
    if (!value || typeof value !== 'object' || seen.has(value)) return;
    seen.add(value);
    if ('kind' in value && value.kind === 'paragraph' && 'lines' in value && 'textBoxes' in value) {
      const p = value as ParagraphLayout;
      result.push(p);
      p.textBoxes.forEach(visit);
      return;
    }
    for (const [key, child] of Object.entries(value)) if (key !== 'source') {
      if (Array.isArray(child)) child.forEach(visit); else visit(child);
    }
  };
  layout.pages.forEach(visit);
  return result;
}
function geometry(layout: DocumentLayout | readonly ParagraphLayout[]) {
  return (Array.isArray(layout) ? layout as ParagraphLayout[] : paragraphs(layout as DocumentLayout)).map(p => ({ bounds: p.flowBounds, lines: p.lines.map(l => ({
    range: l.range, bounds: l.bounds, advance: l.advancePt, baseline: l.baselinePt,
    text: l.placements.flatMap(s => s.kind === 'text' ? [s.text] : []).join(''),
    glyphs: l.placements.flatMap(s => s.kind === 'text' ? s.clusters.map(c => ({
      range: c.range, x: s.origin.xPt + c.offset.xPt, advance: c.advancePt,
    })) : []),
    paint: l.placements.flatMap(s => s.kind === 'text' ? s.paintOps.map(op => ({
      text: op.text, x: s.origin.xPt + op.offset.xPt, y: s.origin.yPt + op.offset.yPt,
      range: op.range, kerning: op.kerning,
    })) : []),
  })) }));
}
function partitions(text: string): string[][] {
  const result = [[text], [...text]];
  for (let i = 1; i < text.length; i++) result.push([text.slice(0, i), text.slice(i)]);
  // Stable pseudo-random partitions: reproducible failures and no fixture matrix.
  let state = 1703;
  for (let n = 0; n < 8; n++) {
    const parts: string[] = []; let start = 0;
    for (let i = 1; i < text.length; i++) {
      state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
      if (state % 3 === 0) { parts.push(text.slice(start, i)); start = i; }
    }
    parts.push(text.slice(start)); result.push(parts);
  }
  return [...new Map(result.map(parts => [JSON.stringify(parts), parts])).values()];
}

describe('formatting-only run boundaries are transparent to text acquisition', () => {
  it.each(['paragraph', 'fixed'] as const)('rejects the narrow T-space prefix consistently in a %s', container => {
    const expected = geometry(acquire(['T i'], 10.8, container));
    expect(expected[0]?.lines.map(l => l.text)).toEqual(['T', ' i']);
    expect(expected[0]?.lines[0]?.glyphs[0]?.x).toBeCloseTo(-.1951171875, 8);
    for (const parts of partitions('T i')) expect(geometry(acquire(parts, 10.8, container)), parts.join('|')).toEqual(expected);
    const accepted = geometry(acquire(['T i'], 16, container));
    expect(accepted[0]?.lines.map(l => l.text)).toEqual(['T ', 'i']);
    for (const parts of partitions('T i')) expect(geometry(acquire(parts, 16, container))).toEqual(accepted);
  });

  // The parser-backed matrix below owns font/threshold variation. This case
  // separately protects intrinsic sizing and each alignment's placement path.
  it('preserves intrinsic widths, container alignment and retained paint across source partitions', () => {
    const text = 'AV T i\tAV-T i';
    const format = { fontFamily: 'Arial', fontSize: 18, kerning: 8 };
    const widths = intrinsic([run(text, format)]);
    for (const parts of partitions(text)) expect(intrinsic(parts.map(text => run(text, format)))).toEqual(widths);
    for (const container of ['paragraph', 'fixed', 'autofit'] as const) {
      for (const alignment of ['left', 'right', 'center', 'both'] as const) {
        const expected = geometry(acquire([text], 35, container, format, alignment));
        for (const parts of partitions(text)) {
          expect(geometry(acquire(parts, 35, container, format, alignment)), `${container}/${alignment}/${parts.join('|')}`).toEqual(expected);
        }
      }
    }
    // ~2.4 s on an idle machine; the default 5 s budget timed out when
    // the suite ran on a loaded machine.
  }, 30_000);

  it('keeps original source ownership after shaping a sequence', () => {
    const layout = acquire(['A', 'V T', ' i'], 100, 'paragraph');
    const runs = textRunsForPage(layout, 0, { scale: 1 });
    expect(runs.map(r => r.text).join('')).toBe('AV T i');
    for (const [index, text] of ['A', 'V T', ' i'].entries()) {
      expect(runs.filter(r => r.sourceRunIndex === index).map(r => r.text).join('')).toBe(text);
    }
  });
});


beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

// Feed the complete WASM wire graph to the production source-store adapter.
// Extracting public runs here would discard the retained typography inputs and
// would miss a content-dependent acquisition key (the r4 regression).
function parsedDocument(parts: string[], wrapper: string, format: Partial<DocxTextRun> = {}, runProperties: string[] = [],
  options: { alignment?: DocParagraph['alignment']; compatibilityMode?: number; widthPt?: number } = {}): DocxDocumentModel {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const O = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const escape = (text: string) => text.replace(/&/gu, '&amp;').replace(/</gu, '&lt;');
  const runs = parts.map((text, index) => {
    // U+2011 is this fixture's marker for the authored OOXML element, which
    // the parser displays as U+002D while retaining its no-break ownership.
    const content = text.split('\t').map(piece => piece.split('\u2011')
      .map(part => `<w:t xml:space="preserve">${escape(part)}</w:t>`)
      .join('<w:noBreakHyphen/>')).join('<w:tab/>');
    const run = `<w:r w:rsidR="0000000${index % 8}">${runProperties[index] ? `<w:rPr>${runProperties[index]}</w:rPr>` : ''}${content}</w:r>`;
    if (['deletion-seam', 'moveFrom-seam', 'deleted-break-seam', 'deleted-tab-seam', 'deleted-math-seam'].includes(wrapper)) {
      const tag = wrapper === 'moveFrom-seam' ? 'moveFrom' : 'del';
      const math = wrapper === 'deleted-math-seam' ? '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>Q</m:t></m:r></m:oMath>' : '';
      const boundary = wrapper === 'deleted-break-seam' ? '<w:br/>' : wrapper === 'deleted-tab-seam' ? '<w:tab/>' : '';
      return `${index ? `<w:${tag} w:id="${index}" w:author="Reviewer"><w:r><w:delText>X</w:delText>${boundary}</w:r>${math}</w:${tag}>` : ''}${run}`;
    }
    if (wrapper === 'ruby-tail' && index === parts.length - 1) return `<w:r><w:ruby><w:rubyPr><w:hps w:val="12"/><w:hpsRaise w:val="6"/><w:hpsBaseText w:val="36"/></w:rubyPr><w:rt><w:r><w:t>hint</w:t></w:r></w:rt><w:rubyBase>${run}</w:rubyBase></w:ruby></w:r>`;
    if (wrapper === 'smart-tags') return `<w:smartTag w:uri="urn:test" w:element="word">${run}</w:smartTag>`;
    if (wrapper === 'revisions') return `<w:ins w:id="${index}" w:author="Reviewer">${run}</w:ins>`;
    if (wrapper === 'hyperlinks') return `<w:hyperlink w:anchor="Destination">${run}</w:hyperlink>`;
    if (wrapper === 'simple-fields') return `<w:fldSimple w:instr="AUTHOR">${run}</w:fldSimple>`;
    if (wrapper === 'bookmarks') return `<w:bookmarkStart w:id="${index}" w:name="B${index}"/>${run}<w:bookmarkEnd w:id="${index}"/>`;
    if (wrapper === 'comments') return `<w:commentRangeStart w:id="${index}"/>${run}<w:commentRangeEnd w:id="${index}"/>`;
    if (wrapper === 'proofing') return `<w:proofErr w:type="spellStart"/>${run}<w:proofErr w:type="spellEnd"/>`;
    if (wrapper === 'complex-fields') return `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>AUTHOR</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${run}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    return run;
  }).join('');
  const p = (indented = false) => `<w:p><w:pPr><w:jc w:val="${options.alignment ?? 'right'}"/><w:spacing w:before="0" w:after="0"/>${indented ? `<w:ind w:right="${Math.round((200 - (options.widthPt ?? 19.8)) * 20)}"/>` : ''}</w:pPr>${runs}</w:p>`;
  const table = (kind: 'fixed' | 'autofit') => `<w:tbl><w:tblPr><w:tblLayout w:type="${kind}"/><w:tblW w:w="396" w:type="dxa"/><w:tblCellMar><w:top w:w="0"/><w:left w:w="0"/><w:bottom w:w="0"/><w:right w:w="0"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="396"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="396" w:type="dxa"/></w:tcPr>${p()}</w:tc></w:tr></w:tbl>`;
  const box = `<w:p><w:r><w:pict><v:shape id="box" type="#_x0000_t202" style="width:19.8pt;height:100pt"><v:textbox inset="0,0,0,0"><w:txbxContent>${p()}</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>`;
  const body = p(true) + table('fixed') + table('autofit') + box + '<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>';
  const files = new Map<string, string>([
    ['word/document.xml', `<w:document xmlns:w="${W}" xmlns:r="${O}" xmlns:v="urn:schemas-microsoft-com:vml"><w:body>${body}<w:sectPr><w:headerReference w:type="default" r:id="hdr"/><w:footerReference w:type="default" r:id="ftr"/><w:pgSz w:w="4000" w:h="20000"/><w:pgMar w:top="2400" w:right="0" w:bottom="2400" w:left="0" w:header="300" w:footer="300"/></w:sectPr></w:body></w:document>`],
    ['word/styles.xml', `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${format.fontFamily ?? 'Arial'}" w:hAnsi="${format.fontFamily ?? 'Arial'}"/><w:sz w:val="${2 * (format.fontSize ?? 18)}"/><w:kern w:val="${2 * (format.kerning ?? 8)}"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`],
    ['word/settings.xml', `<w:settings xmlns:w="${W}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${options.compatibilityMode ?? 15}"/></w:compat></w:settings>`],
    ['word/header1.xml', `<w:hdr xmlns:w="${W}">${table('fixed')}</w:hdr>`],
    ['word/footer1.xml', `<w:ftr xmlns:w="${W}">${table('fixed')}</w:ftr>`],
    ['word/footnotes.xml', `<w:footnotes xmlns:w="${W}"><w:footnote w:id="1">${table('fixed')}</w:footnote></w:footnotes>`],
  ]);
  const relationships = [['hdr', 'header', 'header1.xml'], ['ftr', 'footer', 'footer1.xml'],
    ['styles', 'styles', 'styles.xml'], ['settings', 'settings', 'settings.xml'], ['notes', 'footnotes', 'footnotes.xml']];
  files.set('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.map(([id, type, file]) => `<Relationship Id="${id}" Type="${O}/${type}" Target="${file}"/>`).join('')}</Relationships>`);
  files.set('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${O}/officeDocument" Target="word/document.xml"/></Relationships>`);
  const partTypes = [['document.xml', 'document.main'], ['header1.xml', 'header'], ['footer1.xml', 'footer'],
    ['styles.xml', 'styles'], ['settings.xml', 'settings'], ['footnotes.xml', 'footnotes']];
  files.set('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/>${partTypes.map(([file, type]) => `<Override PartName="/word/${file}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/>`).join('')}</Types>`);
  const archive = new DocxArchive(storeZip(new Map([...files].map(([name, xml]) => [name, new TextEncoder().encode(xml)]))));
  try { return JSON.parse(new TextDecoder().decode(archive.parse())) as DocxDocumentModel; }
  finally { archive.free(); }
}
function acquireParsed(parts: string[], wrapper: string, format?: Partial<DocxTextRun>) {
  const doc = parsedDocument(parts, wrapper, format);
  return layoutDocument(doc, createLayoutServices(doc, { measureContext: context() }));
}

function acquireRuns(runs: DocParagraph['runs'], container: 'paragraph' | 'fixed' | 'autofit') {
  const doc = model({ ...paragraph([], {}, 'right'), runs }, 10.8, container);
  return layoutDocument(doc, createLayoutServices(doc, { measureContext: context() }));
}

describe('complete DOCX parser inputs preserve formatting-only split invariance', () => {
  it('wraps at an ordinary numeric hyphen without depending on source splits', () => {
    const acquire = (parts: string[], properties: string[] = []) => {
      const doc = parsedDocument(parts, 'rsid', {}, properties,
        { alignment: 'left', compatibilityMode: 14, widthPt: 86 });
      return geometry(paragraphs(layoutDocument(doc,
        createLayoutServices(doc, { measureContext: context() })))
        .filter(p => p.source.story === 'body' && p.source.path.length === 1))[0];
    };
    // §17.3.3.18 explicitly contrasts a numeric U+002D break with
    // noBreakHyphen. The visible first prefix advances 85pt in these metrics.
    const whole = acquire(['AAAA 166-67']);
    expect(whole?.lines.map(l => l.text.trim())).toEqual(['AAAA 166-', '67']);
    expect(acquire(['AAAA 166-', '67'])).toEqual(whole);
    expect(acquire(['AAAA 166-', '67'], ['', '<w:b/>'])?.lines.map(l => l.text.trim()))
      .toEqual(['AAAA 166-', '67']);
    expect(acquire(['AAAA 166\u201167'])?.lines.map(l => l.text.trim()))
      .toEqual(['AAAA', '166-67']);
    const ordinary = intrinsic([run('166-67')]);
    const protectedRun = { ...run('166-67'), noBreakRanges: [{ start: 3, end: 4 }] };
    expect(ordinary.minWidthPt).toBeCloseTo(intrinsic([run('166-')]).maxWidthPt, 8);
    expect(ordinary.maxWidthPt).toBeCloseTo(intrinsic([protectedRun]).maxWidthPt, 8);
    expect(intrinsic([protectedRun]).minWidthPt)
      .toBeCloseTo(ordinary.maxWidthPt, 8);
  });
  it('uses the combining base before an ordinary hyphen and keeps a hyphen extension attached', () => {
    const lines = (parts: string[], properties: string[] = []) => {
      const doc = parsedDocument(parts, 'rsid', {}, properties,
        { alignment: 'left', widthPt: 76 });
      return geometry(paragraphs(layoutDocument(doc, createLayoutServices(doc, { measureContext: context() })))
        .filter(p => p.source.story === 'body' && p.source.path.length === 1))[0]?.lines.map(l => l.text.trim());
    };
    // The independent metrics give the complete first prefix 75pt. A mark on
    // the Latin base must not remove its hyphen opportunity; a mark attached
    // to the hyphen stays on its complete extended grapheme.
    expect(lines(['AAAA a\u0301-b'])).toEqual(['AAAA a\u0301-', 'b']);
    expect(lines(['AAAA a', '\u0301-', 'b'], ['', '<w:b/>', '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>']))
      .toEqual(['AAAA a\u0301-', 'b']);
    expect(lines(['AAAA a-\u0301b'])).toEqual(['AAAA a-\u0301', 'b']);
    expect(lines(['AAAA a-', '\u0301b'], ['', '<w:b/>'])).toEqual(['AAAA a-\u0301', 'b']);
    expect(lines(['AAAA a\u2011\u0301b'])).toEqual(['AAAA', 'a-\u0301b']);
    expect(lines(['AAAA a\u2011', '\u0301b'], ['', '<w:b/>'])).toEqual(['AAAA', 'a-\u0301b']);
  });
  it('keeps a parenthesized signed number intact across a genuine formatting seam', () => {
    const doc = parsedDocument(['AA (-', '123)'], 'rsid', {}, ['', '<w:b/>'],
      { alignment: 'left', widthPt: 66 });
    const body = geometry(paragraphs(layoutDocument(doc, createLayoutServices(doc, { measureContext: context() })))
      .filter(p => p.source.story === 'body' && p.source.path.length === 1))[0];
    expect(body?.lines.map(l => l.text.trim())).toEqual(['AA', '(-123)']);
    const hebrew = parsedDocument(['AA א-', 'b'], 'rsid', {}, ['', '<w:b/>'],
      { alignment: 'left', widthPt: 46 });
    const hebrewBody = geometry(paragraphs(layoutDocument(hebrew,
      createLayoutServices(hebrew, { measureContext: context() })))
      .filter(p => p.source.story === 'body' && p.source.path.length === 1))[0];
    expect(hebrewBody?.lines.map(l => l.text.trim())).toEqual(['AA', 'א-b']);
  });
  it('keeps ruby base ownership when script or small-caps splitting emits an unannotated tail', () => {
    const lines = (text: string, properties: string = '') => {
      const doc = parsedDocument(['AAAA ', text], 'ruby-tail', {}, ['', properties], { alignment: 'left', widthPt: 76 });
      return geometry(paragraphs(layoutDocument(doc, createLayoutServices(doc, { measureContext: context() })))
        .filter(p => p.source.story === 'body' && p.source.path.length === 1))[0]?.lines.map(l => l.text.trim());
    };
    // Ruby paints its annotation on the owning piece; existing script/case
    // breaks remain unchanged. The ordinary policy must not add an internal
    // hyphen break to an unannotated tail of that same authored base.
    expect(lines('a\u0301-b')).toEqual(['AAAA a\u0301', '-b']);
    expect(lines('aB-c', '<w:smallCaps/>')).toEqual(['AAAA', 'AB-C']);
  });
  it('admits the same legal hyphen prefix when its tail changes font', () => {
    const lines = (parts: string[], properties: string[] = []) => {
      const doc = parsedDocument(parts, 'rsid', {}, properties,
        { alignment: 'left', widthPt: 86 });
      return geometry(paragraphs(layoutDocument(doc, createLayoutServices(doc, { measureContext: context() })))
        .filter(p => p.source.story === 'body' && p.source.path.length === 1))[0]?.lines.map(l => l.text.trim());
    };
    expect(lines(['AAAA ab-cd'])).toEqual(['AAAA ab-', 'cd']);
    expect(lines(['AAAA ab-c', 'd'], ['', '<w:b/>'])).toEqual(['AAAA ab-', 'cd']);
  });
  it('prefers the CJK to whole-word boundary across actual Latin formatting changes', () => {
    const lines = (parts: string[], properties: string[]) => {
      const doc = parsedDocument(parts, 'rsid', {}, properties,
        { alignment: 'left', widthPt: 65 });
      return geometry(paragraphs(layoutDocument(doc, createLayoutServices(doc, { measureContext: context() })))
        .filter(p => p.source.story === 'body' && p.source.path.length === 1))[0]?.lines.map(l => l.text.trim());
    };
    expect(lines(['日日E-mail'], [])).toEqual(['日日', 'E-mail']);
    expect(lines(['日日', 'E-', 'mail'], ['', '<w:b/>', '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>']))
      .toEqual(['日日', 'E-mail']);
  });
  it('fits the visible justified prefix across a real formatting boundary in mode 14', () => {
    const acquire = (parts: string[], properties: string[]) => {
      const doc = parsedDocument(parts, 'rsid', {}, properties,
        { alignment: 'both', compatibilityMode: 14, widthPt: 14 });
      const layout = layoutDocument(doc, createLayoutServices(doc, { measureContext: context() }));
      return geometry(paragraphs(layout).filter(p => p.source.story === 'body' && p.source.path.length === 1))[0];
    };
    // The independent metrics above give "i i" a 13pt natural advance.
    // Its following 5pt separator is a line edge, so the visible prefix fits
    // the 14pt band. The last authored word remains genuinely bold.
    const whole = acquire(['i i i', ' i'], ['', '<w:b/>']);
    expect(whole?.lines.map(l => l.text.trim())).toEqual(['i i', 'i i']);
    expect(acquire(['i i', ' i', ' i'], ['', '', '<w:b/>'])).toEqual(whole);
  });
  it.each(['deletion-seam', 'moveFrom-seam', 'deleted-break-seam', 'deleted-tab-seam', 'deleted-math-seam'])('keeps omitted %s content out of final-view shaping boundaries', wrapper => {
    const whole = acquireParsed(['T i'], wrapper);
    const doc = parsedDocument(['T', ' i'], wrapper);
    const services = createLayoutServices(doc, { measureContext: context() });
    const final = layoutDocument(doc, services);
    expect(geometry(final)).toEqual(geometry(whole));
    const marked = layoutDocument(doc, services, { showTrackedChanges: true, currentDateMs: 0 });
    expect(geometry(marked).flatMap(p => p.lines.map(l => l.text)).some(text => text.includes('X'))).toBe(true);
    const body = textRunsForPage(final, 0, { scale: 1 }).filter(r => r.source?.story === 'body');
    expect(body.some(r => r.text.includes('X'))).toBe(false);
  });
  it.each(['deletion-seam', 'moveFrom-seam'])('retains one displayed coordinate domain after omitted %s and a real format change', wrapper => {
    const acquire = (parts: string[], properties: string[]) => {
      const doc = parsedDocument(parts, wrapper, {}, properties, { alignment: 'left', widthPt: 100 });
      return geometry(paragraphs(layoutDocument(doc, createLayoutServices(doc, { measureContext: context() })))
        .filter(p => p.source.story === 'body' && p.source.path.length === 1))[0];
    };
    const whole = acquire(['AV', 'B'], ['', '<w:b/>']);
    expect(whole?.lines[0]?.text).toBe('AVB');
    expect(whole?.lines[0]?.range).toEqual({ start: 0, end: 3 });
    expect(acquire(['A', 'V', 'B'], ['', '', '<w:b/>'])).toEqual(whole);
    expect(acquire(['A', 'V', 'B'], ['', '', ''])?.lines[0]?.range).toEqual({ start: 0, end: 3 });
    const markedDoc = parsedDocument(['A', 'V', 'B'], wrapper, {}, ['', '', '<w:b/>'],
      { alignment: 'left', widthPt: 100 });
    const marked = geometry(paragraphs(layoutDocument(markedDoc,
      createLayoutServices(markedDoc, { measureContext: context() }),
      { showTrackedChanges: true, currentDateMs: 0 }))
      .filter(p => p.source.story === 'body' && p.source.path.length === 1))[0];
    expect(marked?.lines[0]?.text).toBe('AXVXB');
    expect(marked?.lines[0]?.range).toEqual({ start: 0, end: 5 });
  });
  it('keeps Latin source seams transparent beside ideographic spaces in every story', () => {
    const whole = acquireParsed(['T i\u3000'], 'rsid');
    const split = acquireParsed(['T', ' i\u3000'], 'rsid');
    expect(geometry(split)).toEqual(geometry(whole));
  });
  it.each([
    ['shadow', '0', '1'], ['frame', '0', '1'], ['themeColor', 'accent1', 'accent2'],
    ['themeTint', '33', '66'], ['themeShade', '33', '66'],
  ])('preserves distinct retained border %s when adjacent public borders match', (attribute, first, second) => {
    const border = (value: string) => `<w:bdr w:val="single" w:sz="8" w:color="000000" w:${attribute}="${value}"/>`;
    const doc = parsedDocument(['A', 'V'], 'rsid', {}, [border(first), border(second)]);
    const layout = layoutDocument(doc, createLayoutServices(doc, { measureContext: context() }));
    const line = paragraphs(layout).find(p => p.source.story === 'body')?.lines[0];
    const facts = line?.placements.flatMap(p => p.kind === 'text' && p.runBorder ? [p.runBorder] : []);
    expect(facts?.map(f => f[attribute as keyof typeof f])).toEqual(
      attribute === 'shadow' || attribute === 'frame' ? [false, true] : [first, second]);
    // ECMA-376 §17.3.2.4: unlike attributes form separate border groups.
    expect(line?.placements.flatMap(p => p.kind === 'text' ? p.runBorderFragments ?? [] : [])).toHaveLength(8);
  });
  it.each(['rsid', 'smart-tags', 'revisions', 'hyperlinks', 'simple-fields', 'complex-fields', 'bookmarks', 'comments', 'proofing'])(
    '%s preserves all story/container geometry with inherited formatting', wrapper => {
      const layout = acquireParsed(['T i'], wrapper);
      expect(new Set(paragraphs(layout).map(p => p.source.story))).toEqual(new Set(['body', 'header', 'footer', 'footnote', 'textbox']));
      const expected = JSON.stringify(geometry(layout));
      // An independent accepted boundary for the native Arial T-space pair.
      expect(geometry(layout).some(p => p.lines.length === 1 && p.lines[0]?.text === 'T i')).toBe(true);
      for (const parts of partitions('T i')) {
        expect(JSON.stringify(geometry(acquireParsed(parts, wrapper))), parts.join('|')).toBe(expected);
      }
    },
  );
  it.each([
    { text: 'AVAT i-AV T i', fontFamily: 'Arial', fontSize: 18, kerning: 8 },
    { text: ' T i AV-T  i ', fontFamily: 'Times New Roman', fontSize: 10, kerning: 10 },
    { text: 'AV T i\tAV-T i', fontFamily: 'Georgia', fontSize: 24, kerning: 24.5 },
    { text: 'AV-T i T  i', fontFamily: 'Arial', fontSize: 18, kerning: 0 },
  ])('keeps parser-backed threshold and separator boundaries for $fontFamily / $kerning', ({ text, ...format }) => {
    const expected = JSON.stringify(geometry(acquireParsed([text], 'proofing', format)));
    for (const parts of partitions(text)) {
      expect(JSON.stringify(geometry(acquireParsed(parts, 'proofing', format))), parts.join('|')).toBe(expected);
    }
  }, 15000);
});

const contextPt: ParagraphLayoutContext = {
    lineGrid: { active: false, pitchPt: null },
    characterGrid: { active: false, kind: null, pitchPt: null, deltaPt: 0 },
    rightIndentGrid: { pitchPt: null, paragraphAllowsAdjustment: true },
    physicalIndentLeftPt: 0, physicalIndentRightPt: 0, firstIndentPt: 0,
    lineSpacing: null, spaceBeforePt: 0, spaceAfterPt: 0, baseRtl: false,
    isJustified: false, stretchLastLine: false, tabStops: [], hasRuby: false, hasEastAsianText: false,
    kinsoku: { enabled: true, lineStartForbidden: new Set(), lineEndForbidden: new Set() }, defaultTabPt: 36,
  };

it('uses the same sequence in the shape text-box adapter', () => {

  const box = (parts: string[]) => {
    const doc = model(paragraph([], {}, 'right'), 100, 'paragraph');
    const measure = context();
    const layout = acquireShapeTextBoxLayout({
      textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
      textBlocks: [{ text: parts.join(''), fontSizePt: 18, alignment: 'right',
        runs: parts.map(text => ({ text, fontSizePt: 18, fontFamily: 'Arial' })) }],
    } as ShapeRun, { xPt: 0, yPt: 0, widthPt: 10.8, heightPt: 100 }, {
      id: 'box', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
      flowDomainId: 'body', context: contextPt, measurer: { context: measure, fontFamilyClasses: {} },
      environment: { pageIndex: 0, totalPages: 1, documentHasEastAsianText: false,
        pageWritingMode: 'horizontal-tb', compat: compatProfile({ compatibilityMode: 15 }),
        layoutServices: createLayoutServices(doc, { measureContext: measure }) },
    });
    if (!layout) throw new Error('Expected text box');
    return geometry(layout.story.blocks.filter((b): b is ParagraphLayout => b.kind === 'paragraph'));
  };
  const expected = box(['T i']);
  for (const parts of partitions('T i')) expect(box(parts)).toEqual(expected);
});


it.each(['font', 'weight', 'threshold'] as const)('retains a real %s boundary in both narrow-cell and intrinsic paths', boundary => {
  const runs = [run('T'), run(' i', boundary === 'font' ? { fontFamily: 'Georgia' }
    : boundary === 'weight' ? { bold: true } : { kerning: 18.5 })];
  const fixed = acquireRuns(runs, 'fixed');
  expect(geometry(fixed)[0]?.lines[0]?.text).toBe('T');
  expect(geometry(fixed)[0]?.lines[0]?.glyphs[0]?.advance).toBeCloseTo(10.9951171875, 8);
  expect(intrinsic(runs)).toEqual({ minWidthPt: 10.9951171875, maxWidthPt: 19.9951171875 });
});

it('keeps source-addressed RTL geometry and later run ranges after sequence shaping', () => {
  const p = paragraph(['A', 'V T', ' i'], { rtl: true }, 'right');
  p.runs.push(run(' final', { bold: true, rtl: true }));
  const doc = model(p, 200, 'paragraph');
  const layout = layoutDocument(doc, createLayoutServices(doc, { measureContext: context() }));
  const runs = textRunsForPage(layout, 0, { scale: 1 });
  for (const text of ['A', 'V T', ' i', ' final']) {
    const index = ['A', 'V T', ' i', ' final'].indexOf(text);
    const owned = runs.filter(r => r.sourceRunIndex === index);
    expect(owned.map(r => r.text).join('')).toBe(text);
    expect(owned.every(r => r.w > 0)).toBe(true);
  }
  expect(paragraphs(layout)[0]?.lines[0]?.range.end).toBe('AV T i final'.length);
});


it('preserves independently resolved substitute faces at an otherwise identical source seam', () => {
  const p = paragraph(['مرحبا', '12'], { fontFamily: 'Scoped Face' }, 'left');
  const doc = model(p, 200, 'paragraph');
  const fonts = createFontResolver([
    { requestedFamily: 'Scoped Face', resolvedFamily: 'Arabic Substitute', source: 'substitute', script: 'arabic' },
  ], { scriptScopedFamilies: { 'scoped face': { script: 'arabic', substituteFamilies: ['Arabic Substitute'] } } });
  const services = createLayoutServices(doc, { measureContext: context() });
  const text = createTextLayoutService({ fonts, measurer: { fingerprint: 'split-script-proof',
    measure: request => ({ advancePt: [...request.text].length * 10, ascentPt: 8, descentPt: 2 }),
  } });
  const layout = layoutDocument(doc, { ...services, text });
  const runs = paragraphs(layout)[0]?.lines[0]?.placements.filter(s => s.kind === 'text') ?? [];
  const digits = runs.find(s => s.kind === 'text' && s.text === '12');
  const arabic = runs.find(s => s.kind === 'text' && s.text === 'مرحبا');
  expect(arabic?.fontRoute.familyList).toContain('Arabic Substitute');
  expect(digits?.fontRoute.familyList).not.toContain('Arabic Substitute');
  expect(digits?.sourceRunIndex).toBe(1);
});

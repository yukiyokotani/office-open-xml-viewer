import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it, vi } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import * as paginationWorkBudget from './layout/pagination-work-budget.js';
import { textRunsForPage } from './text-run-projection.js';
import type { DocxTextRunInfo } from './renderer.js';
import type { DocxDocumentModel } from './types.js';

// A text box nested in a page story (here a header) is acquired with ITS
// owner's section and destination page, not with the body's current location.
// The PAGE field inside the header text box is the independent observable:
// every page's header shows its own page number. The native variant feeds the
// same parsed model with a canonical native BtoT section (the producer's
// private raw flow 2 on its nominal btLr token), so the body runs in its
// counter-clockwise frame while the header story, and the text box inside it,
// stay horizontal and upright; a body text box exercises the upright-physical
// nested story on that native page.

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const V = 'urn:schemas-microsoft-com:vml';
const PAGE = { width: 612, height: 792 };

function textBox(content: string, blocks = false): string {
  return `<w:r><w:pict><v:shape id="box" type="#_x0000_t202" style="width:72pt;height:36pt">`
    + `<v:textbox inset="0,0,0,0"><w:txbxContent>${blocks ? content : `<w:p>${content}</w:p>`}</w:txbxContent></v:textbox>`
    + '</v:shape></w:pict></w:r>';
}

function parsedDocument(fields?: { kind: 'header' | 'footer'; pages: number; format: string }): DocxDocumentModel {
  let body = `<w:p><w:r><w:t>One</w:t></w:r>${textBox('<w:r><w:t>Box</w:t></w:r>')}</w:p>`
    + '<w:p><w:r><w:br w:type="page"/><w:t>Two</w:t></w:r></w:p>';
  let header = `<w:p>${textBox('<w:fldSimple w:instr=" PAGE "><w:r><w:t>9</w:t></w:r></w:fldSimple>')}</w:p>`;
  const kind = fields?.kind ?? 'header';
  if (fields) {
    const page = '<w:fldSimple w:instr=" PAGE "><w:r><w:t>?</w:t></w:r></w:fldSimple>';
    const table = (content: string) => '<w:tbl><w:tblPr><w:tblW w:w="1440" w:type="dxa"/></w:tblPr>'
      + '<w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid><w:tr><w:tc><w:tcPr/>'
      + content + '</w:tc></w:tr></w:tbl>';
    body = Array.from({ length: fields.pages }, (_, index) => '<w:p>'
      + (index ? '<w:pPr><w:pageBreakBefore/></w:pPr>' : '') + page + '</w:p>').join('');
    // Direct fields, table descendants and text boxes inside table cells all
    // read the same destination, including a second detached text-box root.
    header = `<w:p>${page}</w:p>` + table(`<w:p>${page}</w:p>`)
      + `<w:p>${textBox(table(table(`<w:p>${page}${textBox(page)}</w:p>`) + '<w:p/>'), true)}</w:p>`
      + table(`<w:p>${textBox(page)}</w:p>`);
  }
  const files = new Map<string, string>([
    ['word/document.xml', `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:v="${V}"><w:body>${body}`
      + `<w:sectPr><w:${kind}Reference w:type="default" r:id="hdr"/><w:pgSz w:w="12240" w:h="15840"/>`
      + (fields ? `<w:pgNumType w:start="12" w:fmt="${fields.format}"/>` : '')
      + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>'
      + '</w:sectPr></w:body></w:document>'],
    [`word/${kind}1.xml`, `<w:${kind === 'header' ? 'hdr' : 'ftr'} xmlns:w="${W}" xmlns:v="${V}">${header}</w:${kind === 'header' ? 'hdr' : 'ftr'}>`],
    ['word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="hdr" Type="${R}/${kind}" Target="${kind}1.xml"/></Relationships>`],
    ['_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + `<Override PartName="/word/${kind}1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml"/>`
      + '</Types>'],
  ]);
  return parseFiles(files);
}

function parseFiles(files: Map<string, string>): DocxDocumentModel {
  const archive = new DocxArchive(storeZip(new Map([...files].map(([name, xml]) => [name, new TextEncoder().encode(xml)]))));
  try { return JSON.parse(new TextDecoder().decode(archive.parse())) as DocxDocumentModel; }
  finally { archive.free(); }
}

/** The producer facts of a bottomToTop section: nominal btLr plus the
 * private section flow, as a producer emits them. */
function asBottomToTop(doc: DocxDocumentModel): DocxDocumentModel {
  const section = doc.section as DocxDocumentModel['section'] & { __sectionPlacement?: Record<string, unknown> };
  section.textDirection = 'btLr';
  section.__sectionPlacement = { ...(section.__sectionPlacement ?? { sectionId: 'section:0' }), sectionFlow: 'bottomToTop' };
  return doc;
}

function context(): CanvasRenderingContext2D {
  const ctx = { font: '12px serif', fontKerning: 'auto', letterSpacing: '0px',
    measureText(text: string) {
      const size = Number(/([\d.]+)px/u.exec(ctx.font)?.[1] ?? 12);
      const width = [...text].length * size * 0.6;
      return { width, fontBoundingBoxAscent: size * .8, fontBoundingBoxDescent: size * .2,
        actualBoundingBoxLeft: 0, actualBoundingBoxRight: width,
        actualBoundingBoxAscent: size * .8, actualBoundingBoxDescent: size * .2 } as TextMetrics;
    } };
  return ctx as unknown as CanvasRenderingContext2D;
}

/** Physical box of an overlay run (x/y are its physical top-left). */
function physicalBox(run: DocxTextRunInfo) {
  if (run.transform === undefined) return { left: run.x, right: run.x + run.w, top: run.y, bottom: run.y + run.h };
  if (run.transform === 'rotate(90deg)') return { left: run.x - run.h, right: run.x, top: run.y, bottom: run.y + run.w };
  if (run.transform === 'rotate(-90deg)') return { left: run.x, right: run.x + run.h, top: run.y - run.w, bottom: run.y };
  throw new Error(`unexpected overlay transform ${run.transform}`);
}

describe('nested text-box story ownership', () => {
  it.each([
    { label: 'horizontal', native: false, bodyTransform: undefined },
    { label: 'native BtoT', native: true, bodyTransform: 'rotate(-90deg)' },
  ])('acquires a header text box with its own page and section ($label)', ({ native, bodyTransform }) => {
    const parsed = parsedDocument();
    const doc = native ? asBottomToTop(parsed) : parsed;
    const layout = layoutDocument(doc, createLayoutServices(doc, { measureContext: context() }), { currentDateMs: 0 });
    expect(layout.pages).toHaveLength(2);
    for (const [pageIndex, bodyText] of ['One', 'Two'].entries()) {
      const runs = textRunsForPage(layout, pageIndex, { scale: 1 });
      const body = runs.filter(run => run.source?.story === 'body');
      const boxes = runs.filter(run => run.source?.story === 'textbox');
      expect(body.map(run => run.text).join(''), `page ${pageIndex}`).toBe(bodyText);
      for (const run of body) expect(run.transform).toBe(bodyTransform);
      // The header's own text box shows THIS page's number.
      const headerBox = boxes.filter(run => /^\d+$/u.test(run.text));
      expect(headerBox.map(run => run.text).join(''), `page ${pageIndex}`).toBe(String(pageIndex + 1));
      for (const run of headerBox) {
        // Horizontal and upright in the physical header band.
        expect(run.transform).toBeUndefined();
        expect(physicalBox(run).bottom).toBeLessThan(PAGE.height / 2);
      }
      // The body text box belongs to the first page only and stays on it.
      const bodyBox = boxes.filter(run => !/^\d+$/u.test(run.text));
      expect(bodyBox.map(run => run.text).join(''), `page ${pageIndex}`).toBe(pageIndex === 0 ? 'Box' : '');
      for (const run of bodyBox) {
        const box = physicalBox(run);
        expect(box.left >= 0 && box.top >= 0 && box.right <= PAGE.width && box.bottom <= PAGE.height).toBe(true);
      }
    }
  });
});

// ECMA-376 Part 1 §17.6.12: numbering restarts and formats apply to PAGE.
// One page keeps NUMPAGES constant during convergence, exposing a cached
// parent's masked absent-to-resolved destination dependency in nested boxes.
it.each([
  { kind: 'header' as const, pages: 1, format: 'decimal', numbers: ['12'] },
  { kind: 'footer' as const, pages: 1, format: 'decimal', numbers: ['12'] },
  { kind: 'header' as const, pages: 1, format: 'upperRoman', numbers: ['XII'] },
  { kind: 'footer' as const, pages: 2, format: 'upperRoman', numbers: ['XII', 'XIII'] },
])('converges restarted $kind fields through nested tables and shapes ($pages pages, $format)', fields => {
  const doc = parsedDocument(fields);
  const layout = layoutDocument(doc, createLayoutServices(doc, { measureContext: context() }), { currentDateMs: 0 });
  expect(layout.pages).toHaveLength(fields.pages);
  fields.numbers.forEach((number, index) => {
    // Text only: the inline drawings that hold the boxes are listed as well.
    const runs = textRunsForPage(layout, index, { scale: 1 }).filter(run => run.inlineContent === undefined);
    expect(runs.filter(run => run.source?.story === 'body').map(run => run.text)).toEqual([number]);
    expect(runs.filter(run => run.source?.story === fields.kind).map(run => run.text)).toEqual([number, number]);
    expect(runs.filter(run => run.source?.story === 'textbox').map(run => run.text)).toEqual([number, number, number]);
  });
});

// Detached rich text is source work on EVERY selected page, just like inline
// header/footer paragraphs: the selecting story earns its complete text box's
// whole closure, admitted at the acquisition seam before the box's paragraphs
// are debited. The signed margins (§17.6.11) isolate accounting from reserve
// geometry. Seeding the real ledger to a few spare misses makes the credit
// observable on three pages: without it (no admission, or a story weight that
// omits the detached box) the 110-paragraph box exhausts the ledger. A 300-page
// fixture could not observe it, because page-independent paragraph reuse keeps
// its misses far below the 25,000 floor. Unit coverage of the arithmetic lives
// in layout/pagination-work-budget.test.ts; blocking and progressive routes run
// one pagination generator, and progressive admission order is covered in
// layout/canonical-producer-real-model.test.ts.
it.each(['header', 'footer'] as const)('credits a repeated complete text box in a %s before its paragraphs are acquired', kind => {
  const pages = 3;
  const lines = 110;
  const p = (text: string, pageBreak = false) => '<w:p><w:pPr>'
    + (pageBreak ? '<w:pageBreakBefore/>' : '')
    + '<w:spacing w:before="0" w:after="0" w:line="200" w:lineRule="exact"/></w:pPr>'
    + `<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
  const box = textBox(Array.from({ length: lines }, () => p('h')).join(''), true)
    .replace('height:36pt', 'height:1400pt');
  const tag = kind === 'header' ? 'hdr' : 'ftr';
  const doc = parseFiles(new Map([
    ['word/document.xml', `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>`
      + Array.from({ length: pages }, (_, i) => p('b', i > 0)).join('')
      + `<w:sectPr><w:${kind}Reference w:type="default" r:id="story"/>`
      + '<w:pgSz w:w="4000" w:h="30000"/>'
      + '<w:pgMar w:top="-200" w:bottom="-200" w:left="200" w:right="200" w:header="100" w:footer="100"/>'
      + '</w:sectPr></w:body></w:document>'],
    [`word/${kind}.xml`, `<w:${tag} xmlns:w="${W}" xmlns:v="${V}"><w:p>${box}</w:p></w:${tag}>`],
    ['word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="story" Type="${R}/${kind}" Target="${kind}.xml"/></Relationships>`],
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="xml" ContentType="application/xml"/></Types>'],
  ]));
  // Unreferenced definitions cost no layout. They lift the source allowance
  // above its 25,000 floor so page-story credit raises the limit.
  const source = layoutSourceStore({ ...doc,
    endnotes: Array.from({ length: 800 }, (_, index) => ({ id: `unused-${index}`, content: [] })) });
  const createBudget = paginationWorkBudget.createPaginationWorkBudget;
  const budgets: { budget: ReturnType<typeof createBudget>; initialLimit: number }[] = [];
  const factory = vi.spyOn(paginationWorkBudget, 'createPaginationWorkBudget').mockImplementation(store => {
    const budget = createBudget(store);
    const initialLimit = budget.limit;
    // Keep the real ledger and its failure latch, leaving 16 misses: the body
    // fits, but the text box's paragraphs need their owner's page credit.
    for (let i = 0; i < initialLimit - 16; i++) budget.noteMiss();
    budgets.push({ budget, initialLimit });
    return budget;
  });
  try {
    const layout = layoutDocument(source, createLayoutServices(source, { measureContext: context() }), { currentDateMs: 0 });
    expect(layout.pages).toHaveLength(pages);
    expect(budgets).toHaveLength(1);
    expect(budgets[0]!.budget.misses).toBeGreaterThan(budgets[0]!.initialLimit);
    for (let page = 0; page < pages; page++) {
      const texts = textRunsForPage(layout, page, { scale: 1 })
        .filter(run => run.source?.story === 'textbox').map(run => run.text);
      expect(texts, `page ${page}`).toEqual(Array(lines).fill('h'));
    }
  } finally {
    factory.mockRestore();
  }
});

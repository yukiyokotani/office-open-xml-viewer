/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from '../wasm/docx_parser.js';
import { storeZip } from '../conformance/generate.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import { createLayoutServices } from '../layout-runtime.js';
import { layoutDocument } from '../document-layout.js';
import type { TableLayout } from './types.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const encoder = new TextEncoder();
const E0 = 72;
const E1 = 540;
interface Options {
  spacing?: number; border?: number; margin?: number | null; mode?: number | null;
  jc?: string; bidi?: boolean; indent?: number; auto?: boolean; widths?: number[];
  span?: boolean; verticalMerge?: boolean; cellBorder?: boolean; asymmetric?: boolean;
  rowSpacing?: boolean; nested?: boolean; header?: boolean; floating?: boolean;
  spacingScope?: 'row' | 'exception' | 'style';
  vertical?: boolean; long?: boolean; texts?: string[]; pageHeight?: number; rightBorderColor?: string;
}
const para = (text: string) => `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="368" w:lineRule="exact"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
function geometry(o: Options = {}, fragments?: TableLayout[]): TableLayout {
  const widths = o.widths ?? [2000, 2000, 2000];
  const border = (side: string, size = o.border ?? 24) => `<w:${side} w:val="${size ? 'single' : 'nil'}" w:sz="${size}" w:color="${side === 'right' ? o.rightBorderColor ?? '000000' : '000000'}"/>`;
  const tblPr = `<w:tblW w:w="${o.auto ? 0 : widths.reduce((a, b) => a + b, 0)}" w:type="${o.auto ? 'auto' : 'dxa'}"/>`
    + `<w:tblLayout w:type="${o.auto ? 'autofit' : 'fixed'}"/>`
    + (o.spacingScope === 'style' ? '<w:tblStyle w:val="SpacingDerived"/>'
      : o.spacingScope ? '' : `<w:tblCellSpacing w:w="${o.spacing ?? 200}" w:type="dxa"/>`)
    + `<w:jc w:val="${o.jc ?? 'left'}"/>${o.indent === undefined ? '' : `<w:tblInd w:w="${o.indent}" w:type="dxa"/>`}`
    + `${o.bidi ? '<w:bidiVisual/>' : ''}${o.floating ? '<w:tblpPr w:horzAnchor="page" w:vertAnchor="page" w:tblpX="1440" w:tblpY="1440"/>' : ''}`
    + `<w:tblBorders>${['top', 'bottom', 'left', 'right', 'insideH', 'insideV'].map((side) => border(side, o.asymmetric && side === 'right' ? 4 : undefined)).join('')}</w:tblBorders>`
    + (o.margin === null ? '' : `<w:tblCellMar>${['left', 'right', 'top', 'bottom'].map((side) => `<w:${side} w:w="${o.margin ?? 0}" w:type="dxa"/>`).join('')}</w:tblCellMar>`);
  const rows = [0, 1].map((row) => `<w:tr>${o.rowSpacing && row === 1 ? '<w:tblPrEx><w:tblCellSpacing w:w="60" w:type="dxa"/></w:tblPrEx>' : ''}`
    + (o.spacingScope === 'row' ? `<w:trPr><w:tblCellSpacing w:w="${o.spacing ?? 200}" w:type="dxa"/></w:trPr>`
      : o.spacingScope === 'exception' ? `<w:tblPrEx><w:tblCellSpacing w:w="${o.spacing ?? 200}" w:type="dxa"/></w:tblPrEx>` : '')
    + widths.flatMap((_width, col) => o.span && col === 1 ? [] : [`<w:tc><w:tcPr><w:tcW w:w="${o.auto ? 0 : o.span && col === 0 ? widths[0]! + widths[1]! : widths[col]}" w:type="${o.auto ? 'auto' : 'dxa'}"/>`
      + `${o.span && col === 0 ? '<w:gridSpan w:val="2"/>' : ''}${o.verticalMerge && col === 0 ? `<w:vMerge${row === 0 ? ' w:val="restart"' : ''}/>` : ''}`
      + `${o.cellBorder && col === 0 ? `<w:tcBorders>${border('left')}</w:tcBorders>` : ''}${o.vertical ? '<w:textDirection w:val="tbRl"/>' : ''}`
      + `</w:tcPr>${para(o.texts?.[col] ?? (o.long ? 'w '.repeat(100) : ['Q', 'bb', 'ccc'][col] ?? 'Q'))}</w:tc>`]).join('') + '</w:tr>').join('');
  const table = `<w:tbl><w:tblPr>${tblPr}</w:tblPr><w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${rows}</w:tbl>`;
  const outer = `<w:tbl><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc>${table}${para('end')}</w:tc></w:tr></w:tbl>`;
  const mode = o.mode === undefined ? 15 : o.mode;
  const files = new Map<string, string>([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'],
    ['_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="d" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="s" Type="${R}/settings" Target="settings.xml"/><Relationship Id="h" Type="${R}/header" Target="header1.xml"/>${o.spacingScope === 'style' ? `<Relationship Id="st" Type="${R}/styles" Target="styles.xml"/>` : ''}</Relationships>`],
    ['word/document.xml', `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${o.header ? para('body') : o.nested ? outer : table}${para('after')}<w:sectPr><w:headerReference w:type="default" r:id="h"/><w:pgSz w:w="12240" w:h="${o.pageHeight ?? 15840}"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440" w:header="360"/></w:sectPr></w:body></w:document>`],
    ['word/header1.xml', `<w:hdr xmlns:w="${W}">${o.header ? table : ''}${para('header')}</w:hdr>`],
    ['word/settings.xml', `<w:settings xmlns:w="${W}">${mode === null ? '' : `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/></w:compat>`}</w:settings>`],
  ]);
  if (o.spacingScope === 'style') files.set('word/styles.xml', `<w:styles xmlns:w="${W}">`
    + `<w:style w:type="table" w:styleId="SpacingBase"><w:tblPr><w:tblCellSpacing w:w="${o.spacing ?? 200}" w:type="dxa"/></w:tblPr></w:style>`
    + '<w:style w:type="table" w:styleId="SpacingDerived"><w:basedOn w:val="SpacingBase"/></w:style></w:styles>');
  const archive = new DocxArchive(storeZip(new Map([...files].map(([name, xml]) => [name, encoder.encode(xml)]))));
  try {
    const model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
    const context = { font: '10px serif', letterSpacing: '0px', fontKerning: 'none', measureText: (text: string) => ({
      width: [...text].length * 6, fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
      actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
    }) } as unknown as CanvasRenderingContext2D;
    const layout = layoutDocument(model, createLayoutServices(model, { measureContext: context }), { currentDateMs: 0 });
    const found: TableLayout[] = [];
    const seen = new Set<object>();
    const walk = (value: unknown): void => {
      if (value === null || typeof value !== 'object' || seen.has(value)) return;
      seen.add(value);
      if ((value as TableLayout).kind === 'table') found.push(value as TableLayout);
      Object.values(value).forEach(walk);
    };
    walk(layout.pages[0]?.layers);
    const result = found.at(-1);
    if (!result) throw new Error('No retained table');
    if (fragments) {
      layout.pages.slice(1).forEach((page) => walk(page.layers));
      fragments.push(...found);
    }
    return result;
  } finally { archive.free(); }
}
beforeAll(async () => { await init({ module_or_path: await readFile(new URL('../wasm/docx_parser_bg.wasm', import.meta.url)) }); });
const right = (rect: { xPt: number; widthPt: number }) => rect.xPt + rect.widthPt;
const bottom = (rect: { yPt: number; heightPt: number }) => rect.yPt + rect.heightPt;
const fingerprint = (values: Options[]) => createHash('sha256').update(JSON.stringify(values.map((value) => geometry(value)))).digest('hex');

describe('WORD_TABLE_CELL_SPACING_GEOMETRY (#2066)', () => {
  it.each([[60, 24], [200, 4], [200, 0], [200, 24]])('reserves 2s gaps, cell border boxes and vertical outer borders (s=%s, sz=%s)', (spacing, border) => {
    const table = geometry({ spacing, border, margin: 40 });
    const s = spacing / 20; const t = border / 8; const m = 2;
    const [first, second] = table.rows[0]!.cells;
    const last = table.rows[1]!.cells[0]!;
    expect(first!.flowBounds.xPt - table.flowBounds.xPt).toBeCloseTo(2 * s, 8);
    expect(second!.flowBounds.xPt - right(first!.flowBounds)).toBeCloseTo(2 * s, 8);
    expect(right(table.flowBounds) - right(table.rows[0]!.cells[2]!.flowBounds)).toBeCloseTo(2 * s, 8);
    expect(first!.contentBounds.xPt - first!.flowBounds.xPt).toBeCloseTo(t + m, 8);
    expect(first!.contentBounds.widthPt).toBeCloseTo((300 - 8 * s) / 3 - 2 * (t + m), 8);
    expect(first!.contentBounds.yPt - table.flowBounds.yPt).toBeCloseTo(2 * s + 2 * t + m, 8);
    expect(last.contentBounds.yPt - first!.contentBounds.yPt).toBeCloseTo(18.4 + 2 * s + 2 * t + 2 * m, 8);
    expect(bottom(table.flowBounds) - bottom(last.contentBounds)).toBeCloseTo(2 * s + 2 * t + m, 8);
    expect(last.flowBounds.yPt - bottom(first!.flowBounds)).toBeCloseTo(2 * s, 8);
    // Border strokes lie inside the independent cell boxes.
    if (t > 0) expect(table.borders.some((b) => Math.abs(b.from.xPt - (first!.flowBounds.xPt + t / 2)) < 1e-8)).toBe(true);
  });
  it('retains each border side colour while using uniform-width box geometry', () => {
    const table = geometry({ rightBorderColor: 'FF0000' });
    expect(table.borders.map((border) => border.color.toLowerCase())).toContain('#ff0000');
    expect(table.rows[0]!.cells[0]!.contentBounds.xPt - table.flowBounds.xPt).toBeCloseTo(23, 8);
  });
  it('carves fixed gaps proportionally, including horizontal spans', () => {
    const table = geometry({ widths: [1000, 2000, 3000], span: true });
    const [span, last] = table.rows[0]!.cells;
    expect(span!.flowBounds.widthPt).toBeCloseTo(220 * 0.5 + 20, 8);
    expect(last!.flowBounds.widthPt).toBeCloseTo(220 * 0.5, 8);
  });
  it('places mode-15 boxes at the measured border centre with bidi shifts and ignores nonleading indents', () => {
    for (const bidi of [false, true]) for (const jc of ['left', 'start', 'center', 'right', 'end']) {
      const table = geometry({ bidi, jc, indent: 720 });
      const leading = jc === 'left' || jc === 'start';
      const expected = jc === 'center' ? (E0 + E1 - 300) / 2
        : (leading !== bidi) ? E0 + (leading ? 36 : 0) + 1.5 : E1 - 300 - (leading ? 36 : 0) - 1.5;
      expect(table.flowBounds.xPt, `${jc} ${bidi}`).toBeCloseTo(expected - (bidi ? 1.5 : 0), 8);
      if (!leading) expect(geometry({ bidi, jc }).flowBounds).toEqual(table.flowBounds);
    }
  });
  it('hangs the whole content inset in legacy LTR and leaves exactly s in bidi', () => {
    for (const mode of [null, 14]) for (const bidi of [false, true]) {
      const leading = geometry({ mode, bidi, indent: 720, margin: 108 });
      const cell = leading.rows[0]!.cells[0]!;
      expect(bidi ? right(cell.contentBounds) : cell.contentBounds.xPt).toBeCloseTo(bidi ? E1 - 36 - 10 : E0 + 36, 8);
      const end = geometry({ mode, bidi, jc: 'end', indent: 720, margin: 108 });
      const last = end.rows[0]!.cells[2]!;
      expect(bidi ? last.contentBounds.xPt : right(last.contentBounds)).toBeCloseTo(bidi ? E0 + 10 : E1, 8);
      const plain = geometry({ mode, bidi, margin: 108 });
      expect(plain.flowBounds.xPt).toBeCloseTo(bidi ? E1 - 300 : E0, 8);
      const center = geometry({ mode, bidi, jc: 'center', indent: 720 });
      expect(center.flowBounds.xPt).toBeCloseTo((E0 + E1 - 300) / 2, 8);
    }
  });
  it('treats omitted margins as zero without approximating Word quanta', () => {
    expect(geometry({ margin: null })).toEqual(geometry({ margin: 0 }));
  });
  it('content-fits three unpreferred AutoFit tracks at every alignment and measured mode', () => {
    for (const mode of [null, 14, 15]) for (const bidi of [false, true]) for (const jc of ['left', 'center', 'end']) {
      const table = geometry({ auto: true, mode, bidi, jc, margin: 108 });
      expect(table.flowBounds.widthPt, `${mode} ${bidi} ${jc}`).toBeCloseTo(6 + 12 + 18 + 3 * (6 + 10.8) + 80, 8);
      table.rows[0]!.cells.forEach((cell, index) => expect(cell.contentBounds.widthPt).toBeCloseTo((index + 1) * 6, 8));
    }
  });
  it('includes every gap in the nonbinding AutoFit ceiling', () => {
    const fits = geometry({ auto: true, texts: Array(3).fill('w '.repeat(10)) });
    // Three 120pt maxima, each with two 3pt borders; four 20pt gaps.
    expect(fits.flowBounds.widthPt).toBeCloseTo(458, 8);
  });
  it('keeps binding spaced AutoFit byte-identical to the base renderer', () => {
    const binding = { auto: true, texts: Array(3).fill('w '.repeat(11)) };
    // Maxima + borders + gaps need 494pt, exceeding the 468pt band. The
    // entire retained layout must use the base path, including vertical flow.
    expect(geometry(binding).flowBounds.widthPt).toBe(300);
    // Captured at 5b6d291540f122570b801086a52b43e6718b6fbe with this helper.
    expect(fingerprint([binding])).toBe('e1bf5a08d76e3b3a8f07a79d22938e09add612f9ace1af555ed17d6da04c90ce');
    expect(fingerprint([{ auto: true, widths: [5000, 5000, 5000], long: true }]))
      .toBe('d62e5a2c96040b2b97c9950b751a0d3317dfa4adeace656b0b24e56391c3ed94');
  });
  // Captured from the unmodified stacked parent cb1746fad^ using these same
  // synthetic inputs. These guard the explicit byte-identical retained-layout
  // contract, including borders, paragraphs and row/page flow, not a test count.
  it('reserves both cell borders and doubled outer gaps before splitting a row', () => {
    const table = geometry({ long: true, pageHeight: 6000 });
    // A 300pt page with 72pt margins owns a 156pt body. Child selection must
    // reserve the same border/gap geometry that materialization later retains.
    expect(table.flowBounds.heightPt).toBeLessThanOrEqual(156);
    expect(bottom(table.flowBounds)).toBeLessThanOrEqual(228);
  });
  it('reserves standalone outer geometry before admitting a whole row', () => {
    const fragments: TableLayout[] = [];
    geometry({ texts: Array(3).fill('w '.repeat(40)), pageHeight: 6280 }, fragments);
    // The original 167.8pt track fits the 170pt body, but its standalone
    // fragment needs 180.8pt. It must split instead of failing convergence.
    expect(fragments.length).toBeGreaterThan(1);
    for (const table of fragments) {
      expect(table.flowBounds.heightPt).toBeLessThanOrEqual(170);
      expect(bottom(table.flowBounds)).toBeLessThanOrEqual(242);
    }
  });
  it('keeps the previous geometry when gaps alone would exhaust the table or band', () => {
    for (const auto of [false, true]) {
      const table = geometry({ spacing: 2000, auto });
      expect(table.flowBounds.xPt).toBe(E0);
      expect(table.rows[0]!.cells[0]!.flowBounds.xPt - table.flowBounds.xPt).toBe(100);
    }
  });
  it('keeps all no-spacing retained geometry byte-identical to the previous renderer', () => {
    const cases = [null, 11, 12, 14, 15].flatMap((mode) => [false, true].flatMap((bidi) => ['left', 'center', 'end'].flatMap((jc) =>
      [false, true].map((auto) => ({ spacing: 0, mode, bidi, jc, auto, indent: 720, margin: 108 })))));
    expect(fingerprint(cases)).toBe('27c5d15368ef440dec14dd99e2d1ba85a44e1290f3e49641b8110179f573290e');
  });
  it('keeps unmeasured classes byte-identical to the previous renderer', () => {
    expect(fingerprint([{ asymmetric: true }, { cellBorder: true }, { rowSpacing: true }, { nested: true },
      { header: true }, { floating: true }, { vertical: true }, { verticalMerge: true }])).toBe('28a6cf890f22e943b52dc511ea76b9e98f25e5bae672811f81863286987f92d2');
  });
  it.each(['row', 'exception'] as const)('keeps uniform %s spacing byte-identical to the base renderer', (spacingScope) => {
    expect(geometry({ spacingScope }).flowBounds.heightPt).toBeCloseTo(66.8, 8);
    // Captured independently for both sources at 5b6d291540f122570b801086a52b43e6718b6fbe.
    expect(fingerprint([{ spacingScope }])).toBe('1e2aaa8dc6f1d43b1ea27dbc6d7cc73ebc207c5a57dade938ac26e89d2cafbb6');
  });
  it('keeps style-inherited table-level spacing in the measured class', () => {
    expect(geometry({ spacingScope: 'style' })).toEqual(geometry());
  });
  it('places modes 11 and 12 like mode 14 and moves a leading table by a negative indent', () => {
    for (const jc of ['left', 'center', 'right']) {
      for (const bidi of [false, true]) {
        const at14 = geometry({ mode: 14, jc, bidi, indent: 720, margin: 108 });
        for (const mode of [11, 12]) {
          expect(geometry({ mode, jc, bidi, indent: 720, margin: 108 }), `${mode} ${jc} ${bidi}`).toEqual(at14);
        }
      }
    }
    for (const mode of [null, 14, 15]) {
      const none = geometry({ mode, indent: 0 }).flowBounds.xPt;
      expect(geometry({ mode, indent: -360 }).flowBounds.xPt - none, `${mode}`).toBeCloseTo(-18, 8);
    }
  });
});

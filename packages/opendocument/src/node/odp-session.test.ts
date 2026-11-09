// OpenDocument presentations through the Node PPTX session: the claimed
// package is projected into the PPTX slide model (master objects, styles,
// lists, custom shapes, tables), or rejected as a whole only when corrupt or
// encrypted.
import { describe, expect, it } from 'vitest';
import { buildOdpFixture } from '../odp-fixture.js';
import { testOdpSource } from '../test-sources.js';
import { materializePptxPresentation } from './node-facade.js';

type Run = { type: string; text?: string; fontSize?: number; bold?: boolean; color?: string; fontFamily?: string; fieldType?: string };
type Paragraph = { runs: Run[]; marL: number; indent: number; bullet: { type: string; char?: string }; alignment: string };
type Element = {
  type: string; x: number; y: number; width: number; height: number; rotation: number;
  fill?: { fillType: string; color?: string; angle?: number; stops?: { position: number; color: string }[] };
  stroke?: { color: string; width: number };
  textBody?: { paragraphs: Paragraph[]; verticalAnchor: string; lIns: number };
  custGeom?: { cmd: string; x?: number; y?: number }[][];
  textRect?: { x: number; y: number; width: number; height: number };
  rows?: { height: number; cells: { textBody?: { paragraphs: Paragraph[] }; fill?: { color: string }; gridSpan: number; hMerge: boolean }[] }[];
  cols?: number[];
};
type Model = { slideWidth: number; slideHeight: number; slides: { background?: { fillType: string; color?: string }; elements: Element[]; notes?: string; hidden?: boolean }[] };

const open = async (bytes: Uint8Array) =>
  (await materializePptxPresentation(bytes, { odp: testOdpSource() })) as unknown as Model;

const cm = (v: number) => Math.round(v * 360000);

describe('ODP through the PPTX Node session', () => {
  it('maps the page size, master objects, backgrounds and frame text', async () => {
    const model = await open(buildOdpFixture({
      masterContent: '<draw:frame presentation:class="title" svg:x="1cm" svg:y="1cm" svg:width="5cm" svg:height="2cm"><draw:text-box><text:p>template</text:p></draw:text-box></draw:frame>'
        + '<draw:frame presentation:class="page-number" svg:x="20cm" svg:y="14cm" svg:width="2cm" svg:height="1cm"><draw:text-box><text:p><text:page-number>&lt;n&gt;</text:page-number></text:p></draw:text-box></draw:frame>'
        + '<draw:custom-shape draw:style-name="Mgr1" svg:x="0cm" svg:y="0cm" svg:width="28cm" svg:height="1cm"><draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 Z N"/></draw:custom-shape>',
      styles: '<style:style style:name="Mgr1" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#336699" draw:stroke="none"/></style:style>',
      automaticStyles: '<style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties draw:fill="solid" draw:fill-color="#ffeedd"/></style:style>'
        + '<style:style style:name="P1" style:family="paragraph"><style:paragraph-properties fo:text-align="center"/></style:style>'
        + '<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold" fo:font-size="24pt" fo:color="#ff0000"/></style:style>',
      pages: '<draw:page draw:name="One" draw:style-name="dp1" draw:master-page-name="M">'
        + '<draw:frame svg:x="2cm" svg:y="3cm" svg:width="10cm" svg:height="4cm"><draw:text-box><text:p text:style-name="P1">Hello <text:span text:style-name="T1">bold</text:span><text:s text:c="2"/>end</text:p></draw:text-box></draw:frame>'
        + '<presentation:notes><draw:frame presentation:class="notes"><draw:text-box><text:p>speaker</text:p></draw:text-box></draw:frame></presentation:notes>'
        + '</draw:page>',
    }));
    expect(model.slideWidth).toBe(cm(28));
    expect(model.slideHeight).toBe(cm(15.75));
    const slide = model.slides[0]!;
    expect(slide.background).toEqual({ fillType: 'solid', color: 'FFEEDD' });
    expect(slide.notes).toBe('speaker');
    // The title template is not drawn; the page number and the band are.
    const texts = slide.elements.map(e => e.textBody?.paragraphs.flatMap(p => p.runs.map(r => r.text)).join('') ?? '');
    expect(texts).not.toContain('template');
    const band = slide.elements.find(e => e.fill?.color === '336699');
    expect(band).toMatchObject({ x: 0, width: cm(28), height: cm(1) });
    const number = slide.elements.find(e => e.textBody?.paragraphs[0]?.runs[0]?.fieldType === 'slidenum');
    expect(number?.textBody?.paragraphs[0]?.runs[0]?.text).toBe('1');
    const frame = slide.elements.find(e => e.x === cm(2))!;
    const paragraph = frame.textBody!.paragraphs[0]!;
    expect(paragraph.alignment).toBe('ctr');
    expect(paragraph.runs.map(r => r.text)).toEqual(['Hello ', 'bold', '  ', 'end']);
    expect(paragraph.runs[1]).toMatchObject({ bold: true, fontSize: 24, color: 'FF0000' });
    expect(paragraph.runs[0]).toMatchObject({ fontSize: 18, fontFamily: 'Liberation Sans', color: '000000' });
  });

  it('evaluates custom-shape equations, mirrors and text areas', async () => {
    const model = await open(buildOdpFixture({
      pages: '<draw:page draw:master-page-name="M">'
        + '<draw:custom-shape svg:width="4cm" svg:height="2cm" draw:transform="rotate (-1.5707963267949) translate (5cm 6cm)">'
        + '<text:p>t</text:p>'
        + '<draw:enhanced-geometry svg:viewBox="0 0 0 0" draw:mirror-horizontal="true" draw:modifiers="25000" draw:text-areas="0 0 ?f1 logheight" drawooo:enhanced-path="M 0 ?f0 G ?f0 ?f0 180 90 L ?f1 0 Z N" draw:enhanced-path="M 0 0 Z N">'
        + '<draw:equation draw:name="f0" draw:formula="min(logwidth,logheight)*$0 /100000"/><draw:equation draw:name="f1" draw:formula="logwidth-?f0 "/></draw:enhanced-geometry></draw:custom-shape>'
        + '</draw:page>',
    }));
    const shape = model.slides[0]!.elements[0]!;
    expect(shape.rotation).toBeCloseTo(90, 6);
    expect(shape).toMatchObject({ width: cm(4), height: cm(2), flipH: true });
    // rotate about the origin then translate: the centre lands at (4cm, 8cm).
    expect(shape.x + shape.width / 2).toBeCloseTo(cm(4), -1);
    expect(shape.y + shape.height / 2).toBeCloseTo(cm(8), -1);
    expect(shape.custGeom![0]![0]).toEqual({ cmd: 'moveTo', x: 0, y: 0.25 });
    expect(shape.custGeom![0]!.at(-2)).toEqual({ cmd: 'lineTo', x: 0.875, y: 0 });
    // A horizontal mirror places the text area from the right.
    expect(shape.textRect?.width).toBe(cm(3.5));
  });

  it('labels list levels and numbers with the list-style indents', async () => {
    const list = '<text:list-style style:name="L1">'
      + '<text:list-level-style-bullet text:level="1" text:bullet-char="•"><style:list-level-properties text:min-label-width="0.6cm"/><style:text-properties fo:font-size="80%"/></text:list-level-style-bullet>'
      + '<text:list-level-style-number text:level="2" style:num-format="i" style:num-suffix="."><style:list-level-properties text:space-before="1cm" text:min-label-width="0.5cm"/></text:list-level-style-number>'
      + '</text:list-style>';
    const model = await open(buildOdpFixture({
      automaticStyles: list,
      pages: '<draw:page draw:master-page-name="M"><draw:frame svg:x="0cm" svg:y="0cm" svg:width="10cm" svg:height="5cm"><draw:text-box>'
        + '<text:list text:style-name="L1"><text:list-item><text:p>a</text:p><text:list><text:list-item><text:p>b</text:p></text:list-item><text:list-item><text:p>c</text:p></text:list-item></text:list></text:list-item></text:list>'
        + '<text:p>plain</text:p></draw:text-box></draw:frame></draw:page>',
    }));
    const paragraphs = model.slides[0]!.elements[0]!.textBody!.paragraphs;
    expect(paragraphs[0]).toMatchObject({ marL: cm(0.6), indent: -cm(0.6), bullet: { type: 'char', char: '•', sizePct: 80 } });
    expect(paragraphs[1]).toMatchObject({ marL: cm(1.5), indent: -cm(0.5), bullet: { char: 'i.' } });
    expect(paragraphs[2]?.bullet).toMatchObject({ char: 'ii.' });
    expect(paragraphs[3]).toMatchObject({ marL: 0, bullet: { type: 'none' } });
  });

  it('maps tables with spans, cell fills and borders', async () => {
    const model = await open(buildOdpFixture({
      automaticStyles: '<style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="3cm"/></style:style>'
        + '<style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="1cm"/></style:style>'
        + '<style:style style:name="ce1" style:family="table-cell"><loext:graphic-properties draw:fill="solid" draw:fill-color="#bfbfbf"/><style:paragraph-properties fo:border="0.5pt solid #000000"/></style:style>',
      pages: '<draw:page draw:master-page-name="M"><draw:frame svg:x="1cm" svg:y="1cm" svg:width="6cm" svg:height="2cm"><table:table>'
        + '<table:table-column table:style-name="co1" table:number-columns-repeated="2"/>'
        + '<table:table-row table:style-name="ro1"><table:table-cell table:style-name="ce1" table:number-columns-spanned="2"><text:p>wide</text:p></table:table-cell><table:covered-table-cell/></table:table-row>'
        + '<table:table-row table:style-name="ro1"><table:table-cell><text:p>a</text:p></table:table-cell><table:table-cell><text:p>b</text:p></table:table-cell></table:table-row>'
        + '</table:table></draw:frame></draw:page>',
    }));
    const table = model.slides[0]!.elements[0]!;
    expect(table.type).toBe('table');
    expect(table.cols).toEqual([cm(3), cm(3)]);
    expect(table.rows![0]!.cells[0]).toMatchObject({ gridSpan: 2, fill: { color: 'BFBFBF' }, borderL: { width: 6350, color: '000000' } });
    expect(table.rows![0]!.cells[1]!.hMerge).toBe(true);
    expect(table.rows![1]!.cells[1]!.textBody!.paragraphs[0]!.runs[0]!.text).toBe('b');
  });

  it('reads chart objects from their own data table', async () => {
    const ns = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"';
    const row = (label: string, value: number) => `<table:table-row><table:table-cell office:value-type="string"><text:p>${label}</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="${value}"><text:p>${value}</text:p></table:table-cell></table:table-row>`;
    const chart = `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${ns} office:version="1.3"><office:body><office:chart>`
      + '<chart:chart chart:class="chart:bar" svg:width="10cm" svg:height="6cm"><chart:plot-area table:cell-range-address="local-table.$A$1:.$B$3">'
      + '<chart:axis chart:dimension="x" chart:name="primary-x"><chart:categories table:cell-range-address="local-table.$A$2:.$A$3"/></chart:axis><chart:axis chart:dimension="y" chart:name="primary-y"/>'
      + '<chart:series chart:class="chart:bar" chart:values-cell-range-address="local-table.$B$2:.$B$3" chart:label-cell-address="local-table.$B$1"/></chart:plot-area>'
      + '<table:table table:name="local-table"><table:table-header-columns><table:table-column/></table:table-header-columns><table:table-columns><table:table-column/></table:table-columns>'
      + '<table:table-header-rows><table:table-row><table:table-cell><text:p/></table:table-cell><table:table-cell office:value-type="string"><text:p>S1</text:p></table:table-cell></table:table-row></table:table-header-rows>'
      + `<table:table-rows>${row('a', 3)}${row('b', 5)}</table:table-rows></table:table></chart:chart></office:chart></office:body></office:document-content>`;
    const model = await open(buildOdpFixture({
      pages: '<draw:page draw:master-page-name="M"><draw:frame svg:x="1cm" svg:y="1cm" svg:width="10cm" svg:height="6cm"><draw:object xlink:href="./Object 1"/><draw:image xlink:href="./ObjectReplacements/Object 1"/></draw:frame></draw:page>',
      extraFiles: [['Object 1/content.xml', chart]],
    }));
    const element = model.slides[0]!.elements[0] as unknown as { type: string; x: number; chart: { series: { values?: unknown[] }[] } };
    expect(element.type).toBe('chart');
    expect(element.x).toBe(cm(1));
    expect(element.chart.series).toHaveLength(1);
  });

  it('maps gradients, markers and hidden slides; rejects encrypted packages', async () => {
    const model = await open(buildOdpFixture({
      styles: '<draw:gradient draw:name="g1" draw:style="linear" draw:start-color="#000000" draw:end-color="#ffffff" draw:angle="0deg" draw:border="0%"/>'
        + '<draw:marker draw:name="Arrow" svg:viewBox="0 0 20 30" svg:d="M10 0l-10 30h20z"/>',
      automaticStyles: '<style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:fill="gradient" draw:fill-gradient-name="g1" draw:stroke="none"/></style:style>'
        + '<style:style style:name="gr2" style:family="graphic"><style:graphic-properties draw:stroke="solid" svg:stroke-width="0.1cm" svg:stroke-color="#ff0000" draw:marker-end="Arrow" draw:marker-end-width="0.4cm"/></style:style>'
        + '<style:style style:name="dp2" style:family="drawing-page"><style:drawing-page-properties presentation:visibility="hidden"/></style:style>',
      pages: '<draw:page draw:master-page-name="M"><draw:rect draw:style-name="gr1" svg:x="0cm" svg:y="0cm" svg:width="2cm" svg:height="2cm"/>'
        + '<draw:line draw:style-name="gr2" svg:x1="1cm" svg:y1="5cm" svg:x2="6cm" svg:y2="5cm"/></draw:page>'
        + '<draw:page draw:style-name="dp2" draw:master-page-name="M"/>',
    }));
    const [rect, line, arrow] = model.slides[0]!.elements;
    expect(rect!.fill).toMatchObject({ fillType: 'gradient', angle: 90, stops: [{ position: 0, color: '000000' }, { position: 1, color: 'FFFFFF' }] });
    expect(line!.stroke).toMatchObject({ color: 'FF0000', width: cm(0.1) });
    // The 0.4cm × 0.6cm arrow's tip sits on the line end; the line stops under it.
    expect(arrow!.fill).toEqual({ fillType: 'solid', color: 'FF0000' });
    expect(arrow!.x + arrow!.width).toBe(cm(6));
    expect(arrow!.width).toBe(cm(0.6));
    expect(line!.x + line!.width).toBe(cm(6));
    expect((line!.custGeom![0]!.at(-1) as { x: number }).x).toBeCloseTo(4.4 / 5, 6);
    expect(model.slides[1]!.hidden).toBe(true);
    const encrypted = buildOdpFixture({
      pages: '<draw:page/>',
      manifestExtra: '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"><manifest:encryption-data/></manifest:file-entry>',
    });
    await expect(open(encrypted)).rejects.toMatchObject({ code: 'unsupported-encryption' });
  });
});

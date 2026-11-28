import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it, vi } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import { textRunsForPage } from './text-run-projection.js';
import { buildDocxHighlightLayer } from './find-highlight-layer.js';
import type { DocxTextRunInfo } from './types.js';

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const encode = (text: string) => new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${text}`);

/** Every character advances half an em (6pt at 12pt). */
function measureContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  const size = (): number => Number.parseFloat(/([\d.]+)px/u.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px', fontKerning: 'normal',
    measureText: (value: string) => ({
      width: [...value].length * size() * 0.5,
      fontBoundingBoxAscent: size() * 0.8, fontBoundingBoxDescent: size() * 0.2,
      actualBoundingBoxAscent: size() * 0.8, actualBoundingBoxDescent: size() * 0.2,
    }),
  } as unknown as CanvasRenderingContext2D;
}

function runsOf(paragraph: string, vertical: boolean, sectionExtra = '', textWidthTwips = 9360): DocxTextRunInfo[] {
  const section = `<w:pgSz w:w="${textWidthTwips + 2880}" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" `
    + 'w:header="0" w:footer="0" w:gutter="0"/>' + (vertical ? '<w:textDirection w:val="tbRl"/>' : '') + sectionExtra;
  const bytes = storeZip(new Map([
    ['[Content_Types].xml', encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')],
    ['_rels/.rels', encode(`<Relationships xmlns="${PKG}"><Relationship Id="rId1" `
      + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/_rels/document.xml.rels', encode(`<Relationships xmlns="${PKG}"><Relationship Id="rIdT" `
      + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')],
    ['word/settings.xml', encode(`<w:settings ${NS}><w:compat><w:compatSetting w:name="compatibilityMode" `
      + 'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')],
    ['word/document.xml', encode(`<w:document ${NS}><w:body>${paragraph}<w:sectPr>${section}</w:sectPr></w:body></w:document>`)],
  ]));
  const archive = new DocxArchive(bytes);
  let parsed;
  try {
    parsed = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse())));
  } finally { archive.free(); }
  const model = parsed.document;
  const layout = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext() }), { currentDateMs: 0 });
  return textRunsForPage(layout, 0, { scale: 1 });
}

type Box = Readonly<{ left: number; top: number; width: number; height: number }>;
/** The find highlight of each character of `run`, in page px (a 1000px page). */
function highlights(run: DocxTextRunInfo): Box[] {
  const element = () => ({ style: { cssText: '' }, children: [] as unknown[], appendChild(child: unknown) { this.children.push(child); } });
  vi.stubGlobal('document', { createElement: element });
  try {
    return [...run.text].map((_, index) => {
      const layer = element();
      buildDocxHighlightLayer(layer as never, [run], [{ active: false, slices: [{ runIndex: 0, start: index, end: index + 1 }] }],
        1000, 1000, () => (value: string) => [...value].length * 6);
      const css = (layer.children[0] as { style: { cssText: string } }).style.cssText;
      const value = (name: string) => Number(new RegExp(`${name}:([-\\d.e]+)%`, 'u').exec(css)![1]) * 10;
      return { left: value('left'), top: value('top'), width: value('width'), height: value('height') };
    });
  } finally {
    vi.unstubAllGlobals();
  }
}

const close = (value: number, expected: number) => expect(value).toBeCloseTo(expected, 6);

describe('find highlights cover the glyph box: the natural advance from the pen, without tracking', () => {
  const trackings = [-40, 0, 160] as const;
  const kinds = ['LTR', 'RTL', 'tate-chu-yoko', 'fitText'] as const;
  const cases = trackings.flatMap((tracking) => kinds.map((kind) => [kind, tracking] as const));
  it.each(cases)('%s with w:spacing %i', (kind, tracking) => {
    const trackPt = tracking / 20;
    const spacing = `<w:spacing w:val="${tracking}"/><w:sz w:val="24"/>`;
    const text = kind === 'RTL' ? 'אבג' : kind === 'tate-chu-yoko' ? '29' : 'abc';
    const rPr = kind === 'RTL' ? `<w:rtl/>${spacing}`
      : kind === 'tate-chu-yoko' ? `${spacing}<w:eastAsianLayout w:id="1" w:vert="1"/>`
        : kind === 'fitText' ? `${spacing}<w:fitText w:val="480" w:id="1"/>` : spacing;
    const pPr = kind === 'RTL' ? '<w:pPr><w:bidi/></w:pPr>' : '';
    const runs = runsOf(`<w:p>${pPr}<w:r><w:rPr>${rPr}</w:rPr><w:t>${text}</w:t></w:r></w:p>`, kind === 'tate-chu-yoko');
    const run = runs.find((candidate) => candidate.text === text)!;
    const found = highlights(run);
    if (kind === 'tate-chu-yoko') {
      // Upright side by side across the run: each glyph 6pt across, one
      // tracked advance apart, over the run's cell along it.
      for (const box of found) close(box.height, 6);
      const [first, second] = found as [Box, Box];
      close(Math.hypot(second.left - first.left, second.top - first.top), 6 + trackPt);
      return;
    }
    // Each glyph is 6pt wide from its pen; pens are one advance apart.
    for (const box of found) close(box.width, 6);
    // fitText (24pt) over three 6pt glyphs replaces w:spacing with the gap
    // that fills the region: (24 - 18) / 2 between glyphs, none after the last.
    const pitch = kind === 'fitText' ? (24 - 18) / 2 : trackPt;
    const lefts = found.map((box) => box.left);
    if (kind === 'RTL') {
      // Read right to left: the first glyph ends at the run's right edge.
      close(lefts[0]! + 6, run.x + run.w);
      for (let index = 1; index < lefts.length; index += 1) close(lefts[index - 1]! - lefts[index]!, 6 + pitch);
    } else {
      close(lefts[0]!, run.x);
      if (kind === 'fitText') close(lefts.at(-1)! + 6, run.x + 24);
      for (let index = 1; index < lefts.length; index += 1) close(lefts[index]! - lefts[index - 1]!, 6 + pitch);
    }
  });
});

describe('find highlights leave out what is added between glyphs', () => {
  /** Every character's highlight of every run whose text is in `texts`. */
  const allHighlights = (runs: DocxTextRunInfo[], keep: (run: DocxTextRunInfo) => boolean) =>
    runs.filter(keep).flatMap((run) => highlights(run).map((box, index) => ({ run, index, box })));
  const ordered = (found: ReturnType<typeof allHighlights>) => {
    // Glyph boxes of one run never overlap: each starts at or after the
    // previous one's end, in paint order.
    const byRun = new Map<DocxTextRunInfo, Box[]>();
    for (const { run, box } of found) byRun.set(run, [...(byRun.get(run) ?? []), box]);
    for (const boxes of byRun.values()) {
      const sorted = [...boxes].sort((left, right) => left.left - right.left);
      for (let index = 1; index < sorted.length; index += 1) {
        expect(sorted[index]!.left).toBeGreaterThanOrEqual(sorted[index - 1]!.left + sorted[index - 1]!.width - 1e-6);
      }
    }
  };

  it('distributed CJK text: each highlight is its 6pt glyph, the distributed space left out', () => {
    // 31pt band (620tw), jc=distribute: lines are spread to the band.
    const text = '\u6f22\u5b57\u6771\u4eac\u5927\u962a\u4eac\u90fd';
    const runs = runsOf(`<w:p><w:pPr><w:jc w:val="distribute"/></w:pPr><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`,
      false, '', 620);
    const found = allHighlights(runs, (run) => /[\u4e00-\u9fff]/u.test(run.text));
    expect(found).toHaveLength([...text].length);
    for (const { box } of found) close(box.width, 6);
    ordered(found);
    // Some line is spread: there are gaps between its glyphs.
    const gaps = runs.filter((run) => run.text.length > 1).some((run) => run.w > run.text.length * 6 + 1e-6);
    expect(gaps).toBe(true);
  });

  it.each([
    ['CJK', '\u6f22\u5b57'],
    ['Latin', 'ab'],
  ])('%s text on a snapToChars character grid: each highlight is its 6pt glyph, the cell padding left out', (_, text) => {
    const grid = '<w:docGrid w:type="snapToChars" w:linePitch="360" w:charSpace="8192"/>';
    const runs = runsOf(`<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`, false, grid);
    const found = allHighlights(runs, (run) => run.text.length > 0 && [...run.text].every((character) => text.includes(character)));
    expect(found).toHaveLength([...text].length);
    for (const { box } of found) close(box.width, 6);
    ordered(found);
  });

  it('kashida-justified Arabic: each highlight is its 6pt glyph, the elongation left out', () => {
    const word = '\u0628\u062a\u062a\u0628\u062a';
    const text = Array.from({ length: 12 }, () => word).join(' ');
    const runs = runsOf(`<w:p><w:pPr><w:bidi/><w:jc w:val="highKashida"/></w:pPr><w:r><w:rPr><w:rtl/><w:sz w:val="24"/></w:rPr>`
      + `<w:t>${text}</w:t></w:r></w:p>`, false, '', 2600);
    const found = allHighlights(runs, (run) => /[\u0600-\u06ff]/u.test(run.text))
      .filter(({ run, index }) => run.text[index] !== ' ');
    expect(found.length).toBeGreaterThan(0);
    for (const { box } of found) close(box.width, 6);
    ordered(found);
  });
});

describe('find highlights only what its clips show', () => {
  it('gives no highlight to text painted below an exact-height row', () => {
    // 12pt text in a 12pt exact row, its line exactly 40pt: the glyphs paint
    // at the bottom of the line box, below the row's clip.
    const table = '<w:tbl><w:tblPr><w:tblW w:w="2000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>'
      + '<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="240" w:hRule="exact"/></w:trPr>'
      + '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:before="0" w:after="0" '
      + 'w:line="800" w:lineRule="exact"/></w:pPr><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>Gone</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';
    const run = runsOf(table, false).find((candidate) => candidate.text === 'Gone')!;
    expect(run.clips).toBeDefined();
    const element = () => ({ style: { cssText: '' }, children: [] as unknown[], appendChild(child: unknown) { this.children.push(child); } });
    vi.stubGlobal('document', { createElement: element });
    try {
      const layer = element();
      buildDocxHighlightLayer(layer as never, [run], [{ active: false, slices: [{ runIndex: 0, start: 0, end: 4 }] }],
        1000, 1000, () => (value: string) => [...value].length * 6);
      const boxes = (node: { children: unknown[] }): number => node.children.reduce<number>((sum, child) =>
        sum + 1 + boxes(child as { children: unknown[] }), 0);
      expect(boxes(layer)).toBe(0);
    } finally {
      vi.unstubAllGlobals();
    }
  });
});

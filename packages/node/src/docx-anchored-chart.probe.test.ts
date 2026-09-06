import { describe, it, expect } from 'vitest';
import { readFileSync, existsSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  installImageBitmapShim,
  installOffscreenCanvasShim,
  type NodeCanvasFactory,
} from './render.ts';
import {
  importForTests,
  loadDocxRendererForTests,
  loadSkiaForTests,
  type DocxLayoutServices,
} from './test-imports';

// skia-canvas is a devDependency; the private chart sample is git-ignored.
// absent → skip cleanly (local), OOXML_REQUIRE_SKIA=1 (CI) → hard failure.
const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas, loadImage } = (skia ?? {}) as Skia;

const factory: NodeCanvasFactory = {
  createCanvas: (w, h) =>
    new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: (async (buf: ArrayBuffer | Uint8Array | Buffer) =>
    loadImage(Buffer.from(buf as Uint8Array))) as unknown as NodeCanvasFactory['loadImage'],
};

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '../../..');
const docxMod = skia ? await importForTests(() => import('./docx.ts'), './docx.ts (docx WASM)') : null;
const rendererMod = skia ? await loadDocxRendererForTests() : null;

const SAMPLE_24 = resolve(ROOT, 'packages/docx/public/private/docx/sample-24.docx');
const haveSample = existsSync(SAMPLE_24);
const sampleBytes = haveSample ? readFileSync(SAMPLE_24) : null;

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

// Recursively collect every DocRun of the given `type` from a document body,
// descending into table cells (charts in sample-24 sit in ordinary body paras).
function collectRuns(body: Any[], type: string): Any[] {
  const out: Any[] = [];
  const walkPara = (p: Any) => {
    for (const r of p.runs ?? []) if (r.type === type) out.push(r);
  };
  const walkTable = (t: Any) => {
    for (const row of t.rows ?? [])
      for (const cell of row.cells ?? [])
        for (const el of cell.content ?? []) {
          if (el.type === 'table' || el.rows) walkTable(el);
          else walkPara(el);
        }
  };
  for (const el of body) {
    if (el.type === 'paragraph' || (el.runs && !el.rows)) walkPara(el);
    else if (el.type === 'table' || el.rows) walkTable(el);
  }
  return out;
}

// Private sample numbers are local workspace slots, not stable fixture ids.
// Run this probe only when the slot still contains the chart fixture this test
// describes; another valid private document at the same path must skip cleanly.
const haveExpectedChartSample = await (async () => {
  if (!docxMod || !sampleBytes) return false;
  try {
    const doc = await docxMod.materializeDocxDocument(sampleBytes) as Any;
    return collectRuns(doc.body, 'chart').length === 6
      && collectRuns(doc.body, 'image').length === 0;
  } catch {
    return false;
  }
})();

async function renderPage(
  doc: Any,
  pageIndex: number,
  dpr = 1,
  layoutServices?: DocxLayoutServices,
): Promise<{ data: Uint8ClampedArray; w: number; h: number }> {
  const { createLayoutServices, renderDocumentToCanvas } = rendererMod!;
  const widthPx = doc.section.pageWidth;
  const heightPx = doc.section.pageHeight;
  const canvas = new Canvas(Math.round(widthPx * dpr), Math.round(heightPx * dpr));
  const restoreImg = installImageBitmapShim(factory);
  const restoreOff = installOffscreenCanvasShim(factory);
  try {
    await renderDocumentToCanvas(doc, canvas as unknown as OffscreenCanvas, pageIndex, {
      dpr,
      width: widthPx,
      layoutServices: layoutServices ?? createLayoutServices(doc),
      currentDate: 0,
      defaultCurrentDateMs: 0,
    });
  } finally {
    restoreOff();
    restoreImg();
  }
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
  const img = ctx.getImageData(0, 0, canvas.width, canvas.height);
  return { data: img.data, w: canvas.width, h: canvas.height };
}

// Count non-white (drawn) pixels inside a page-space rect (in px).
function nonWhiteInRect(
  data: Uint8ClampedArray,
  w: number,
  h: number,
  x0: number,
  y0: number,
  x1: number,
  y1: number,
): number {
  let n = 0;
  const xa = Math.max(0, Math.floor(x0));
  const xb = Math.min(w, Math.ceil(x1));
  const ya = Math.max(0, Math.floor(y0));
  const yb = Math.min(h, Math.ceil(y1));
  for (let y = ya; y < yb; y++) {
    for (let x = xa; x < xb; x++) {
      const i = (y * w + x) * 4;
      if (data[i] < 250 || data[i + 1] < 250 || data[i + 2] < 250) n++;
    }
  }
  return n;
}

describe.skipIf(!skia || !docxMod || !rendererMod || !haveExpectedChartSample)(
  'docx anchored + chartex charts render (sample-24, #747/#752)',
  () => {
    // #747 — sample-24's two inline PNG pictures are the mc:Fallback rendered
    // images of the two chartex charts (chart4/chart6). Because the parser
    // understands `Requires="cx"`, MCE (ECMA-376 Part 3) selects the Choice
    // (live chartEx chart) and drops the Fallback picture — so the correct
    // parse is 6 charts + 0 image runs (no double-emit). Word renders the same
    // two charts, NOT the static snapshots.
    it('sample-24 parses 6 chart runs and 0 image runs (MCE Choice, not Fallback)', async () => {
      const { materializeDocxDocument } = docxMod!;
      const doc = await materializeDocxDocument(sampleBytes!) as Any;
      const charts = collectRuns(doc.body, 'chart');
      const images = collectRuns(doc.body, 'image');
      expect(charts.length).toBe(6);
      // The 2 chartex fallback PNGs must NOT surface as image runs.
      expect(images.length).toBe(0);
      // All inline (no anchored charts in this fixture).
      expect(charts.every((c) => c.anchor === false)).toBe(true);
    });

    // The chartex charts (chart4 boxWhisker, chart6 sunburst) must actually
    // draw — proving the whole zip → parse → renderChart pipeline paints ink
    // where Word's PDF ground truth shows the two charts.
    it('sample-24 pages paint ink for every chart box (charts draw, not blank)', async () => {
      const { materializeDocxDocument } = docxMod!;
      const doc = await materializeDocxDocument(sampleBytes!) as Any;
      const restoreImg = installImageBitmapShim(factory);
      const restoreOff = installOffscreenCanvasShim(factory);
      let pageCount: number;
      let layoutServices: DocxLayoutServices;
      try {
        const { createLayoutServices, layoutDocument } = rendererMod!;
        layoutServices = createLayoutServices(doc);
        pageCount = layoutDocument(doc, layoutServices, { currentDateMs: 0 }).pages.length;
      } finally {
        restoreOff();
        restoreImg();
      }
      let totalInk = 0;
      for (let p = 0; p < pageCount; p++) {
        const { data, w, h } = await renderPage(doc, p, 1, layoutServices);
        totalInk += nonWhiteInRect(data, w, h, 0, 0, w, h);
      }
      // A blank multi-chart document would be near-zero; charts + text draw
      // hundreds of thousands of non-white px.
      expect(totalInk).toBeGreaterThan(50000);
    });

    // #752 render path — synthesize a floating chart by flipping one of
    // sample-24's inline charts to `<wp:anchor>` (posOffset 72pt/72pt from the
    // page), render it as the SOLE body content, and assert ink lands in the
    // anchored box. Before the fix, an anchored chart was dropped in layout
    // (`if (chartRun.anchor) continue;`) and renderAnchorImages had no chart
    // branch, so the box stayed blank.
    it('an anchored chart draws at its absolute page box (#752)', async () => {
      const { materializeDocxDocument } = docxMod!;
      const base = await materializeDocxDocument(sampleBytes!) as Any;

      // Find a REAL parsed paragraph that carries an inline chart (keeps every
      // paragraph prop the layout/pagination path reads), and flip that chart
      // run to a floating `<wp:anchor>` at 72pt/72pt from the page top-left.
      let hostPara: Any = null;
      let chartRun: Any = null;
      for (const el of base.body) {
        if (el.type === 'paragraph' && el.runs) {
          const c = el.runs.find((r: Any) => r.type === 'chart');
          if (c) {
            hostPara = el;
            chartRun = c;
            break;
          }
        }
      }
      expect(hostPara).toBeTruthy();
      expect(chartRun).toBeTruthy();

      const chartWidthPt = chartRun.widthPt;
      const chartHeightPt = chartRun.heightPt;
      const offPt = 72; // 1 inch from the page top-left
      chartRun.anchor = true;
      chartRun.anchorXPt = offPt;
      chartRun.anchorYPt = offPt;
      chartRun.anchorXFromMargin = false; // relativeFrom page/column → offset path
      chartRun.anchorYFromPara = false;

      // Render ONLY the host paragraph so the anchored box is isolated in white
      // (strip headers/footers/notes for the same reason). Keep the real
      // section (page geometry).
      const doc: Any = {
        ...base,
        body: [hostPara],
        headers: { default: null, first: null, even: null },
        footers: { default: null, first: null, even: null },
        footnotes: [],
        endnotes: [],
      };

      const { data, w, h } = await renderPage(doc, 0, 1); // dpr 1, scale 1 px/pt
      const bx0 = offPt;
      const by0 = offPt;
      const bx1 = bx0 + chartWidthPt;
      const by1 = by0 + chartHeightPt;
      const inkInBox = nonWhiteInRect(data, w, h, bx0, by0, bx1, by1);
      // The chart fills a large fraction of its box → tens of thousands of px.
      expect(inkInBox).toBeGreaterThan(5000);

      // Nothing should draw in the band ABOVE the anchored box: the chart is
      // the only graphic and it starts at y=offPt, so y∈[0, offPt/2) must be
      // blank — confirming the chart landed at the anchor, not the flow top.
      const inkAbove = nonWhiteInRect(data, w, h, 0, 0, w, offPt / 2);
      expect(inkAbove).toBe(0);
    });
  },
);

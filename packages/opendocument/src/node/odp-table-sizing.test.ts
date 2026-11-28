// Tables of an `odp` source keep the ODF row-height rule in the Node session
// renderer by the option that claimed the input, not by a model field: a
// provider whose tables omit `odfRowHeights` paints them the same way.
import { describe, expect, it, vi } from 'vitest';
import { buildOdpFixture } from '../odp-fixture.js';
import { testOdpSource } from '../test-sources.js';
import { materializePptxPresentation, openPptxPresentation, renderSlideNode } from './node-facade.js';
import type { NodeCanvasFactory, NodeCanvasLike } from '../../../node/src/index.ts';

// Two 10pt rows in a 20pt frame (no slack) holding 18pt text, so PowerPoint
// growth would enlarge the first row and the ODF rule keeps it.
const deck = () => buildOdpFixture({
  automaticStyles: '<style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="6cm"/></style:style>'
    + '<style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="10pt"/></style:style>'
    + '<style:style style:name="ce1" style:family="table-cell"><style:paragraph-properties fo:border="0.5pt solid #000000"/></style:style>',
  pages: '<draw:page draw:master-page-name="M"><draw:frame svg:x="0cm" svg:y="0cm" svg:width="6cm" svg:height="20pt"><table:table>'
    + '<table:table-column table:style-name="co1"/>'
    + '<table:table-row table:style-name="ro1"><table:table-cell table:style-name="ce1"><text:p>tall</text:p></table:table-cell></table:table-row>'
    + '<table:table-row table:style-name="ro1"><table:table-cell table:style-name="ce1"><text:p>next</text:p></table:table-cell></table:table-row>'
    + '</table:table></draw:frame></draw:page>',
});

type TableModel = { type: string; odfRowHeights?: boolean };
type SlideModel = { elements: TableModel[] };

/** A canvas recording the y of every horizontal stroke, in slide points. */
function strokeCanvas(pxPerPt: number): { canvas: NodeCanvasLike; rows: () => number[] } {
  const ys = new Set<number>();
  let from = { x: 0, y: 0 };
  const state: Record<string, unknown> = {
    font: '10px sans-serif', fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'start', textBaseline: 'alphabetic', globalAlpha: 1, letterSpacing: '0px',
  };
  const canvas = { width: 1, height: 1 } as NodeCanvasLike;
  const ctx = new Proxy(state, {
    get(_target, prop: string) {
      if (prop === 'canvas') return canvas;
      if (prop in state) return state[prop];
      if (prop === 'measureText') {
        return (text: string) => ({
          width: text.length * 7, actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
          fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
        });
      }
      if (prop === 'moveTo') return (x: number, y: number) => { from = { x, y }; };
      if (prop === 'lineTo') return (x: number, y: number) => {
        if (Math.abs(y - from.y) < 0.01 && Math.abs(x - from.x) > 1) ys.add(Math.round(y / pxPerPt));
        from = { x, y };
      };
      if (prop === 'getTransform') return () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
      if (prop === 'createLinearGradient' || prop === 'createRadialGradient') return () => ({ addColorStop() {} });
      return () => undefined;
    },
    set(_target, prop: string, value) { state[prop] = value; return true; },
  });
  (canvas as { getContext: unknown }).getContext = () => ctx;
  return { canvas, rows: () => [...ys].sort((a, b) => a - b) };
}

function fakeFactory(): NodeCanvasFactory {
  const fake = (): NodeCanvasLike => ({ width: 1, height: 1, getContext: vi.fn() as unknown as NodeCanvasLike['getContext'] });
  return { createCanvas: vi.fn(() => fake()), loadImage: vi.fn(async () => ({ width: 1, height: 1 })) };
}

/** Drop the parser's marker, as an application provider might. */
function unmark(slide: SlideModel): void {
  for (const element of slide.elements) if (element.type === 'table') delete element.odfRowHeights;
}

const WIDTH = 960;
// 28cm slide = 793.7pt.
const PX_PER_PT = WIDTH / (28 / 2.54 * 72);

describe('ODP table sizing in the Node session renderer', () => {
  it('keeps the ODF rule for an unmarked table from the odp option', async () => {
    const session = await openPptxPresentation(deck(), { odp: testOdpSource() });
    try {
      for await (const slide of session) {
        unmark(slide as unknown as SlideModel);
        const { canvas, rows } = strokeCanvas(PX_PER_PT);
        await session.renderSlide(canvas, slide, { width: WIDTH, dpr: 1, factory: fakeFactory() });
        expect(rows()).toEqual([0, 10, 20]);
      }
    } finally {
      await session.close();
    }
  });

  it('keeps it for a materialized slide too', async () => {
    const presentation = await materializePptxPresentation(deck(), { odp: testOdpSource() });
    unmark(presentation.slides[0] as unknown as SlideModel);
    const { canvas, rows } = strokeCanvas(PX_PER_PT);
    await renderSlideNode(canvas, presentation, 0, { width: WIDTH, dpr: 1, factory: fakeFactory() });
    expect(rows()).toEqual([0, 10, 20]);
  });

  it('follows the claiming option when one provider fills both', async () => {
    // legacyPpt claims first: its unmarked table takes PowerPoint growth.
    const source = testOdpSource();
    const presentation = await materializePptxPresentation(deck(), { legacyPpt: source as never, odp: source });
    unmark(presentation.slides[0] as unknown as SlideModel);
    const { canvas, rows } = strokeCanvas(PX_PER_PT);
    await renderSlideNode(canvas, presentation, 0, { width: WIDTH, dpr: 1, factory: fakeFactory() });
    expect(rows()[1]).toBeGreaterThan(10);
  });
});

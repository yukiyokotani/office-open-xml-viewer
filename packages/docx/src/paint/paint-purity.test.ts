import { describe, expect, it } from 'vitest';
import type { SectionLayoutContext } from '../layout-context.js';
import type { DocumentLayout } from '../layout/types.js';
import { paintLayoutPage } from './canvas-page.js';

describe('paintLayoutPage', () => {
  it('paints retained geometry without measuring text', async () => {
    const fills: Array<{ fill: string; args: number[] }> = [];
    let currentFill = '';
    const context = {
      get fillStyle() { return currentFill; },
      set fillStyle(value: string | CanvasGradient | CanvasPattern) { currentFill = String(value); },
      save() {},
      restore() {},
      setTransform() {},
      transform() {},
      clearRect() {},
      fillRect(...args: number[]) {
        fills.push({ fill: currentFill, args });
      },
      measureText() {
        throw new Error('paint must not measure text');
      },
    } as unknown as CanvasRenderingContext2D;
    const target = {
      width: 0,
      height: 0,
      getContext: () => context,
    } as unknown as HTMLCanvasElement;
    const layout: DocumentLayout = {
      pages: [{
        pageIndex: 0,
        geometry: {
          xPt: 0,
          yPt: 0,
          widthPt: 100,
          heightPt: 200,
          contentTopPt: 10,
          contentBottomPt: 190,
        },
        flowDomains: [{
          id: 'body',
          kind: 'body',
          bounds: { xPt: 10, yPt: 10, widthPt: 80, heightPt: 180 },
        }],
        section: {} as SectionLayoutContext,
        sectionOccurrenceId: 'section:0', parityBlank: false, bookmarkStarts: [],
        pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:0' },
        sectionRegions: [{
          id: 'region:0', sectionOccurrenceId: 'section:0',
          coordinateSpace: { writingMode: 'horizontal-tb', logicalToPhysical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 } },
          blockStartPt: 10, blockEndPt: 190, flowDomainIds: ['body'],
          section: {} as SectionLayoutContext,
        }],
        layers: {
          paintOrder: [{ layer: 'body', nodeId: 'drawing-1', coordinateSpace: 'logical-body-points', logicalBlock: { blockStartPt: 20, blockExtentPt: 40 } }],
          background: [],
          behindText: [],
          header: [],
          body: [{
            kind: 'drawing',
            id: 'drawing-1',
            source: { story: 'body', storyInstance: 'body', path: [0] },
            flowBounds: { xPt: 10, yPt: 20, widthPt: 30, heightPt: 40 },
            inkBounds: { xPt: 10, yPt: 20, widthPt: 30, heightPt: 40 },
            advancePt: 40,
            ordinaryFlow: true,
            flowDomainId: 'body',
            commands: [{
              kind: 'fill-rect',
              rect: { xPt: 10, yPt: 20, widthPt: 30, heightPt: 40 },
              fill: '#ff0000',
            }],
          }],
          notes: [],
          front: [],
          footer: [],
        },
        readingOrder: ['drawing-1'],
      }],
      diagnostics: [],
    };

    await expect(paintLayoutPage(layout, 0, target, { scale: 1, dpr: 1 })).resolves.toBeUndefined();
    expect(fills).toEqual([{ fill: '#ff0000', args: [10, 20, 30, 40] }]);
  });

  it('rejects missing and duplicate paint references instead of dropping content', async () => {
    const context = {
      save() {}, restore() {}, setTransform() {}, transform() {}, clearRect() {}, fillRect() {},
      fillStyle: '',
    } as unknown as CanvasRenderingContext2D;
    const target = { width: 0, height: 0, getContext: () => context } as unknown as HTMLCanvasElement;
    const node = {
      kind: 'drawing' as const,
      id: 'drawing-1',
      source: { story: 'body' as const, storyInstance: 'body', path: [0] },
      flowBounds: { xPt: 10, yPt: 20, widthPt: 30, heightPt: 40 },
      inkBounds: { xPt: 10, yPt: 20, widthPt: 30, heightPt: 40 },
      advancePt: 40,
      ordinaryFlow: true,
      flowDomainId: 'body',
      commands: [],
    };
    const page = {
      pageIndex: 0,
      geometry: { xPt: 0, yPt: 0, widthPt: 100, heightPt: 200, contentTopPt: 10, contentBottomPt: 190 },
      flowDomains: [{ id: 'body', kind: 'body' as const, bounds: { xPt: 10, yPt: 10, widthPt: 80, heightPt: 180 } }],
      section: {} as SectionLayoutContext,
      sectionOccurrenceId: 'section:0', parityBlank: false, bookmarkStarts: [],
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:0' },
      sectionRegions: [{
        id: 'region:0', sectionOccurrenceId: 'section:0',
        coordinateSpace: { writingMode: 'horizontal-tb' as const, logicalToPhysical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 } },
        blockStartPt: 10, blockEndPt: 190, flowDomainIds: ['body'],
        section: {} as SectionLayoutContext,
      }],
      layers: {
        paintOrder: [{ layer: 'body' as const, nodeId: 'missing', coordinateSpace: 'logical-body-points' as const, logicalBlock: { blockStartPt: 20, blockExtentPt: 40 } }],
        background: [], behindText: [], header: [], body: [node], notes: [], front: [], footer: [],
      },
      readingOrder: [node.id],
    };
    const missing: DocumentLayout = { pages: [page], diagnostics: [] };
    await expect(paintLayoutPage(missing, 0, target, { scale: 1, dpr: 1 })).rejects.toThrow(/missing/i);

    const duplicate: DocumentLayout = {
      pages: [{ ...page, layers: { ...page.layers, body: [node, node], paintOrder: [
        { layer: 'body', nodeId: node.id, coordinateSpace: 'logical-body-points', logicalBlock: { blockStartPt: 20, blockExtentPt: 40 } },
        { layer: 'body', nodeId: node.id, coordinateSpace: 'logical-body-points', logicalBlock: { blockStartPt: 20, blockExtentPt: 40 } },
      ] } }],
      diagnostics: [],
    };
    await expect(paintLayoutPage(duplicate, 0, target, { scale: 1, dpr: 1 })).rejects.toThrow(/duplicate/i);
  });

  it('dispatches retained tables through the canonical page painter', async () => {
    const fills: unknown[] = [];
    let currentFill = '';
    const context = {
      get fillStyle() { return currentFill; },
      set fillStyle(value: string | CanvasGradient | CanvasPattern) { currentFill = String(value); },
      save() {}, restore() {}, setTransform() {}, transform() {}, clearRect() {},
      beginPath() {}, rect() {}, clip() {}, translate() {}, rotate() {}, scale() {},
      fillRect(x: number, y: number, width: number, height: number) {
        fills.push([x, y, width, height, currentFill]);
      },
      strokeRect() {}, setLineDash() {}, moveTo() {}, lineTo() {}, stroke() {}, fill() {},
      strokeStyle: '', lineWidth: 1,
    } as unknown as CanvasRenderingContext2D;
    const target = { width: 0, height: 0, getContext: () => context } as unknown as HTMLCanvasElement;
    const bounds = { xPt: 10, yPt: 20, widthPt: 80, heightPt: 16 };
    const cell = {
      kind: 'table-cell', id: 'cell-0',
      source: { story: 'body', storyInstance: 'body', path: [0, 0, 0] },
      flowDomainId: 'body', ordinaryFlow: true,
      flowBounds: bounds, inkBounds: bounds, contentBounds: bounds,
      advancePt: 16, verticalMerge: 'none', vAlign: 'top',
      background: { color: '#abcdef' }, blocks: [],
    };
    const row = {
      kind: 'table-row', id: 'row-0',
      source: { story: 'body', storyInstance: 'body', path: [0, 0] },
      flowDomainId: 'body', ordinaryFlow: true,
      flowBounds: bounds, inkBounds: bounds, advancePt: 16, cells: [cell],
    };
    const table = {
      kind: 'table', id: 'table-0',
      source: { story: 'body', storyInstance: 'body', path: [0] },
      flowDomainId: 'body', ordinaryFlow: true,
      flowBounds: bounds, inkBounds: bounds, advancePt: 16,
      columnWidthsPt: [80], rows: [row], borders: [],
    };
    const layout = {
      pages: [{
        pageIndex: 0,
        geometry: { xPt: 0, yPt: 0, widthPt: 100, heightPt: 200, contentTopPt: 10, contentBottomPt: 190 },
        flowDomains: [{ id: 'body', kind: 'body', bounds: { xPt: 10, yPt: 10, widthPt: 80, heightPt: 180 } }],
        section: {} as SectionLayoutContext,
        sectionOccurrenceId: 'section:0', parityBlank: false, bookmarkStarts: [],
        pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:0' },
        sectionRegions: [{
          id: 'region:0', sectionOccurrenceId: 'section:0',
          coordinateSpace: { writingMode: 'horizontal-tb', logicalToPhysical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 } },
          blockStartPt: 10, blockEndPt: 190, flowDomainIds: ['body'],
          section: {} as SectionLayoutContext,
        }],
        layers: {
          paintOrder: [{ layer: 'body', nodeId: 'table-0', coordinateSpace: 'logical-body-points', logicalBlock: { blockStartPt: 20, blockExtentPt: 16 } }],
          background: [], behindText: [], header: [], body: [table], notes: [], front: [], footer: [],
        },
        readingOrder: ['table-0'],
      }],
      diagnostics: [],
    } as unknown as DocumentLayout;

    await paintLayoutPage(layout, 0, target, { scale: 1, dpr: 1 });

    expect(fills).toContainEqual([10, 20, 80, 16, '#abcdef']);
  });
});

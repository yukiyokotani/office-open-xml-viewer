import { describe, expect, it } from 'vitest';
import { createCanvasFontRoute } from '@silurus/ooxml-core';
import {
  canonicalLogicalToPhysical,
  composeAffine,
  mapAffinePoint,
  mapAffineRect,
} from '../layout/affine.js';
import { assertDocumentLayout } from '../layout/invariants.js';
import type {
  DocumentLayout,
  DrawingLayout,
  LayoutRect,
  Matrix2DData,
  PageOccurrenceCoordinateSpace,
  PagePaintNode,
  ParagraphLayout,
  PaintReadyTableLayout,
  ResolvedFloatingTablePlacementLayout,
  TableCellLayout,
  TableLayout,
  TableRowLayout,
  TextBoxLayout,
  WritingMode,
} from '../layout/types.js';
import type { SectionLayoutContext } from '../layout-context.js';
import { paintLayoutPage } from './canvas-page.js';

const rect = (xPt: number, yPt: number, widthPt: number, heightPt: number): LayoutRect => ({
  xPt, yPt, widthPt, heightPt,
});

const fontRoute = createCanvasFontRoute('Test Sans', 'native');

function paragraph(
  id: string,
  domain: string,
  bounds = rect(11, 17, 23, 7),
  options: { hyperlink?: string; tateChuYoko?: boolean; drawings?: readonly DrawingLayout[] } = {},
): ParagraphLayout {
  return {
    kind: 'paragraph', id,
    source: { story: 'body', storyInstance: 'body', path: [0] },
    flowDomainId: domain, ordinaryFlow: true,
    flowBounds: bounds, inkBounds: bounds, advancePt: bounds.heightPt,
    spacing: { beforePt: 0, afterPt: 0 }, contextualSpacing: false,
    lines: [{
      range: { start: 0, end: id.length }, bounds, baselinePt: bounds.yPt + 5, advancePt: bounds.heightPt,
      placements: [{
        kind: 'text', text: id, range: { start: 0, end: id.length },
        origin: { xPt: bounds.xPt, yPt: bounds.yPt + 5 }, bounds,
        advancePt: bounds.widthPt,
        clusters: [{ range: { start: 0, end: id.length }, offset: { xPt: 0, yPt: 0 }, advancePt: bounds.widthPt }],
        paintOps: [{
          text: id, range: { start: 0, end: id.length }, offset: { xPt: 0, yPt: 0 },
          letterSpacingPt: 0, scaleX: 1, direction: 'ltr', kerning: 'auto',
          writingMode: 'horizontal-tb',
        }],
        color: { kind: 'explicit', color: '#111111' }, fontRoute, fontSizePt: 10,
        fontWeight: 400, fontStyle: 'normal', direction: 'ltr', decorations: [],
        ...(options.hyperlink ? { hyperlink: options.hyperlink } : {}),
        ...(options.tateChuYoko ? { tateChuYoko: true } : {}),
      }],
    }],
    borders: [], resources: [], drawings: options.drawings ?? [], textBoxes: [], events: [], exclusions: [],
  };
}

function table(
  id: string,
  domain: string,
  bounds: LayoutRect,
  child?: ParagraphLayout,
  background = '#abcdef',
): PaintReadyTableLayout {
  const cell: TableCellLayout = {
    kind: 'table-cell', id: `${id}:cell`,
    source: { story: 'body', storyInstance: 'body', path: [0, 0, 0] },
    flowDomainId: domain, ordinaryFlow: true,
    flowBounds: bounds, inkBounds: bounds, contentBounds: bounds,
    advancePt: bounds.heightPt, verticalMerge: 'none', vAlign: 'top',
    background: { color: background },
    blocks: child ? [{ layout: child, offsetPt: 0, advancePt: child.advancePt }] : [],
  };
  const row: TableRowLayout = {
    kind: 'table-row', id: `${id}:row`,
    source: { story: 'body', storyInstance: 'body', path: [0, 0] },
    flowDomainId: domain, ordinaryFlow: true,
    flowBounds: bounds, inkBounds: bounds, advancePt: bounds.heightPt,
    heightPt: bounds.heightPt, contentHeightPt: bounds.heightPt, cells: [cell],
  };
  return {
    kind: 'table', id,
    source: { story: 'body', storyInstance: 'body', path: [0] },
    flowDomainId: domain, ordinaryFlow: true,
    flowBounds: bounds, inkBounds: bounds, advancePt: bounds.heightPt,
    columnWidthsPt: [bounds.widthPt], rows: [row], borders: [],
    paintReadyFloatingTables: { kind: 'none' },
  };
}

type RegionInput = Readonly<{
  id: string;
  domain: string;
  mode: WritingMode;
  node: PagePaintNode;
  coordinateSpace?: PageOccurrenceCoordinateSpace;
  logicalBlockExtentPt?: number;
}>;

function documentFor(regions: readonly RegionInput[], pageWidthPt = 333): DocumentLayout {
  const geometry = { ...rect(0, 0, pageWidthPt, 517), contentTopPt: 0, contentBottomPt: 517 };
  return {
    pages: [{
      pageIndex: 0, geometry,
      section: {} as SectionLayoutContext,
      sectionOccurrenceId: 'section:0', parityBlank: false, bookmarkStarts: [],
      pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:0' },
      sectionRegions: regions.map(({ id, domain, mode }) => ({
        id, sectionOccurrenceId: id,
        coordinateSpace: {
          writingMode: mode,
          logicalToPhysical: canonicalLogicalToPhysical(mode, pageWidthPt),
        },
        blockStartPt: 0, blockEndPt: 517, flowDomainIds: [domain],
        section: {} as SectionLayoutContext,
      })),
      flowDomains: regions.map(({ domain, mode, node }) => ({
        id: domain, kind: 'body',
        bounds: (regions.find((item) => item.domain === domain)?.coordinateSpace ?? 'logical-body-points')
          === 'logical-body-points'
          ? mapAffineRect(canonicalLogicalToPhysical(mode, pageWidthPt), node.flowBounds)
          : node.flowBounds,
      })),
      layers: {
        paintOrder: regions.map(({ node, coordinateSpace = 'logical-body-points', logicalBlockExtentPt }) => ({
          layer: 'body', nodeId: node.id, coordinateSpace,
          logicalBlock: {
            blockStartPt: node.flowBounds.yPt,
            blockExtentPt: logicalBlockExtentPt ?? node.flowBounds.heightPt,
          },
        })),
        background: [], behindText: [], header: [], body: regions.map(({ node }) => node),
        notes: [], front: [], footer: [],
      },
      readingOrder: regions.map(({ node }) => node.id),
    }],
    diagnostics: [],
  };
}

interface RecordedRect extends LayoutRect { readonly fill: string }

function canvasTarget() {
  let ctm: Matrix2DData = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
  const stack: Matrix2DData[] = [];
  const texts: Array<{ text: string; point: { xPt: number; yPt: number } }> = [];
  const fills: RecordedRect[] = [];
  const clips: LayoutRect[] = [];
  const operations: string[] = [];
  let clears = 0;
  const context = {
    globalAlpha: 1, fillStyle: '', strokeStyle: '', lineWidth: 1, font: '',
    textAlign: 'left', textBaseline: 'alphabetic', direction: 'ltr',
    letterSpacing: '0px', fontKerning: 'auto',
    save() { stack.push({ ...ctm }); },
    restore() { ctm = stack.pop()!; },
    setTransform(a: number, b: number, c: number, d: number, e: number, f: number) {
      ctm = { a, b, c, d, e, f };
    },
    transform(a: number, b: number, c: number, d: number, e: number, f: number) {
      ctm = composeAffine(ctm, { a, b, c, d, e, f });
    },
    translate(x: number, y: number) { this.transform(1, 0, 0, 1, x, y); },
    scale(x: number, y: number) { this.transform(x, 0, 0, y, 0, 0); },
    rotate(angle: number) { this.transform(Math.cos(angle), Math.sin(angle), -Math.sin(angle), Math.cos(angle), 0, 0); },
    clearRect() { clears += 1; },
    fillText(text: string, x: number, y: number) { texts.push({ text, point: mapAffinePoint(ctm, { xPt: x, yPt: y }) }); },
    fillRect(x: number, y: number, width: number, height: number) {
      fills.push({ ...mapAffineRect(ctm, rect(x, y, width, height)), fill: String(this.fillStyle) });
      operations.push(`fill:${String(this.fillStyle)}`);
    },
    beginPath() {},
    rect(x: number, y: number, width: number, height: number) { clips.push(mapAffineRect(ctm, rect(x, y, width, height))); },
    clip() {}, strokeRect() {}, setLineDash() {}, moveTo() {}, lineTo() {},
    stroke() { operations.push(`stroke:${String(this.strokeStyle)}`); },
    fill() {}, drawImage() {},
  };
  const proxy = new Proxy(context, {
    get(target, property, receiver) {
      if (property === 'measureText') throw new Error('canonical paint must not measure');
      return Reflect.get(target, property, receiver);
    },
  });
  const target = {
    width: 0, height: 0,
    getContext: () => proxy,
  } as unknown as HTMLCanvasElement;
  return { target, texts, fills, clips, operations, clearCount: () => clears };
}

describe('canonical page coordinate paint', () => {
  it.each<WritingMode>(['horizontal-tb', 'vertical-rl', 'vertical-lr'])(
    'maps one real paragraph exactly once in %s',
    async (mode) => {
      const node = paragraph(mode, 'body');
      const layout = documentFor([{ id: 'region', domain: 'body', mode, node }]);
      const canvas = canvasTarget();
      await paintLayoutPage(layout, 0, canvas.target, { scale: 1.25, dpr: 2 });

      const expected = mapAffinePoint(
        composeAffine(
          { a: 2.5, b: 0, c: 0, d: 2.5, e: 0, f: 0 },
          canonicalLogicalToPhysical(mode, 333),
        ),
        { xPt: 11, yPt: 22 },
      );
      expect(canvas.texts).toEqual([{ text: mode, point: expected }]);
      expect(canvas.target.width).toBe(Math.ceil(333 * 2.5));
      expect(canvas.target.height).toBe(Math.ceil(517 * 2.5));
      expect(canvas.clearCount()).toBe(1);
    },
  );

  it('uses each continuous region transform independently on one page', async () => {
    const horizontal = paragraph('horizontal', 'h', rect(5, 7, 12, 6));
    const vertical = paragraph('vertical', 'v', rect(31, 41, 12, 6));
    const layout = documentFor([
      { id: 'h-region', domain: 'h', mode: 'horizontal-tb', node: horizontal },
      { id: 'v-region', domain: 'v', mode: 'vertical-rl', node: vertical },
    ]);
    expect(() => assertDocumentLayout(layout)).not.toThrow();
    const canvas = canvasTarget();
    await paintLayoutPage(layout, 0, canvas.target, { scale: 1, dpr: 1 });

    expect(canvas.texts.map(({ text, point }) => [text, point])).toEqual([
      ['horizontal', { xPt: 5, yPt: 12 }],
      ['vertical', { xPt: 287, yPt: 31 }],
    ]);
  });

  it('reports vertical run geometry in physical CSS coordinates without DPR', async () => {
    const node = paragraph('run', 'body', rect(11, 17, 23, 7), {
      hyperlink: 'https://example.test/', tateChuYoko: true,
    });
    const layout = documentFor([{ id: 'region', domain: 'body', mode: 'vertical-rl', node }]);
    const runs: unknown[] = [];
    await paintLayoutPage(layout, 0, canvasTarget().target, {
      scale: 2, dpr: 3, onTextRun: (run) => runs.push(run),
    });

    expect(runs).toEqual([expect.objectContaining({
      text: 'run', x: 632, y: 22, w: 46, h: 14, fontSize: 20,
      transform: 'rotate(90deg)',
      hyperlink: { kind: 'external', url: 'https://example.test/' },
      eastAsianVert: true,
    })]);
  });

  it('bypasses the region transform for an upright physical table and keeps its logical block charge', async () => {
    const child = paragraph('cell', 'body', rect(202, 42, 20, 8));
    const node = table('upright', 'body', rect(200, 40, 80, 30), child);
    const layout = documentFor([{
      id: 'region', domain: 'body', mode: 'vertical-rl', node,
      coordinateSpace: 'upright-physical-page-points', logicalBlockExtentPt: 80,
    }]);
    expect(() => assertDocumentLayout(layout)).not.toThrow();
    const canvas = canvasTarget();
    await paintLayoutPage(layout, 0, canvas.target, { scale: 2, dpr: 1.5 });

    expect(canvas.fills).toContainEqual({ ...rect(600, 120, 240, 90), fill: '#abcdef' });
    expect(canvas.texts).toContainEqual({ text: 'cell', point: { xPt: 600, yPt: 135 } });
    expect(layout.pages[0]!.layers.paintOrder[0]!.logicalBlock).toEqual({
      blockStartPt: 40, blockExtentPt: 80,
    });
  });

  it('keeps distinct-margin exact clips and auto growth on canonical physical Y', async () => {
    const pageWidthPt = 333;
    const blockCursorPt = 73;
    const tableWidthPt = 80;
    const physicalLeftPt = pageWidthPt - blockCursorPt - tableWidthPt;
    const physicalTopPt = 37;
    const exactBounds = rect(physicalLeftPt, physicalTopPt, tableWidthPt, 80);
    const autoBounds = rect(physicalLeftPt, physicalTopPt + 80, tableWidthPt, 110);
    const exactChild = paragraph('exact-overflow', 'body', rect(0, 0, 30, 10));
    const autoChild = paragraph('auto-growth', 'body', rect(0, 0, 30, 10));
    const cell = (
      id: string,
      bounds: LayoutRect,
      child: ParagraphLayout,
      offsetPt: number,
      clipBounds?: LayoutRect,
    ): TableCellLayout => ({
      kind: 'table-cell', id, source: child.source, flowDomainId: 'body', ordinaryFlow: true,
      flowBounds: bounds, inkBounds: bounds, contentBounds: bounds, advancePt: bounds.heightPt,
      verticalMerge: 'none', vAlign: 'top', background: { color: id === 'exact-cell' ? '#aa1111' : '#11aa11' },
      blocks: [{ layout: child, offsetPt, advancePt: child.advancePt }],
      ...(clipBounds ? { clipBounds } : {}),
    });
    const row = (id: string, bounds: LayoutRect, cells: readonly TableCellLayout[]): TableRowLayout => ({
      kind: 'table-row', id, source: cells[0]!.source, flowDomainId: 'body', ordinaryFlow: true,
      flowBounds: bounds, inkBounds: bounds, advancePt: bounds.heightPt,
      heightPt: bounds.heightPt, contentHeightPt: bounds.heightPt, cells,
    });
    const node: PaintReadyTableLayout = {
      kind: 'table', id: 'exact-auto-upright', source: exactChild.source,
      flowDomainId: 'body', ordinaryFlow: true,
      flowBounds: rect(physicalLeftPt, physicalTopPt, tableWidthPt, 190),
      inkBounds: rect(physicalLeftPt, physicalTopPt, tableWidthPt, 190),
      advancePt: 190, columnWidthsPt: [tableWidthPt], borders: [],
      rows: [
        row('exact-row', exactBounds, [cell('exact-cell', exactBounds, exactChild, 90, exactBounds)]),
        row('auto-row', autoBounds, [cell('auto-cell', autoBounds, autoChild, 90)]),
      ],
      paintReadyFloatingTables: { kind: 'none' },
    };
    const layout = documentFor([{
      id: 'region', domain: 'body', mode: 'vertical-rl', node,
      coordinateSpace: 'upright-physical-page-points', logicalBlockExtentPt: tableWidthPt,
    }], pageWidthPt);
    const canvas = canvasTarget();
    await paintLayoutPage(layout, 0, canvas.target, { scale: 1, dpr: 1 });

    expect(physicalLeftPt).toBe(180);
    expect(canvas.clips).toContainEqual(exactBounds);
    expect(canvas.fills).toContainEqual({ ...exactBounds, fill: '#aa1111' });
    expect(canvas.fills).toContainEqual({ ...autoBounds, fill: '#11aa11' });
    expect(canvas.texts.find(({ text }) => text === 'auto-growth')?.point.yPt)
      .toBeGreaterThan(autoBounds.yPt + 80);
    expect(node.advancePt).toBe(190);
    expect(layout.pages[0]!.layers.paintOrder[0]!.logicalBlock?.blockExtentPt).toBe(tableWidthPt);
  });

  it('keeps a page-owned nested drawing physical beneath a logical vertical host', async () => {
    const physicalRect = rect(250, 30, 9, 5);
    const drawing: DrawingLayout = {
      kind: 'drawing', id: 'page-anchor', source: { story: 'body', storyInstance: 'body', path: [0, 1] },
      flowDomainId: 'body', ordinaryFlow: false, flowBounds: physicalRect, inkBounds: physicalRect,
      advancePt: 0, commands: [{ kind: 'fill-rect', rect: physicalRect, fill: '#fedcba' }],
      anchorLayer: {
        occurrenceId: 'anchor:0', behindDoc: true, relativeHeight: 1, sourceOrder: 0,
        coordinateSpace: 'physical-page-points',
        horizontalOwnership: 'page', verticalOwnership: 'page',
      },
    };
    const host = paragraph('host', 'body', rect(11, 17, 23, 7), { drawings: [drawing] });
    const layout = documentFor([{ id: 'region', domain: 'body', mode: 'vertical-rl', node: host }]);
    const canvas = canvasTarget();
    await paintLayoutPage(layout, 0, canvas.target, { scale: 2, dpr: 2 });

    expect(canvas.fills).toContainEqual({ ...rect(1000, 120, 36, 20), fill: '#fedcba' });
    expect(canvas.texts).toContainEqual({ text: 'host', point: { xPt: 1244, yPt: 44 } });
  });

  it('keeps partial anchor relocation ownership on one non-singular physical frame', async () => {
    const bounds = rect(250, 30, 9, 5);
    const anchored = (
      id: string,
      fill: string,
      horizontalOwnership: 'page' | 'host',
      verticalOwnership: 'page' | 'host',
    ): DrawingLayout => ({
      kind: 'drawing', id, source: { story: 'body', storyInstance: 'body', path: [0, 1] },
      flowDomainId: 'body', ordinaryFlow: false, flowBounds: bounds, inkBounds: bounds,
      advancePt: 0, commands: [{ kind: 'fill-rect', rect: bounds, fill }],
      anchorLayer: {
        occurrenceId: `anchor:${id}`, behindDoc: true, relativeHeight: 1, sourceOrder: 0,
        coordinateSpace: 'physical-page-points',
        horizontalOwnership, verticalOwnership,
      },
    });
    const host = paragraph('partial', 'body', rect(11, 17, 23, 7), {
      drawings: [
        anchored('horizontal-page', '#110000', 'page', 'host'),
        anchored('vertical-page', '#001100', 'host', 'page'),
      ],
    });
    const layout = documentFor([{ id: 'region', domain: 'body', mode: 'vertical-rl', node: host }]);
    const canvas = canvasTarget();
    await paintLayoutPage(layout, 0, canvas.target, { scale: 1, dpr: 1 });

    expect(canvas.fills).toContainEqual({ ...bounds, fill: '#110000' });
    expect(canvas.fills).toContainEqual({ ...bounds, fill: '#001100' });
    const partialFills = canvas.fills.filter((fill) => fill.fill === '#110000' || fill.fill === '#001100');
    expect(partialFills.every((fill) => fill.widthPt > 0 && fill.heightPt > 0)).toBe(true);
  });

  it('re-enters the physical page below translated table and vertical text-box frames', async () => {
    const physical = rect(250, 30, 9, 5);
    const pageDrawing = (id: string, fill: string): DrawingLayout => ({
      kind: 'drawing', id, source: { story: 'body', storyInstance: 'body', path: [0, 1] },
      flowDomainId: 'body', ordinaryFlow: false, flowBounds: physical, inkBounds: physical,
      advancePt: 0, commands: [{ kind: 'fill-rect', rect: physical, fill }],
      anchorLayer: {
        occurrenceId: `anchor:${id}`, behindDoc: true, relativeHeight: 1, sourceOrder: 0,
        coordinateSpace: 'physical-page-points',
        horizontalOwnership: 'page', verticalOwnership: 'page',
      },
    });
    const tableChild = paragraph('table-child', 'body', rect(5, 7, 20, 8), {
      drawings: [pageDrawing('table-page-anchor', '#aa0000')],
    });
    const tableNode = table('translated-table', 'body', rect(50, 70, 80, 30), tableChild);
    const textBoxChild = paragraph('textbox-child', 'body', rect(0, 0, 20, 8), {
      drawings: [pageDrawing('textbox-page-anchor', '#00aa00')],
    });
    const textBox: TextBoxLayout = {
      kind: 'textbox', id: 'vertical-textbox', source: textBoxChild.source,
      flowDomainId: 'textbox:0', ordinaryFlow: false,
      flowBounds: rect(70, 90, 40, 20), inkBounds: rect(70, 90, 40, 20), advancePt: 0,
      paragraphs: [textBoxChild], writingMode: 'vertical-rl', verticalMode: 'vert',
      insets: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
    };
    const textBoxHost = {
      ...paragraph('textbox-host', 'body', rect(20, 40, 30, 10)),
      textBoxes: [textBox],
    };
    const layout = documentFor([
      { id: 'table-region', domain: 'body', mode: 'vertical-rl', node: tableNode },
      { id: 'textbox-region', domain: 'body-2', mode: 'vertical-rl', node: { ...textBoxHost, flowDomainId: 'body-2' } },
    ]);
    const canvas = canvasTarget();
    await paintLayoutPage(layout, 0, canvas.target, { scale: 1, dpr: 1 });

    expect(canvas.fills).toContainEqual({ ...physical, fill: '#aa0000' });
    expect(canvas.fills).toContainEqual({ ...physical, fill: '#00aa00' });
  });

  it('paints relocated upright floats in destination source order before the parent border', async () => {
    const domain = 'page:1:region:destination:column:1';
    const outer = table('outer', domain, rect(180, 40, 100, 40));
    const nested = table('nested', domain, rect(220, 48, 25, 12), undefined, '#123456');
    const second = table('nested-2', domain, rect(245, 64, 25, 12), undefined, '#654321');
    const placement = (
      occurrenceId: string,
      child: PaintReadyTableLayout,
      exclusionBounds: LayoutRect,
      sourceBlockIndex: number,
    ): ResolvedFloatingTablePlacementLayout => {
      const source = {
        kind: 'floating-table-placement' as const, occurrenceId, ownership: 'source' as const,
        physicalPageIndex: 1, displayPageNumber: 7, hostCellId: 'outer:cell',
        sourceBlockIndex, anchorBlockIndex: sourceBlockIndex, tableId: child.id, overlap: 'never' as const,
        positioning: {} as never, anchorBounds: child.flowBounds, child,
      };
      return {
        kind: 'resolved-floating-table-placement', occurrenceId,
        xPt: child.flowBounds.xPt, yPt: child.flowBounds.yPt,
        bounds: child.flowBounds, exclusionBounds, overlap: 'never', child, source,
      };
    };
    const resolved = placement('float:0', nested, rect(218, 46, 29, 16), 0);
    const resolvedSecond = placement('float:1', second, rect(243, 62, 29, 16), 1);
    const fragment = {
      ...outer,
      paintReadyFloatingTables: {
        kind: 'resolved' as const,
        coordinateSpace: 'upright-physical-page-points' as const,
        unresolved: [], placements: [resolved, resolvedSecond],
      },
      borders: [{
        edge: 'right' as const, from: { xPt: 280, yPt: 40 }, to: { xPt: 280, yPt: 80 },
        color: '#000000', widthPt: 1, authoredStyle: 'single', style: 'solid' as const,
      }],
    };
    const onePage = documentFor([{
      id: 'destination', domain, mode: 'vertical-rl', node: fragment,
      coordinateSpace: 'upright-physical-page-points', logicalBlockExtentPt: 100,
    }]);
    const destination = {
      ...onePage.pages[0]!, pageIndex: 1,
      pageNumber: { displayNumber: 7, format: 'decimal', sectionOccurrenceId: 'section:0' },
    };
    const blank = {
      ...destination, pageIndex: 0, parityBlank: true,
      pageNumber: { displayNumber: 6, format: 'decimal', sectionOccurrenceId: 'section:0' },
      flowDomains: [], sectionRegions: [],
      layers: {
        paintOrder: [], background: [], behindText: [], header: [], body: [],
        notes: [], front: [], footer: [],
      },
      readingOrder: [],
    };
    const layout: DocumentLayout = { pages: [blank, destination], diagnostics: [] };
    expect(() => assertDocumentLayout(layout)).not.toThrow();
    const canvas = canvasTarget();
    await expect(paintLayoutPage(layout, 1, canvas.target, { scale: 1, dpr: 1 })).resolves.toBeUndefined();

    expect(canvas.fills).toContainEqual({ ...rect(220, 48, 25, 12), fill: '#123456' });
    expect(canvas.fills).toContainEqual({ ...rect(245, 64, 25, 12), fill: '#654321' });
    expect(canvas.operations.indexOf('fill:#123456')).toBeLessThan(
      canvas.operations.indexOf('fill:#654321'),
    );
    expect(canvas.operations.indexOf('fill:#654321')).toBeLessThan(
      canvas.operations.indexOf('stroke:#000000'),
    );
    expect(resolved).toMatchObject({
      bounds: rect(220, 48, 25, 12), exclusionBounds: rect(218, 46, 29, 16),
      source: { physicalPageIndex: 1, displayPageNumber: 7 },
    });
    expect(resolved.child.flowDomainId).toBe(domain);
    expect(resolvedSecond.source.sourceBlockIndex).toBe(1);
  });
});

import { describe, expect, it } from 'vitest';
import type { SectionLayoutContext } from '../layout-context.js';
import { physicalSectionGeometry } from './context.js';
import type { BodyLayoutInput } from './body-layout-input.js';
import type { SyntheticBodyLayoutKernel } from './body-layout-kernel.test-support.js';
import { attachSyntheticBodyLayoutKernel as attachBodyLayoutKernel } from './body-layout-kernel.test-support.js';
import { NoteCapacityExceededError } from './body-layout-kernel.js';
import { paginateBody } from './body-paginator.js';
import type { LayoutServices, ParagraphLayout, SourceRef, TableLayout } from './types.js';

const emptyFlowRegistrySnapshot = () => ({
  floats: {
    coordinateSpace: 'logical-page-points' as const,
    flowDomainId: 'body',
    entries: [],
    nextParagraphId: 0,
  },
  drawingCollisions: {
    coordinateSpace: 'logical-page-points' as const,
    flowDomainId: 'body',
    entries: [],
  },
});

const source = (index: number): SourceRef => ({
  story: 'body', storyInstance: 'body', path: [index],
});

const section: SectionLayoutContext = {
  geometry: {
    pageWidth: 200, pageHeight: 100,
    marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 5, footerDistance: 5,
  },
  columns: [{ xPt: 10, wPt: 180 }],
  columnSeparator: false,
  grid: { kind: 'none', linePitchPt: null, charSpacePt: null },
  textDirection: 'lrTb', verticalAlignment: 'top',
};

const paragraph = (id: string, src: SourceRef, heightPt: number): ParagraphLayout => ({
  kind: 'paragraph', id, source: src, flowDomainId: 'acquisition',
  flowBounds: { xPt: 0, yPt: 0, widthPt: 180, heightPt },
  inkBounds: { xPt: 0, yPt: 0, widthPt: 180, heightPt },
  advancePt: heightPt, ordinaryFlow: true,
  spacing: { beforePt: 0, afterPt: 0 }, contextualSpacing: false,
  lines: [], borders: [], resources: [], drawings: [], textBoxes: [], events: [], exclusions: [],
});

const paragraphWithFootnote = (
  id: string,
  src: SourceRef,
  heightPt: number,
  noteId: string,
  kind: 'footnote' | 'endnote' = 'footnote',
): ParagraphLayout => {
  const bounds = { xPt: 0, yPt: 0, widthPt: 180, heightPt };
  return {
    ...paragraph(id, src, heightPt),
    lines: [{
      range: { start: 0, end: 1 }, bounds, baselinePt: heightPt, advancePt: heightPt,
      placements: [{
        kind: 'text', range: { start: 0, end: 1 }, origin: { xPt: 0, yPt: heightPt },
        bounds: { ...bounds, widthPt: 1 }, advancePt: 1, decorations: [],
        noteReference: { kind, id: noteId },
      }],
    }],
  } as unknown as ParagraphLayout;
};

const paragraphWithEndnotes = (
  id: string,
  src: SourceRef,
  heightPt: number,
  noteIds: readonly string[],
): ParagraphLayout => {
  const retained = paragraphWithFootnote(id, src, heightPt, noteIds[0]!, 'endnote');
  const line = retained.lines[0]!;
  const placement = line.placements[0]!;
  return {
    ...retained,
    lines: [{
      ...line,
      placements: noteIds.map((noteId) => ({
        ...placement,
        noteReference: { kind: 'endnote' as const, id: noteId },
      })),
    }],
  };
};

const table = (id: string, src: SourceRef, heightPt: number): TableLayout => ({
  kind: 'table', id, source: src, flowDomainId: 'acquisition',
  flowBounds: { xPt: 0, yPt: 0, widthPt: 180, heightPt },
  inkBounds: { xPt: 0, yPt: 0, widthPt: 180, heightPt },
  advancePt: heightPt, ordinaryFlow: true, columnWidthsPt: [180], rows: [], borders: [],
});

const tableWithFootnote = (
  id: string,
  src: SourceRef,
  heightPt: number,
  noteId: string,
  widthPt = 180,
): TableLayout => {
  const bounds = { xPt: 0, yPt: 0, widthPt, heightPt };
  const noteParagraph = {
    ...paragraph(`${id}:paragraph`, src, heightPt),
    lines: [{
      range: { start: 0, end: 0 }, bounds, baselinePt: heightPt, advancePt: heightPt,
      placements: [{
        kind: 'text', range: { start: 0, end: 0 }, origin: { xPt: 0, yPt: heightPt },
        bounds: { ...bounds, widthPt: 0 }, advancePt: 0,
        noteReference: { kind: 'footnote', id: noteId },
      }],
    }],
  } as unknown as ParagraphLayout;
  return {
    ...table(id, src, heightPt),
    flowBounds: bounds,
    inkBounds: bounds,
    columnWidthsPt: [widthPt],
    rows: [{
      kind: 'table-row', id: `${id}:row`, source: src, flowDomainId: 'acquisition',
      flowBounds: bounds, inkBounds: bounds, advancePt: heightPt, ordinaryFlow: true,
      heightPt, contentHeightPt: heightPt,
      cells: [{
        kind: 'table-cell', id: `${id}:cell`, source: src, flowDomainId: 'acquisition',
        flowBounds: bounds, inkBounds: bounds, advancePt: heightPt, ordinaryFlow: true,
        contentBounds: bounds, verticalMerge: 'none', vAlign: 'top',
        blocks: [{ layout: noteParagraph, offsetPt: 0, advancePt: heightPt }],
      }],
    }],
  };
};

const bodyOwner = () => ({
  sectionOccurrenceId: 'section:0', source: source(2), startType: 'nextPage' as const,
  context: section,
  pageNumbering: { start: null, format: null }, titlePage: false, evenAndOddHeaders: false,
  headers: { default: null, first: null, even: null },
  footers: { default: null, first: null, even: null },
  pageBordersAuthored: false, pageBorders: null,
  pageLayout: {
    physicalGeometry: section.geometry, columns: null, textDirection: 'lrTb' as const,
    gutterPt: 0, rtlGutter: false, mirrorMargins: false, gutterAtTop: false,
    bookFoldPrinting: false, bookFoldRevPrinting: false, printTwoOnOne: false,
  },
});

describe('canonical body producer', () => {
  it('is the sole page owner and returns retained page nodes', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const layouts = [paragraph('p0', source(0), 60), paragraph('p1', source(1), 60)];
    const kernel: SyntheticBodyLayoutKernel = {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => ({
          layout: layouts[input.source.path[0]!]!, blockExtentPt: 60, fragmentation: { kind: 'indivisible' },
        }),
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    };
    attachBodyLayoutKernel(services, kernel);
    const owner = {
      sectionOccurrenceId: 'section:0', source: source(2), startType: 'nextPage' as const,
      context: section,
      pageNumbering: { start: null, format: null }, titlePage: false, evenAndOddHeaders: false,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      pageBordersAuthored: false, pageBorders: null,
      pageLayout: {
        physicalGeometry: section.geometry, columns: null, textDirection: 'lrTb',
        gutterPt: 0, rtlGutter: false, mirrorMargins: false, gutterAtTop: false,
        bookFoldPrinting: false, bookFoldRevPrinting: false, printTwoOnOne: false,
      },
    };
    const input: BodyLayoutInput = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: owner,
      sequence: [0, 1].map((index) => ({
        kind: 'body-block' as const,
        block: {
          kind: 'paragraph' as const, source: source(index), pageBreakBefore: false,
          keepLines: false, keepNext: false, widowControl: true,
          spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
        },
      })),
    };

    const layout = paginateBody(input, services, { currentDateMs: 0 });

    expect(layout.pages).toHaveLength(2);
    expect(layout.pages.map((page) => page.layers.body.map((node) => node.source.path[0])))
      .toEqual([[0], [1]]);
    expect(layout.pages.map((page) => page.readingOrder.length)).toEqual([1, 1]);
    expect(Object.isFrozen(layout)).toBe(true);
  });

  it('retains one canonical boundary across collapsed fractional paragraph spacing', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const fractionalSection: SectionLayoutContext = {
      ...section,
      geometry: {
        ...section.geometry,
        pageHeight: 500,
        marginTop: 236.24326171875,
        marginBottom: 10,
      },
    };
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => {
          const layout = paragraph(`p${input.source.path[0]}`, input.source, 20);
          return {
            layout: {
              ...layout,
              spacing: { beforePt: 0, afterPt: 10 },
            },
            blockExtentPt: 20,
            fragmentation: { kind: 'indivisible' },
          };
        },
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const owner = {
      ...bodyOwner(),
      context: fractionalSection,
      pageLayout: {
        ...bodyOwner().pageLayout,
        physicalGeometry: fractionalSection.geometry,
      },
    };
    const input: BodyLayoutInput = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: owner,
      sequence: [0, 1].map((index) => ({
        kind: 'body-block' as const,
        block: {
          kind: 'paragraph' as const, source: source(index), pageBreakBefore: false,
          keepLines: false, keepNext: false, widowControl: true,
          spaceBeforePt: 0, spaceAfterPt: 10, contextualSpacing: true, styleId: 'same',
        },
      })),
    };

    const layout = paginateBody(input, services, { currentDateMs: 0 });
    const [first, second] = layout.pages[0]!.layers.body;

    expect(first!.flowBounds.yPt + first!.flowBounds.heightPt).toBe(second!.flowBounds.yPt);
  });

  it('retains endnotes after terminal body flow in authored order', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => ({
          layout: paragraphWithEndnotes('body', input.source, 20, ['end-1', 'end-2']),
          blockExtentPt: 20,
          fragmentation: { kind: 'splittable', lineEndBoundaries: [{ segIndex: 0, charOffset: 1 }] },
        }),
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: ({ referenceIds }) => referenceIds.length * 10,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const input: BodyLayoutInput = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: bodyOwner(),
      sequence: [{
        kind: 'body-block',
        block: {
          kind: 'paragraph', source: source(0), pageBreakBefore: false,
          keepLines: false, keepNext: false, widowControl: true,
          spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
        },
      }],
      endnoteIds: ['end-1', 'end-2'],
    };

    const layout = paginateBody(input, services, { currentDateMs: 0 });
    const page = layout.pages[0]!;
    const note = page.layers.notes[0]!;

    expect(layout.diagnostics).toEqual([]);
    expect(note).toMatchObject({
      kind: 'note',
      source: { story: 'endnote', storyInstance: 'end-1' },
      flowDomainId: 'endnotes:page:0',
      flowBounds: { yPt: 30, heightPt: 20 },
    });
    expect(page.flowDomains.find((domain) => domain.kind === 'endnote')?.sectionRegionId)
      .toBe(page.sectionRegions[0]!.id);
    expect(page.readingOrder).toEqual([page.layers.body[0]!.id, note.id]);
  });

  it('reports a retained endnote story that cannot fit the terminal flow region', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => ({
          layout: paragraphWithFootnote('body', input.source, 20, 'end-1', 'endnote'),
          blockExtentPt: 20,
          fragmentation: { kind: 'splittable', lineEndBoundaries: [{ segIndex: 0, charOffset: 1 }] },
        }),
        measureTable: () => { throw new Error('unused'); },
        layoutNotes: () => {
          throw new NoteCapacityExceededError('endnote', 0, 'endnotes:page:0');
        },
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });

    const layout = paginateBody({
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: bodyOwner(),
      sequence: [{
        kind: 'body-block',
        block: {
          kind: 'paragraph', source: source(0), pageBreakBefore: false,
          keepLines: false, keepNext: false, widowControl: true,
          spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
        },
      }],
      endnoteIds: ['end-1'],
    }, services, { currentDateMs: 0 });

    expect(layout.pages[0]!.layers.notes).toEqual([]);
    expect(layout.diagnostics).toEqual([
      expect.objectContaining({
        code: 'UNSUPPORTED_FEATURE',
        severity: 'error',
        source: expect.objectContaining({ story: 'endnote', storyInstance: 'end-1' }),
        message: expect.stringContaining('do not fit the retained terminal flow region'),
      }),
    ]);
  });

  it.each([
    new Error('endnote implementation bug'),
    new NoteCapacityExceededError('endnote', 1, 'endnotes:page:1'),
  ])('does not misreport an unrelated failure as the terminal endnote capacity limit', (failure) => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => ({
          layout: paragraphWithFootnote('body', input.source, 20, 'end-1', 'endnote'),
          blockExtentPt: 20,
          fragmentation: { kind: 'splittable', lineEndBoundaries: [{ segIndex: 0, charOffset: 1 }] },
        }),
        measureTable: () => { throw new Error('unused'); },
        layoutNotes: () => { throw failure; },
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });

    expect(() => paginateBody({
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: bodyOwner(),
      sequence: [{
        kind: 'body-block',
        block: {
          kind: 'paragraph', source: source(0), pageBreakBefore: false,
          keepLines: false, keepNext: false, widowControl: true,
          spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
        },
      }],
      endnoteIds: ['end-1'],
    }, services, { currentDateMs: 0 })).toThrow(failure.message);
  });

  it('places vertical footnotes against the logical block edge and retains physical bounds', () => {
    const verticalSection: SectionLayoutContext = {
      ...section,
      textDirection: 'tbRl',
    };
    const owner = {
      ...bodyOwner(),
      context: verticalSection,
      pageLayout: {
        ...bodyOwner().pageLayout,
        physicalGeometry: physicalSectionGeometry(verticalSection.geometry),
        textDirection: 'tbRl' as const,
      },
    };
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => ({
          layout: paragraphWithFootnote('vertical-body', input.source, 20, 'vertical-note'),
          blockExtentPt: 20,
          fragmentation: { kind: 'splittable', lineEndBoundaries: [{ segIndex: 0, charOffset: 1 }] },
        }),
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: ({ referenceIds }) => referenceIds.length > 0 ? 10 : 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const input: BodyLayoutInput = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: owner,
      sequence: [{
        kind: 'body-block',
        block: {
          kind: 'paragraph', source: source(0), pageBreakBefore: false,
          keepLines: false, keepNext: false, widowControl: true,
          spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
        },
      }],
    };

    const layout = paginateBody(input, services, { currentDateMs: 0 });
    const page = layout.pages[0]!;
    const note = page.layers.notes[0]!;
    const domain = page.flowDomains.find((candidate) => candidate.kind === 'footnote')!;

    expect(note.flowBounds).toMatchObject({ yPt: 80, heightPt: 10 });
    expect(domain.logicalBounds).toEqual({ xPt: 10, yPt: 80, widthPt: 180, heightPt: 10 });
    expect(domain.physicalBounds).toEqual({ xPt: 10, yPt: 10, widthPt: 10, heightPt: 180 });
    expect(domain.sectionRegionId).toBe(page.sectionRegions[0]!.id);
  });

  it('retains nextColumn section ownership as disjoint same-page column regions', () => {
    const columns = [{ xPt: 10, wPt: 80 }, { xPt: 110, wPt: 80 }];
    const twoColumnSection: SectionLayoutContext = {
      ...section,
      columns,
      verticalAlignment: 'center',
    };
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => {
          const layout = paragraph(`p${input.source.path[0]}`, input.source, 20);
          return {
            layout: {
              ...layout,
              flowBounds: { ...layout.flowBounds, widthPt: 80 },
              inkBounds: { ...layout.inkBounds, widthPt: 80 },
            },
            blockExtentPt: 20,
            fragmentation: { kind: 'indivisible' },
          };
        },
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const owner = (sectionOccurrenceId: string, startType: 'nextPage' | 'nextColumn') => ({
      sectionOccurrenceId,
      source: source(sectionOccurrenceId === 'section:outgoing' ? 10 : 11),
      startType,
      context: twoColumnSection,
      pageNumbering: { start: null, format: null },
      titlePage: false,
      evenAndOddHeaders: false,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      pageBordersAuthored: false,
      pageBorders: null,
      pageLayout: {
        physicalGeometry: twoColumnSection.geometry,
        columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] },
        textDirection: 'lrTb',
        gutterPt: 0,
        rtlGutter: false,
        mirrorMargins: false,
        gutterAtTop: false,
        bookFoldPrinting: false,
        bookFoldRevPrinting: false,
        printTwoOnOne: false,
      },
    });
    const outgoing = owner('section:outgoing', 'nextPage');
    const incoming = owner('section:incoming', 'nextColumn');
    const bodyBlock = (index: number) => ({
      kind: 'body-block' as const,
      block: {
        kind: 'paragraph' as const,
        source: source(index),
        pageBreakBefore: false,
        keepLines: false,
        keepNext: false,
        widowControl: true,
        spaceBeforePt: 0,
        spaceAfterPt: 0,
        contextualSpacing: false,
        styleId: null,
      },
    });

    const layout = paginateBody({
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: outgoing,
      sequence: [
        bodyBlock(0),
        { kind: 'begin-section', source: source(9), section: incoming },
        bodyBlock(1),
      ],
    }, services, { currentDateMs: 0 });

    expect(layout.pages).toHaveLength(1);
    expect(layout.pages[0]!.sectionRegions.map((region) => ({
      section: region.sectionOccurrenceId,
      columns: region.columnIndexes,
      block: [region.blockStartPt, region.blockEndPt],
    }))).toEqual([
      { section: 'section:outgoing', columns: [0], block: [10, 90] },
      { section: 'section:incoming', columns: [1], block: [10, 90] },
    ]);
    expect(layout.pages[0]!.layers.body.map((node) => node.flowDomainId)).toEqual([
      layout.pages[0]!.sectionRegions[0]!.flowDomainIds[0],
      layout.pages[0]!.sectionRegions[1]!.flowDomainIds[0],
    ]);
    expect(layout.pages[0]!.layers.body.map((node) => node.flowBounds.yPt))
      .toEqual([40, 40]);
  });

  it.each([
    { name: 'LTR to RTL', outgoingBidi: false, incomingBidi: true },
    { name: 'RTL to LTR', outgoingBidi: true, incomingBidi: false },
  ])('rejects an overlapping $name nextColumn ownership cutover before page mutation', ({
    outgoingBidi,
    incomingBidi,
  }) => {
    const columns = [
      { xPt: 10, wPt: 50 },
      { xPt: 75, wPt: 50 },
      { xPt: 140, wPt: 50 },
    ];
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => {
          const layout = paragraph(`p${input.source.path[0]}`, input.source, 20);
          return {
            layout: {
              ...layout,
              flowBounds: { ...layout.flowBounds, widthPt: 50 },
              inkBounds: { ...layout.inkBounds, widthPt: 50 },
            },
            blockExtentPt: 20,
            fragmentation: { kind: 'indivisible' },
          };
        },
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const owner = (
      id: string,
      startType: 'nextPage' | 'nextColumn',
      sectionBidi: boolean,
    ) => {
      const context: SectionLayoutContext = { ...section, columns, sectionBidi };
      return {
        sectionOccurrenceId: id,
        source: source(id === 'section:outgoing' ? 10 : 11),
        startType,
        context,
        pageNumbering: { start: null, format: null },
        titlePage: false,
        evenAndOddHeaders: false,
        headers: { default: null, first: null, even: null },
        footers: { default: null, first: null, even: null },
        pageBordersAuthored: false,
        pageBorders: null,
        pageLayout: {
          physicalGeometry: context.geometry,
          columns: { count: 3, spacePt: 15, equalWidth: true, sep: false, cols: [] },
          textDirection: 'lrTb',
          gutterPt: 0,
          rtlGutter: false,
          mirrorMargins: false,
          gutterAtTop: false,
          bookFoldPrinting: false,
          bookFoldRevPrinting: false,
          printTwoOnOne: false,
        },
      };
    };
    const outgoing = owner('section:outgoing', 'nextPage', outgoingBidi);
    const incoming = owner('section:incoming', 'nextColumn', incomingBidi);

    expect(() => paginateBody({
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: outgoing,
      sequence: [
        {
          kind: 'body-block',
          block: {
            kind: 'paragraph', source: source(0), pageBreakBefore: false,
            keepLines: false, keepNext: false, widowControl: true,
            spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
          },
        },
        { kind: 'begin-section', source: source(9), section: incoming },
      ],
    }, services, { currentDateMs: 0 })).toThrow(expect.objectContaining({
      code: 'NEXT_COLUMN_DESTINATION_UNAVAILABLE',
      reason: 'physical-overlap',
    }));

    const recovered = paginateBody({
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: outgoing,
      sequence: [
        {
          kind: 'body-block',
          block: {
            kind: 'paragraph', source: source(0), pageBreakBefore: false,
            keepLines: false, keepNext: false, widowControl: true,
            spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
          },
        },
        { kind: 'authored-break', source: source(1), break: 'page' },
        {
          kind: 'body-block',
          block: {
            kind: 'paragraph', source: source(2), pageBreakBefore: false,
            keepLines: false, keepNext: false, widowControl: true,
            spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
          },
        },
        { kind: 'begin-section', source: source(9), section: incoming },
      ],
    }, services, { currentDateMs: 0 });

    expect(recovered.pages).toHaveLength(1);
    expect(recovered.pages[0]!.pageIndex).toBe(0);
    expect(recovered.pages[0]!.layers.body.map((node) => node.source.path)).toEqual([[0]]);
    expect(recovered.diagnostics).toContainEqual(expect.objectContaining({
      code: 'UNSUPPORTED_FEATURE',
      severity: 'error',
      source: source(9),
    }));
  });

  it('keeps pageBreakBefore on the first RTL population column of an empty page', () => {
    const rtlSection: SectionLayoutContext = {
      ...section,
      sectionBidi: true,
      columns: [{ xPt: 10, wPt: 80 }, { xPt: 110, wPt: 80 }],
    };
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => {
          const layout = paragraph('rtl-page-start', input.source, 20);
          return {
            layout: {
              ...layout,
              flowBounds: { ...layout.flowBounds, widthPt: 80 },
              inkBounds: { ...layout.inkBounds, widthPt: 80 },
            },
            blockExtentPt: 20,
            fragmentation: { kind: 'indivisible' },
          };
        },
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const owner = {
      ...bodyOwner(),
      context: rtlSection,
      pageLayout: {
        ...bodyOwner().pageLayout,
        physicalGeometry: rtlSection.geometry,
        columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] },
      },
    };

    const layout = paginateBody({
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: owner,
      sequence: [{
        kind: 'body-block',
        block: {
          kind: 'paragraph', source: source(0), pageBreakBefore: true,
          keepLines: false, keepNext: false, widowControl: true,
          spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
        },
      }],
    }, services, { currentDateMs: 0 });

    expect(layout.pages).toHaveLength(1);
    expect(layout.pages[0]!.layers.body[0]!.flowDomainId)
      .toContain(':column:1');
  });

  it('commits neither registry component for a rejected body candidate', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const baseFloatEntries = Object.freeze([]);
    const baseCollisionEntries = Object.freeze([]);
    const commits: unknown[] = [];
    let secondMeasurements = 0;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: () => { throw new Error('unused'); },
        measureTable: (request) => {
          const index = request.input.source.path[0]!;
          if (index === 0) {
            return { layout: table('first', request.input.source, 60), blockExtentPt: 60 };
          }
          secondMeasurements += 1;
          const flowRegistryDelta = {
            floats: {
              coordinateSpace: 'logical-page-points' as const,
              flowDomainId: request.location.flowDomainId,
              baseEntries: baseFloatEntries,
              baseNextParagraphId: 0,
              nextParagraphId: 1,
              entries: [{
                kind: 'shape' as const,
                occurrenceId: 'candidate-float',
                paragraphId: 0,
                bounds: { xPt: 0, yPt: 0, widthPt: 10, heightPt: 10 },
                exclusionBounds: { xPt: 0, yPt: 0, widthPt: 10, heightPt: 10 },
              }],
            },
            drawingCollisions: {
              coordinateSpace: 'logical-page-points' as const,
              flowDomainId: request.location.flowDomainId,
              baseEntries: baseCollisionEntries,
              baseEntryCount: 0,
              entries: [{
                occurrenceId: 'candidate-collision',
                bounds: { xPt: 0, yPt: 0, widthPt: 10, heightPt: 10 },
                horizontalOwnership: 'page' as const,
                verticalOwnership: 'page' as const,
              }],
            },
          };
          return request.availableBlockExtentPt < 40
            ? {
                layout: table('rejected', request.input.source, 40),
                blockExtentPt: 40,
                requiresFreshFlowRegion: true,
                flowRegistryDelta,
              }
            : {
                layout: table('accepted', request.input.source, 40),
                blockExtentPt: 40,
                flowRegistryDelta,
              };
        },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: (delta) => { commits.push(delta); },
      }),
    });
    const input = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: bodyOwner(),
      sequence: [0, 1].map((index) => ({
        kind: 'body-block',
        block: { kind: 'table', source: source(index) },
      })),
    } as unknown as BodyLayoutInput;

    paginateBody(input, services, { currentDateMs: 0 });

    expect(secondMeasurements).toBe(2);
    expect(commits).toHaveLength(1);
    expect(commits[0]).toMatchObject({
      floats: { entries: [{ occurrenceId: 'candidate-float' }] },
      drawingCollisions: { entries: [{ occurrenceId: 'candidate-collision' }] },
    });
  });

  it('acquires an adjacent-table sequence as one cursor-bearing logical table', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const acquiredInputs: string[] = [];
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: () => { throw new Error('unused'); },
        measureTable: ({ input }) => {
          acquiredInputs.push(input.kind);
          return { layout: table('logical-table', input.source, 30), blockExtentPt: 30 };
        },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const owner = {
      sectionOccurrenceId: 'section:0', source: source(2), startType: 'nextPage' as const,
      context: section,
      pageNumbering: { start: null, format: null }, titlePage: false, evenAndOddHeaders: false,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      pageBordersAuthored: false, pageBorders: null,
      pageLayout: {
        physicalGeometry: section.geometry, columns: null, textDirection: 'lrTb',
        gutterPt: 0, rtlGutter: false, mirrorMargins: false, gutterAtTop: false,
        bookFoldPrinting: false, bookFoldRevPrinting: false, printTwoOnOne: false,
      },
    };
    const layout = paginateBody({
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: owner,
      sequence: [{
        kind: 'adjacent-table-group', logicalSequenceId: 'logical:0', source: source(0),
        tables: [
          { kind: 'table', source: source(0), rowCount: 1 },
          { kind: 'table', source: source(1), rowCount: 2 },
        ],
      }],
    }, services, { currentDateMs: 0 });

    expect(acquiredInputs).toEqual(['adjacent-table-group']);
    expect(layout.pages[0]!.layers.body.map((node) => node.source.path)).toEqual([[0]]);
  });

  it('admits a table fragment whose charge matches the region within floating-point precision', () => {
    // The region extent (blockEnd - blockStart) and the fragment charge (a sum
    // of retained row advances) reconstruct the same authored boundary through
    // different arithmetic, so they can disagree by single-ULP drift. With
    // zero footnote reserve the retry request is identical to the first, so a
    // raw `>` fit test can never converge on such a fragment.
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: () => { throw new Error('unused'); },
        measureTable: (request) => {
          if (!request.cursor) {
            const extentPt = request.availableBlockExtentPt
              + Number.EPSILON * request.availableBlockExtentPt / 2;
            return {
              layout: table('head-slice', request.input.source, extentPt),
              blockExtentPt: extentPt,
              nextCursor: {
                kind: 'table' as const,
                cursor: { rowIndex: 1, rowFragmentIndex: 0, cells: [] },
              },
            };
          }
          return {
            layout: table('tail-slice', request.input.source, 30),
            blockExtentPt: 30,
          };
        },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const input = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: bodyOwner(),
      sequence: [{
        kind: 'body-block',
        block: { kind: 'table', source: source(0) },
      }],
    } as unknown as BodyLayoutInput;

    const layout = paginateBody(input, services, { currentDateMs: 0 });

    expect(layout.pages).toHaveLength(2);
    expect(layout.pages[0]!.layers.body[0]!.id).toContain('head-slice');
    expect(layout.pages[1]!.layers.body[0]!.id).toContain('tail-slice');
  });

  it('still reports non-convergence for table overshoot beyond floating-point precision', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: () => { throw new Error('unused'); },
        measureTable: (request) => {
          const extentPt = request.availableBlockExtentPt + 1;
          return {
            layout: table('oversized-slice', request.input.source, extentPt),
            blockExtentPt: extentPt,
            nextCursor: {
              kind: 'table' as const,
              cursor: { rowIndex: 1, rowFragmentIndex: 0, cells: [] },
            },
          };
        },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const input = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: bodyOwner(),
      sequence: [{
        kind: 'body-block',
        block: { kind: 'table', source: source(0) },
      }],
    } as unknown as BodyLayoutInput;

    expect(() => paginateBody(input, services, { currentDateMs: 0 }))
      .toThrow('Table footnote admission did not converge');
  });

  it('admits references from each accepted split-table slice only', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const admissions: Array<{ pageIndex: number; referenceIds: readonly string[] }> = [];
    let measuredPageIndex = -1;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: () => { throw new Error('unused'); },
        measureTable: (request) => {
          measuredPageIndex = request.location.pageIndex;
          return request.cursor
            ? { layout: table('terminal-slice', request.input.source, 30), blockExtentPt: 30 }
            : {
                layout: tableWithFootnote('reference-slice', request.input.source, 30, 'fn-first'),
                blockExtentPt: 30,
                nextCursor: {
                  kind: 'table' as const,
                  cursor: { rowIndex: 1, rowFragmentIndex: 0, cells: [] },
                },
              };
        },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: ({ referenceIds }) => {
          admissions.push({ pageIndex: measuredPageIndex, referenceIds: [...referenceIds] });
          return referenceIds.length === 0 ? 0 : 5;
        },
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const legacySourceWideInput = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: bodyOwner(),
      sequence: [{
        kind: 'body-block',
        block: { kind: 'table', source: source(0), footnoteReferenceIds: ['fn-first'] },
      }],
    } as unknown as BodyLayoutInput;

    const layout = paginateBody(legacySourceWideInput, services, { currentDateMs: 0 });

    expect(layout.pages).toHaveLength(2);
    expect(admissions).toEqual([
      { pageIndex: 0, referenceIds: ['fn-first'] },
    ]);
  });

  it('does not union source-wide references across an adjacent table group', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const admissions: string[][] = [];
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: () => { throw new Error('unused'); },
        measureTable: ({ input }) => ({
          layout: tableWithFootnote('retained-left', input.source, 30, 'left'), blockExtentPt: 30,
        }),
        measureStoryExtent: () => 0,
        measureFootnoteReserve: ({ referenceIds }) => {
          admissions.push([...referenceIds]);
          return 0;
        },
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const legacySourceWideInput = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: bodyOwner(),
      sequence: [{
        kind: 'adjacent-table-group', logicalSequenceId: 'logical:0', source: source(0),
        tables: [
          { kind: 'table', source: source(0), rowCount: 1, footnoteReferenceIds: ['left'] },
          { kind: 'table', source: source(1), rowCount: 1, footnoteReferenceIds: ['right'] },
        ],
      }],
    } as unknown as BodyLayoutInput;

    paginateBody(legacySourceWideInput, services, { currentDateMs: 0 });

    expect(admissions).toEqual([['left']]);
  });

  it('retains a zero-reserve footnote reference across a continuous section boundary', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => {
          const index = input.source.path[0]!;
          const layout = index === 0
            ? paragraphWithFootnote('note', input.source, 10, 'zero-height')
            : paragraph('after', input.source, 10);
          return {
            layout,
            blockExtentPt: 10,
            fragmentation: index === 0
              ? { kind: 'splittable', lineEndBoundaries: [{ segIndex: 0, charOffset: 1 }] }
              : { kind: 'indivisible' },
          };
        },
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: () => ({ fullExtentPt: 10, leadContentExtentPt: 10 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const incoming = {
      ...bodyOwner(),
      sectionOccurrenceId: 'section:1',
      startType: 'continuous' as const,
    };
    const block = (index: number) => ({
      kind: 'body-block' as const,
      block: {
        kind: 'paragraph' as const, source: source(index), pageBreakBefore: false,
        keepLines: false, keepNext: false, widowControl: true,
        spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
      },
    });

    const layout = paginateBody({
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: bodyOwner(),
      sequence: [
        block(0),
        { kind: 'begin-section', source: source(2), section: incoming },
        block(1),
      ],
    }, services, { currentDateMs: 0 });

    expect(layout.pages.map((page) => page.layers.body.map((node) => node.source.path[0])))
      .toEqual([[0], [1]]);
  });

  it('uses retained reference presence, not reserve height, for first-on-page measurement', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const admissions: Array<Readonly<{
      referenceIds: readonly string[];
      firstOnPage: boolean;
    }>> = [];
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => {
          const index = input.source.path[0]!;
          return {
            layout: paragraphWithFootnote(`note-${index}`, input.source, 10, `fn-${index}`),
            blockExtentPt: 10,
            fragmentation: { kind: 'splittable', lineEndBoundaries: [{ segIndex: 0, charOffset: 1 }] },
          };
        },
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: ({ referenceIds, firstOnPage }) => {
          if (referenceIds.length > 0) {
            admissions.push({ referenceIds: [...referenceIds], firstOnPage });
          }
          return 0;
        },
        measureFollowingBlock: () => ({ fullExtentPt: 10, leadContentExtentPt: 10 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const input: BodyLayoutInput = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: bodyOwner(),
      sequence: [0, 1].map((index) => ({
        kind: 'body-block' as const,
        block: {
          kind: 'paragraph' as const, source: source(index), pageBreakBefore: false,
          keepLines: false, keepNext: false, widowControl: true,
          spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
        },
      })),
    };

    paginateBody(input, services, { currentDateMs: 0 });

    expect(admissions.filter(({ referenceIds }) => referenceIds[0] === 'fn-0'))
      .not.toHaveLength(0);
    expect(admissions.filter(({ referenceIds }) => referenceIds[0] === 'fn-0')
      .every(({ firstOnPage }) => firstOnPage)).toBe(true);
    expect(admissions.filter(({ referenceIds }) => referenceIds[0] === 'fn-1'))
      .not.toHaveLength(0);
    expect(admissions.filter(({ referenceIds }) => referenceIds[0] === 'fn-1')
      .every(({ firstOnPage }) => !firstOnPage)).toBe(true);
  });

  it('admits a placed paragraph footnote before selecting its flow region', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => {
          const index = input.source.path[0]!;
          if (index === 1) {
            const layout = paragraphWithFootnote(
              'frame-note',
              input.source,
              20,
              'frame-footnote',
            );
            return {
              layout: { ...layout, ordinaryFlow: false },
              blockExtentPt: 0,
              fragmentation: { kind: 'splittable', lineEndBoundaries: [{ segIndex: 0, charOffset: 1 }] },
              placement: {
                coordinateSpace: 'logical-body' as const,
                xPt: 10,
                yPt: 70,
                sectionFlowOwnership: 'host-flow' as const,
              },
              relocationBlockExtentPt: 20,
            };
          }
          const heightPt = 60;
          return {
            layout: paragraph(`p-${index}`, input.source, heightPt),
            blockExtentPt: heightPt,
            fragmentation: { kind: 'indivisible' },
          };
        },
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: ({ referenceIds }) => (
          referenceIds.includes('frame-footnote') ? 10 : 0
        ),
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const input: BodyLayoutInput = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: bodyOwner(),
      sequence: [0, 1, 2].map((index) => ({
        kind: 'body-block' as const,
        block: {
          kind: 'paragraph' as const, source: source(index), pageBreakBefore: false,
          keepLines: false, keepNext: false, widowControl: true,
          spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
        },
      })),
    };

    const layout = paginateBody(input, services, { currentDateMs: 0 });

    expect(layout.pages.map((page) => page.layers.body.map((node) => node.source.path[0])))
      .toEqual([[0], [1, 2]]);
  });

  it('retains a placed paragraph footnote across a continuous section boundary', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => {
          const index = input.source.path[0]!;
          if (index === 0) {
            const layout = paragraphWithFootnote(
              'frame-note',
              input.source,
              10,
              'frame-zero-height',
            );
            return {
              layout: { ...layout, ordinaryFlow: false },
              blockExtentPt: 0,
              fragmentation: { kind: 'splittable', lineEndBoundaries: [{ segIndex: 0, charOffset: 1 }] },
              placement: {
                coordinateSpace: 'logical-body' as const,
                xPt: 10,
                yPt: 10,
                sectionFlowOwnership: 'host-flow' as const,
              },
              relocationBlockExtentPt: 10,
            };
          }
          return {
            layout: paragraph('after', input.source, 10),
            blockExtentPt: 10,
            fragmentation: { kind: 'indivisible' },
          };
        },
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: () => ({ fullExtentPt: 10, leadContentExtentPt: 10 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const incoming = {
      ...bodyOwner(),
      sectionOccurrenceId: 'section:1',
      startType: 'continuous' as const,
    };
    const block = (index: number) => ({
      kind: 'body-block' as const,
      block: {
        kind: 'paragraph' as const, source: source(index), pageBreakBefore: false,
        keepLines: false, keepNext: false, widowControl: true,
        spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
      },
    });

    const layout = paginateBody({
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: bodyOwner(),
      sequence: [
        block(0),
        { kind: 'begin-section', source: source(2), section: incoming },
        block(1),
      ],
    }, services, { currentDateMs: 0 });

    expect(layout.pages.map((page) => page.layers.body.map((node) => node.source.path[0])))
      .toEqual([[0], [1]]);
  });

  it('moves a page-owned placed footnote to a fresh physical page before committing it', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const acquiredAt: Array<Readonly<{ pageIndex: number; columnIndex: number }>> = [];
    const admissions: Array<Readonly<{
      pageIndex: number;
      referenceIds: readonly string[];
      firstOnPage: boolean;
    }>> = [];
    const committedRegistryEntries: string[] = [];
    let measuredPageIndex = -1;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input, location }) => {
          const index = input.source.path[0]!;
          if (index === 0) {
            return {
              layout: paragraph('prior', input.source, 75),
              blockExtentPt: 75,
              fragmentation: { kind: 'indivisible' },
            };
          }
          acquiredAt.push({ pageIndex: location.pageIndex, columnIndex: location.columnIndex });
          measuredPageIndex = location.pageIndex;
          const layout = paragraphWithFootnote('page-frame', input.source, 10, 'page-note');
          return {
            layout: { ...layout, ordinaryFlow: false },
            blockExtentPt: 0,
            fragmentation: { kind: 'indivisible' },
            placement: {
              coordinateSpace: 'logical-body' as const,
              xPt: 10,
              yPt: 10,
              sectionFlowOwnership: 'page' as const,
            },
            flowRegistryDelta: {
              floats: {
                coordinateSpace: 'logical-page-points' as const,
                flowDomainId: location.flowDomainId,
                baseEntries: [],
                baseNextParagraphId: 0,
                nextParagraphId: 1,
                entries: [{
                  kind: 'frame' as const,
                  occurrenceId: `frame-at-page-${location.pageIndex}`,
                  paragraphId: 0,
                  bounds: { xPt: 10, yPt: 10, widthPt: 20, heightPt: 10 },
                  exclusionBounds: { xPt: 10, yPt: 10, widthPt: 20, heightPt: 10 },
                }],
              },
            },
          };
        },
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: ({ referenceIds, firstOnPage }) => {
          if (referenceIds.length > 0) {
            admissions.push({
              pageIndex: measuredPageIndex,
              referenceIds: [...referenceIds],
              firstOnPage,
            });
          }
          return referenceIds.includes('page-note') ? 10 : 0;
        },
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: (delta) => {
          committedRegistryEntries.push(...(delta.floats?.entries.map((entry) => entry.occurrenceId) ?? []));
        },
      }),
    });
    const input: BodyLayoutInput = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: bodyOwner(),
      sequence: [0, 1].map((index) => ({
        kind: 'body-block' as const,
        block: {
          kind: 'paragraph' as const, source: source(index), pageBreakBefore: false,
          keepLines: false, keepNext: false, widowControl: true,
          spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
        },
      })),
    };

    const layout = paginateBody(input, services, { currentDateMs: 0 });

    expect(layout.pages.map((page) => page.layers.body.map((node) => node.source.path[0])))
      .toEqual([[0], [1]]);
    expect(acquiredAt).toEqual([
      { pageIndex: 0, columnIndex: 0 },
      { pageIndex: 1, columnIndex: 0 },
    ]);
    expect(admissions).toEqual([
      { pageIndex: 0, referenceIds: ['page-note'], firstOnPage: true },
      { pageIndex: 1, referenceIds: ['page-note'], firstOnPage: true },
    ]);
    expect(committedRegistryEntries).toEqual(['frame-at-page-1']);
  });

  it('moves host-flow placed footnote overflow past all same-page columns', () => {
    const twoColumnSection: SectionLayoutContext = {
      ...section,
      columns: [{ xPt: 10, wPt: 80 }, { xPt: 110, wPt: 80 }],
    };
    const owner = {
      ...bodyOwner(),
      context: twoColumnSection,
      pageLayout: {
        ...bodyOwner().pageLayout,
        columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] },
      },
    };
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const acquiredAt: Array<Readonly<{ pageIndex: number; columnIndex: number }>> = [];
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input, location }) => {
          const index = input.source.path[0]!;
          if (index === 0) {
            const retained = paragraph('prior', input.source, 80);
            return {
              layout: {
                ...retained,
                flowBounds: { ...retained.flowBounds, widthPt: 80 },
                inkBounds: { ...retained.inkBounds, widthPt: 80 },
              },
              blockExtentPt: 80,
              fragmentation: { kind: 'indivisible' },
            };
          }
          acquiredAt.push({ pageIndex: location.pageIndex, columnIndex: location.columnIndex });
          const retained = paragraphWithFootnote('host-frame', input.source, 4, 'host-note');
          return {
            layout: {
              ...retained,
              ordinaryFlow: false,
              flowBounds: { ...retained.flowBounds, widthPt: 80 },
              inkBounds: { ...retained.inkBounds, widthPt: 80 },
            },
            blockExtentPt: 0,
            fragmentation: { kind: 'indivisible' },
            placement: {
              coordinateSpace: 'logical-body' as const,
              xPt: location.cursorPt.xPt,
              yPt: location.cursorPt.yPt,
              sectionFlowOwnership: 'host-flow' as const,
            },
            relocationBlockExtentPt: 4,
          };
        },
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: ({ referenceIds }) => referenceIds.includes('host-note') ? 10 : 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const input: BodyLayoutInput = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: owner,
      sequence: [{
        kind: 'body-block',
        block: {
          kind: 'paragraph', source: source(0), pageBreakBefore: false,
          keepLines: false, keepNext: false, widowControl: true,
          spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
        },
      }, {
        kind: 'authored-break',
        source: source(0),
        break: 'column',
      }, {
        kind: 'body-block',
        block: {
          kind: 'paragraph', source: source(1), pageBreakBefore: false,
          keepLines: false, keepNext: false, widowControl: true,
          spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
        },
      }],
    };

    const layout = paginateBody(input, services, { currentDateMs: 0 });

    expect(layout.pages.map((page) => page.layers.body.map((node) => node.source.path[0])))
      .toEqual([[0], [1]]);
    expect(acquiredAt).toEqual([
      { pageIndex: 0, columnIndex: 1 },
      { pageIndex: 1, columnIndex: 0 },
    ]);
  });

  it('moves host-flow placed footnote overflow to a safe same-page column', () => {
    const twoColumnSection: SectionLayoutContext = {
      ...section,
      columns: [{ xPt: 10, wPt: 80 }, { xPt: 110, wPt: 80 }],
    };
    const owner = {
      ...bodyOwner(),
      context: twoColumnSection,
      pageLayout: {
        ...bodyOwner().pageLayout,
        columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] },
      },
    };
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const acquiredAt: Array<Readonly<{ pageIndex: number; columnIndex: number }>> = [];
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input, location }) => {
          const index = input.source.path[0]!;
          if (index === 0) {
            const retained = paragraphWithFootnote('prior', input.source, 60, 'prior-note');
            return {
              layout: {
                ...retained,
                flowBounds: { ...retained.flowBounds, widthPt: 80 },
                inkBounds: { ...retained.inkBounds, widthPt: 80 },
              },
              blockExtentPt: 60,
              fragmentation: { kind: 'splittable', lineEndBoundaries: [{ segIndex: 0, charOffset: 1 }] },
            };
          }
          acquiredAt.push({ pageIndex: location.pageIndex, columnIndex: location.columnIndex });
          const retained = paragraphWithFootnote('host-frame', input.source, 4, 'host-note');
          return {
            layout: {
              ...retained,
              ordinaryFlow: false,
              flowBounds: { ...retained.flowBounds, widthPt: 80 },
              inkBounds: { ...retained.inkBounds, widthPt: 80 },
            },
            blockExtentPt: 0,
            fragmentation: { kind: 'indivisible' },
            placement: {
              coordinateSpace: 'logical-body' as const,
              xPt: location.cursorPt.xPt,
              yPt: location.cursorPt.yPt,
              sectionFlowOwnership: 'host-flow' as const,
            },
            relocationBlockExtentPt: 4,
          };
        },
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: ({ referenceIds }) => referenceIds.length > 0 ? 10 : 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const input: BodyLayoutInput = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: owner,
      sequence: [0, 1].map((index) => ({
        kind: 'body-block' as const,
        block: {
          kind: 'paragraph' as const, source: source(index), pageBreakBefore: false,
          keepLines: false, keepNext: false, widowControl: true,
          spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
        },
      })),
    };

    const layout = paginateBody(input, services, { currentDateMs: 0 });

    expect(layout.pages.map((page) => page.layers.body.map((node) => node.source.path[0])))
      .toEqual([[0, 1]]);
    expect(acquiredAt).toEqual([
      { pageIndex: 0, columnIndex: 0 },
      { pageIndex: 0, columnIndex: 1 },
    ]);
  });

  it('moves an ordinary paragraph footnote past prior-column content it would invade', () => {
    const twoColumnSection: SectionLayoutContext = {
      ...section,
      columns: [{ xPt: 10, wPt: 80 }, { xPt: 110, wPt: 80 }],
    };
    const owner = {
      ...bodyOwner(),
      context: twoColumnSection,
      pageLayout: {
        ...bodyOwner().pageLayout,
        columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] },
      },
    };
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const acquiredAt: Array<Readonly<{ pageIndex: number; columnIndex: number }>> = [];
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input, location }) => {
          const index = input.source.path[0]!;
          if (index === 0) {
            const retained = paragraph('prior', input.source, 75);
            return {
              layout: {
                ...retained,
                flowBounds: { ...retained.flowBounds, widthPt: 80 },
                inkBounds: { ...retained.inkBounds, widthPt: 80 },
              },
              blockExtentPt: 75,
              fragmentation: { kind: 'indivisible' },
            };
          }
          acquiredAt.push({ pageIndex: location.pageIndex, columnIndex: location.columnIndex });
          const retained = paragraphWithFootnote('ordinary-note', input.source, 10, 'ordinary-note');
          return {
            layout: {
              ...retained,
              flowBounds: { ...retained.flowBounds, widthPt: 80 },
              inkBounds: { ...retained.inkBounds, widthPt: 80 },
            },
            blockExtentPt: 10,
            fragmentation: { kind: 'splittable', lineEndBoundaries: [{ segIndex: 0, charOffset: 1 }] },
          };
        },
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: ({ referenceIds }) =>
          referenceIds.includes('ordinary-note') ? 10 : 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const input: BodyLayoutInput = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: owner,
      sequence: [{
        kind: 'body-block',
        block: {
          kind: 'paragraph', source: source(0), pageBreakBefore: false,
          keepLines: false, keepNext: false, widowControl: true,
          spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
        },
      }, {
        kind: 'authored-break',
        source: source(0),
        break: 'column',
      }, {
        kind: 'body-block',
        block: {
          kind: 'paragraph', source: source(1), pageBreakBefore: false,
          keepLines: false, keepNext: false, widowControl: true,
          spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
        },
      }],
    };

    const layout = paginateBody(input, services, { currentDateMs: 0 });

    expect(layout.pages.map((page) => page.layers.body.map((node) => node.source.path[0])))
      .toEqual([[0], [1]]);
    expect(acquiredAt).toEqual([
      { pageIndex: 0, columnIndex: 1 },
      { pageIndex: 1, columnIndex: 0 },
    ]);
  });

  it('moves a table footnote past prior-column content it would invade', () => {
    const twoColumnSection: SectionLayoutContext = {
      ...section,
      columns: [{ xPt: 10, wPt: 80 }, { xPt: 110, wPt: 80 }],
    };
    const owner = {
      ...bodyOwner(),
      context: twoColumnSection,
      pageLayout: {
        ...bodyOwner().pageLayout,
        columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] },
      },
    };
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const acquiredAt: Array<Readonly<{ pageIndex: number; columnIndex: number }>> = [];
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => {
          const retained = paragraph('prior', input.source, 75);
          return {
            layout: {
              ...retained,
              flowBounds: { ...retained.flowBounds, widthPt: 80 },
              inkBounds: { ...retained.inkBounds, widthPt: 80 },
            },
            blockExtentPt: 75,
            fragmentation: { kind: 'indivisible' },
          };
        },
        measureTable: ({ input, location }) => {
          acquiredAt.push({ pageIndex: location.pageIndex, columnIndex: location.columnIndex });
          return {
            layout: tableWithFootnote('table-note', input.source, 10, 'table-note', 80),
            blockExtentPt: 10,
          };
        },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: ({ referenceIds }) =>
          referenceIds.includes('table-note') ? 10 : 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const input: BodyLayoutInput = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: owner,
      sequence: [{
        kind: 'body-block',
        block: {
          kind: 'paragraph', source: source(0), pageBreakBefore: false,
          keepLines: false, keepNext: false, widowControl: true,
          spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
        },
      }, {
        kind: 'authored-break',
        source: source(0),
        break: 'column',
      }, {
        kind: 'body-block',
        block: { kind: 'table', source: source(1) },
      }],
    };

    const layout = paginateBody(input, services, { currentDateMs: 0 });

    expect(layout.pages.map((page) => page.layers.body.map((node) => node.source.path[0])))
      .toEqual([[0], [1]]);
    expect(acquiredAt).toEqual([
      { pageIndex: 0, columnIndex: 1 },
      { pageIndex: 1, columnIndex: 0 },
    ]);
  });

  it('admits all retained frame-group references on the owner occurrence', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => {
          const index = input.source.path[0]!;
          if (index === 0) {
            return {
              layout: paragraph('prior', input.source, 75),
              blockExtentPt: 75,
              fragmentation: { kind: 'indivisible' },
            };
          }
          const ownLayout = index === 1
            ? paragraph('frame-owner', input.source, 10)
            : paragraphWithFootnote('frame-member', input.source, 10, 'later-member-note');
          return {
            layout: { ...ownLayout, ordinaryFlow: false },
            blockExtentPt: 0,
            fragmentation: { kind: 'indivisible' },
            placement: {
              coordinateSpace: 'logical-body' as const,
              xPt: 10,
              yPt: 10,
              sectionFlowOwnership: 'page' as const,
            },
            retainedFootnoteReferenceIds: index === 1 ? ['later-member-note'] : [],
          };
        },
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: ({ referenceIds }) => (
          referenceIds.includes('later-member-note') ? 10 : 0
        ),
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const input: BodyLayoutInput = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: bodyOwner(),
      sequence: [0, 1, 2].map((index) => ({
        kind: 'body-block' as const,
        block: {
          kind: 'paragraph' as const, source: source(index), pageBreakBefore: false,
          keepLines: false, keepNext: false, widowControl: true,
          spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
        },
      })),
    };

    const layout = paginateBody(input, services, { currentDateMs: 0 });

    expect(layout.pages.map((page) => page.layers.body.map((node) => node.source.path[0])))
      .toEqual([[0], [1, 2]]);
  });

  it('rejects a placed footnote reserve that cannot fit a fresh physical page', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => {
          const layout = paragraphWithFootnote('oversized-note', input.source, 10, 'too-tall');
          return {
            layout: { ...layout, ordinaryFlow: false },
            blockExtentPt: 0,
            fragmentation: { kind: 'indivisible' },
            placement: {
              coordinateSpace: 'logical-body' as const,
              xPt: 10,
              yPt: 10,
              sectionFlowOwnership: 'page' as const,
            },
          };
        },
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: ({ referenceIds }) => referenceIds.includes('too-tall') ? 81 : 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const input: BodyLayoutInput = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: bodyOwner(),
      sequence: [{
        kind: 'body-block',
        block: {
          kind: 'paragraph', source: source(0), pageBreakBefore: false,
          keepLines: false, keepNext: false, widowControl: true,
          spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
        },
      }],
    };

    expect(() => paginateBody(input, services, { currentDateMs: 0 })).toThrow(expect.objectContaining({
      code: 'FOOTNOTE_RESERVE_EXCEEDS_FRESH_PAGE',
    }));
  });

  it('rejects an ordinary paragraph footnote reserve that cannot fit a fresh physical page', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => ({
          layout: paragraphWithFootnote('oversized-note', input.source, 10, 'too-tall'),
          blockExtentPt: 10,
          fragmentation: { kind: 'splittable', lineEndBoundaries: [{ segIndex: 0, charOffset: 1 }] },
        }),
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: ({ referenceIds }) => referenceIds.includes('too-tall') ? 81 : 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const input: BodyLayoutInput = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: bodyOwner(),
      sequence: [{
        kind: 'body-block',
        block: {
          kind: 'paragraph', source: source(0), pageBreakBefore: false,
          keepLines: false, keepNext: false, widowControl: true,
          spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
        },
      }],
    };

    expect(() => paginateBody(input, services, { currentDateMs: 0 })).toThrow(expect.objectContaining({
      code: 'FOOTNOTE_RESERVE_EXCEEDS_FRESH_PAGE',
    }));
  });

  it('rejects a table fragment and footnote reserve that cannot fit a fresh physical page', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: () => { throw new Error('unused'); },
        measureTable: ({ input }) => ({
          layout: tableWithFootnote('oversized-note', input.source, 10, 'too-tall'),
          blockExtentPt: 10,
        }),
        measureStoryExtent: () => 0,
        measureFootnoteReserve: ({ referenceIds }) => referenceIds.includes('too-tall') ? 71 : 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const input: BodyLayoutInput = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: bodyOwner(),
      sequence: [{
        kind: 'body-block',
        block: { kind: 'table', source: source(0) },
      }],
    };

    expect(() => paginateBody(input, services, { currentDateMs: 0 })).toThrow(expect.objectContaining({
      code: 'FOOTNOTE_RESERVE_EXCEEDS_FRESH_PAGE',
    }));
  });

  it('keeps a paragraph with its note-bearing successor when the pair only fits a fresh page', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => {
          const index = input.source.path[0]!;
          const retained = index === 2
            ? paragraphWithFootnote('successor', input.source, 10, 'successor-note')
            : paragraph(`p${index}`, input.source, index === 0 ? 20 : 10);
          return {
            layout: retained,
            blockExtentPt: retained.advancePt,
            fragmentation: index === 2
              ? { kind: 'splittable', lineEndBoundaries: [{ segIndex: 0, charOffset: 1 }] }
              : { kind: 'indivisible' },
          };
        },
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: ({ referenceIds }) =>
          referenceIds.includes('successor-note') ? 50 : 0,
        measureFollowingBlock: ({ input }) => ({
          fullExtentPt: input.source.path[0] === 0 ? 20 : 10,
          leadContentExtentPt: input.source.path[0] === 0 ? 20 : 10,
          leadFootnoteReferenceIds: input.source.path[0] === 2 ? ['successor-note'] : [],
        }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const input: BodyLayoutInput = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: bodyOwner(),
      sequence: [0, 1, 2].map((index) => ({
        kind: 'body-block' as const,
        block: {
          kind: 'paragraph' as const, source: source(index), pageBreakBefore: false,
          keepLines: false, keepNext: index === 1, widowControl: true,
          spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
        },
      })),
    };

    const layout = paginateBody(input, services, { currentDateMs: 0 });

    expect(layout.pages.map((page) => page.layers.body.map((node) => node.source.path[0])))
      .toEqual([[0], [1, 2]]);
  });

  it('suppresses leading spacing when a keepNext unit moves to an automatic page', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const keepMeasurements: boolean[] = [];
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input, suppressSpaceBefore }) => {
          const index = input.source.path[0]!;
          if (index === 1) keepMeasurements.push(suppressSpaceBefore);
          const extent = index === 0
            ? 65
            : index === 1
              ? 36 + (suppressSpaceBefore ? 0 : 15)
              : 20;
          return {
            layout: paragraph(`p${index}`, input.source, extent),
            blockExtentPt: extent,
            fragmentation: { kind: 'indivisible' },
          };
        },
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: ({ input }) => ({
          fullExtentPt: input.source.path[0] === 2 ? 20 : 0,
          leadContentExtentPt: input.source.path[0] === 2 ? 20 : 0,
        }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const input: BodyLayoutInput = {
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: bodyOwner(),
      sequence: [0, 1, 2].map((index) => ({
        kind: 'body-block' as const,
        block: {
          kind: 'paragraph' as const, source: source(index), pageBreakBefore: false,
          keepLines: false, keepNext: index === 1, widowControl: true,
          spaceBeforePt: index === 1 ? 15 : 0,
          spaceAfterPt: 0, contextualSpacing: false, styleId: null,
        },
      })),
    };

    const layout = paginateBody(input, services, { currentDateMs: 0 });

    expect(layout.pages.map((page) => page.layers.body.map((node) => node.source.path[0])))
      .toEqual([[0], [1, 2]]);
    expect(keepMeasurements[0]).toBe(false);
    expect(keepMeasurements.at(-1)).toBe(true);
  });

  it('admits nextColumn footnote relocation against the complete next-page interval', () => {
    const twoColumnSection: SectionLayoutContext = {
      ...section,
      columns: [{ xPt: 10, wPt: 80 }, { xPt: 110, wPt: 80 }],
    };
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => {
          if (input.source.path[0] === 0) {
            return {
              layout: paragraph('lead', input.source, 30),
              blockExtentPt: 30,
              fragmentation: { kind: 'indivisible' },
            };
          }
          const layout = paragraphWithFootnote(
            'next-column-footnote',
            input.source,
            10,
            'next-column-note',
          );
          return {
            layout: {
              ...layout,
              ordinaryFlow: false,
              flowBounds: { ...layout.flowBounds, widthPt: 80 },
              inkBounds: { ...layout.inkBounds, widthPt: 80 },
            },
            blockExtentPt: 0,
            fragmentation: { kind: 'indivisible' },
            placement: {
              coordinateSpace: 'logical-body' as const,
              xPt: 10,
              yPt: 10,
              sectionFlowOwnership: 'host-flow' as const,
            },
            relocationBlockExtentPt: 65,
            retainedFootnoteReferenceIds: ['next-column-note'],
          };
        },
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: ({ referenceIds }) =>
          referenceIds.includes('next-column-note') ? 5 : 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const owner = (
      id: string,
      context: SectionLayoutContext,
      startType: 'nextPage' | 'continuous' | 'nextColumn',
    ) => ({
      sectionOccurrenceId: id,
      source: source(id === 'section:initial' ? 10 : id === 'section:mid' ? 11 : 12),
      startType,
      context,
      pageNumbering: { start: null, format: null },
      titlePage: false,
      evenAndOddHeaders: false,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      pageBordersAuthored: false,
      pageBorders: null,
      pageLayout: {
        physicalGeometry: context.geometry,
        columns: context.columns.length === 1
          ? null
          : { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] },
        textDirection: 'lrTb',
        gutterPt: 0,
        rtlGutter: false,
        mirrorMargins: false,
        gutterAtTop: false,
        bookFoldPrinting: false,
        bookFoldRevPrinting: false,
        printTwoOnOne: false,
      },
    });
    const initial = owner('section:initial', section, 'nextPage');
    const midPage = owner('section:mid', twoColumnSection, 'continuous');
    const incoming = owner('section:incoming', twoColumnSection, 'nextColumn');
    const layout = paginateBody({
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: initial,
      sequence: [
        {
          kind: 'body-block',
          block: {
            kind: 'paragraph', source: source(0), pageBreakBefore: false,
            keepLines: false, keepNext: false, widowControl: true,
            spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
          },
        },
        { kind: 'begin-section', source: source(10), section: midPage },
        { kind: 'begin-section', source: source(11), section: incoming },
        {
          kind: 'body-block',
          block: {
            kind: 'paragraph', source: source(1), pageBreakBefore: false,
            keepLines: false, keepNext: false, widowControl: true,
            spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
          },
        },
      ],
    }, services, { currentDateMs: 0 });

    expect(layout.pages).toHaveLength(2);
    expect(layout.pages[1]!.layers.body.map((node) => node.source.path[0])).toEqual([1]);
    expect(layout.pages[1]!.sectionRegions[0]!.blockStartPt).toBe(10);
    expect(layout.pages[1]!.sectionRegions[0]!.blockEndPt).toBe(90);
  });

  it('anchors a continuous section page-number restart to its shared first page', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const heights = [20, 60, 30];
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => ({
          layout: paragraph(`p${input.source.path[0]}`, input.source, heights[input.source.path[0]!]!),
          blockExtentPt: heights[input.source.path[0]!]!, fragmentation: { kind: 'indivisible' },
        }),
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const owner = (id: string, startType: 'nextPage' | 'continuous', start: number | null) => ({
      sectionOccurrenceId: id, source: source(3), startType, context: section,
      pageNumbering: { start, format: null }, titlePage: false, evenAndOddHeaders: false,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      pageBordersAuthored: false, pageBorders: null,
      pageLayout: {
        physicalGeometry: section.geometry, columns: null, textDirection: 'lrTb',
        gutterPt: 0, rtlGutter: false, mirrorMargins: false, gutterAtTop: false,
        bookFoldPrinting: false, bookFoldRevPrinting: false, printTwoOnOne: false,
      },
    });
    const block = (index: number) => ({
      kind: 'body-block' as const,
      block: {
        kind: 'paragraph' as const, source: source(index), pageBreakBefore: false,
        keepLines: false, keepNext: false, widowControl: true,
        spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
      },
    });
    const first = owner('section:0', 'nextPage', null);
    const second = owner('section:1', 'continuous', 50);
    const layout = paginateBody({
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: first,
      sequence: [block(0), { kind: 'begin-section', source: source(3), section: second }, block(1), block(2)],
    }, services, { currentDateMs: 0 });

    expect(layout.pages.map((page) => page.pageNumber.displayNumber)).toEqual([1, 51]);
  });

  it('retains an upright physical table placement independently from its logical flow charge', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => ({
          layout: paragraph('follower', input.source, 10),
          blockExtentPt: 10,
          fragmentation: { kind: 'indivisible' },
        }),
        measureTable: ({ input }) => ({
          layout: table('upright', input.source, 30), blockExtentPt: 20,
          placement: {
            coordinateSpace: 'upright-physical', xPt: 40, yPt: 25,
            sectionFlowOwnership: 'host-flow',
          },
        }),
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const owner = {
      sectionOccurrenceId: 'section:0', source: source(1), startType: 'nextPage' as const,
      context: section, pageNumbering: { start: null, format: null },
      titlePage: false, evenAndOddHeaders: false,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      pageBordersAuthored: false, pageBorders: null,
      pageLayout: {
        physicalGeometry: section.geometry, columns: null, textDirection: 'lrTb',
        gutterPt: 0, rtlGutter: false, mirrorMargins: false, gutterAtTop: false,
        bookFoldPrinting: false, bookFoldRevPrinting: false, printTwoOnOne: false,
      },
    };
    const layout = paginateBody({
      source: { story: 'body', storyInstance: 'body', path: [] }, initialSection: owner,
      sequence: [
        { kind: 'body-block', block: { kind: 'table', source: source(0) } },
        {
          kind: 'body-block',
          block: {
            kind: 'paragraph', source: source(1), pageBreakBefore: false,
            keepLines: false, keepNext: false, widowControl: true,
            spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
          },
        },
      ],
    }, services, { currentDateMs: 0 });

    const body = layout.pages[0]!.layers.body;
    expect(body[0]!.ordinaryFlow).toBe(false);
    expect(body[0]!.advancePt).toBe(30);
    expect(body[0]!.flowBounds).toMatchObject({ xPt: 40, yPt: 25, heightPt: 20 });
    expect(body[1]!.flowBounds.yPt).toBe(30);
    expect(layout.pages[0]!.layers.roots[0]).toMatchObject({ coordinateSpace: 'upright-physical' });
  });

  it('assigns deterministic distinct occurrence ids to one table continued across two columns', () => {
    const twoColumnSection: SectionLayoutContext = {
      ...section,
      columns: [{ xPt: 10, wPt: 80 }, { xPt: 110, wPt: 80 }],
    };
    const paginate = () => {
      const services = Object.freeze({
        text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
      }) as LayoutServices;
      attachBodyLayoutKernel(services, {
        openBodyLayoutSession: () => ({
          hasPaginationFields: false,
          measureParagraph: () => { throw new Error('unused'); },
          measureTable: ({ input, cursor }) => {
            const retained = table('retained-table', input.source, cursor ? 20 : 80);
            const layout = {
              ...retained,
              flowBounds: { ...retained.flowBounds, widthPt: 80 },
              inkBounds: { ...retained.inkBounds, widthPt: 80 },
              columnWidthsPt: [80],
            };
            return cursor
            ? { layout, blockExtentPt: 20 }
            : {
                layout, blockExtentPt: 80,
                nextCursor: {
                  kind: 'table',
                  cursor: { rowIndex: 1, rowFragmentIndex: 0, cells: [] },
                },
              };
          },
          measureStoryExtent: () => 0,
          measureFootnoteReserve: () => 0,
          measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
          measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
          resetPageAcquisition: () => undefined,
          moveAcquisitionCursor: () => undefined,
          flowRegistrySnapshot: emptyFlowRegistrySnapshot,
          commitFlowRegistryDelta: () => undefined,
        }),
      });
      const owner = {
        sectionOccurrenceId: 'section:two-column', source: source(1), startType: 'nextPage' as const,
        context: twoColumnSection, pageNumbering: { start: null, format: null },
        titlePage: false, evenAndOddHeaders: false,
        headers: { default: null, first: null, even: null },
        footers: { default: null, first: null, even: null },
        pageBordersAuthored: false, pageBorders: null,
        pageLayout: {
          physicalGeometry: twoColumnSection.geometry,
          columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] },
          textDirection: 'lrTb', gutterPt: 0, rtlGutter: false, mirrorMargins: false,
          gutterAtTop: false, bookFoldPrinting: false, bookFoldRevPrinting: false,
          printTwoOnOne: false,
        },
      };
      return paginateBody({
        source: { story: 'body', storyInstance: 'body', path: [] }, initialSection: owner,
        sequence: [{ kind: 'body-block', block: { kind: 'table', source: source(0) } }],
      }, services, { currentDateMs: 0 });
    };

    const layout = paginate();
    expect(layout.pages).toHaveLength(1);
    const body = layout.pages[0]!.layers.body;
    expect(body).toHaveLength(2);
    expect(body.map((node) => node.flowDomainId)).toEqual([
      layout.pages[0]!.sectionRegions[0]!.flowDomainIds[0],
      layout.pages[0]!.sectionRegions[0]!.flowDomainIds[1],
    ]);
    expect(new Set(body.map((node) => node.id)).size).toBe(2);
    expect(layout.pages[0]!.readingOrder).toEqual(body.map((node) => node.id));
    expect(paginate().pages[0]!.layers.body.map((node) => node.id))
      .toEqual(body.map((node) => node.id));
  });

  it('aligns an upright physical table by its logical allocation span', () => {
    const centeredSection: SectionLayoutContext = { ...section, verticalAlignment: 'center' };
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: () => { throw new Error('unused'); },
        measureTable: ({ input }) => ({
          layout: table('upright', input.source, 30), blockExtentPt: 20,
          placement: {
            coordinateSpace: 'upright-physical', xPt: 40, yPt: 25,
            sectionFlowOwnership: 'host-flow',
          },
        }),
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: () => ({ fullExtentPt: 0, leadContentExtentPt: 0 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const owner = {
      sectionOccurrenceId: 'section:centered', source: source(1), startType: 'nextPage' as const,
      context: centeredSection, pageNumbering: { start: null, format: null },
      titlePage: false, evenAndOddHeaders: false,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      pageBordersAuthored: false, pageBorders: null,
      pageLayout: {
        physicalGeometry: centeredSection.geometry, columns: null, textDirection: 'lrTb',
        gutterPt: 0, rtlGutter: false, mirrorMargins: false, gutterAtTop: false,
        bookFoldPrinting: false, bookFoldRevPrinting: false, printTwoOnOne: false,
      },
    };
    const layout = paginateBody({
      source: { story: 'body', storyInstance: 'body', path: [] }, initialSection: owner,
      sequence: [{ kind: 'body-block', block: { kind: 'table', source: source(0) } }],
    }, services, { currentDateMs: 0 });
    const retained = layout.pages[0]!.layers.body[0]!;

    expect(retained.advancePt).toBe(30);
    expect(retained.flowBounds.yPt).toBe(55);
    expect(layout.pages[0]!.layers.roots[0]).toEqual({
      layer: 'body', node: retained, coordinateSpace: 'upright-physical',
    });
    expect(layout.pages[0]!.layers.roots[0]!.node).toBe(retained);
  });

  it('moves a feasible keepNext chain together before accepting any chain member', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const heights = [30, 20, 20, 20];
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => ({
          layout: paragraph(`p${input.source.path[0]}`, input.source, heights[input.source.path[0]!]!),
          blockExtentPt: heights[input.source.path[0]!]!, fragmentation: { kind: 'indivisible' },
        }),
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: ({ input }) => ({
          fullExtentPt: heights[input.source.path[0]!]!,
          leadContentExtentPt: heights[input.source.path[0]!]!,
        }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const owner = {
      sectionOccurrenceId: 'section:0', source: source(3), startType: 'nextPage' as const,
      context: section, pageNumbering: { start: null, format: null },
      titlePage: false, evenAndOddHeaders: false,
      headers: { default: null, first: null, even: null }, footers: { default: null, first: null, even: null },
      pageBordersAuthored: false, pageBorders: null,
      pageLayout: {
        physicalGeometry: section.geometry, columns: null, textDirection: 'lrTb',
        gutterPt: 0, rtlGutter: false, mirrorMargins: false, gutterAtTop: false,
        bookFoldPrinting: false, bookFoldRevPrinting: false, printTwoOnOne: false,
      },
    };
    const layout = paginateBody({
      source: { story: 'body', storyInstance: 'body', path: [] }, initialSection: owner,
      sequence: heights.map((_height, index) => ({
        kind: 'body-block' as const,
        block: {
          kind: 'paragraph' as const, source: source(index), pageBreakBefore: false,
          keepLines: false, keepNext: index === 1 || index === 2, widowControl: true,
          spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
        },
      })),
    }, services, { currentDateMs: 0 });

    expect(layout.pages.map((page) => page.layers.body.map((node) => node.source.path[0])))
      .toEqual([[0], [1, 2, 3]]);
  });

  it('bridges an undecorated empty keepNext mark through the following paragraph', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const heights = [30, 20, 20, 20];
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => {
          const index = input.source.path[0]!;
          const layout = paragraph(`p${index}`, input.source, heights[index]!);
          return {
            layout: index === 1
              ? {
                  ...layout,
                  inkBounds: { ...layout.inkBounds, widthPt: 0 },
                  paragraphMark: { hidden: false, bounds: { ...layout.inkBounds, widthPt: 0 } },
                }
              : layout,
            blockExtentPt: heights[index]!, fragmentation: { kind: 'indivisible' },
          };
        },
        measureTable: () => { throw new Error('unused'); },
        measureFollowingBlock: ({ input }) => ({
          fullExtentPt: heights[input.source.path[0]!]!,
          leadContentExtentPt: heights[input.source.path[0]!]!,
        }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const owner = {
      sectionOccurrenceId: 'section:0', source: source(3), startType: 'nextPage' as const,
      context: section, pageNumbering: { start: null, format: null },
      titlePage: false, evenAndOddHeaders: false,
      headers: { default: null, first: null, even: null }, footers: { default: null, first: null, even: null },
      pageBordersAuthored: false, pageBorders: null,
      pageLayout: {
        physicalGeometry: section.geometry, columns: null, textDirection: 'lrTb',
        gutterPt: 0, rtlGutter: false, mirrorMargins: false, gutterAtTop: false,
        bookFoldPrinting: false, bookFoldRevPrinting: false, printTwoOnOne: false,
      },
    };
    const layout = paginateBody({
      source: { story: 'body', storyInstance: 'body', path: [] }, initialSection: owner,
      sequence: heights.map((_height, index) => ({
        kind: 'body-block' as const,
        block: {
          kind: 'paragraph' as const, source: source(index), pageBreakBefore: false,
          keepLines: false, keepNext: index === 1, widowControl: true,
          spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
          inkless: index === 1,
        },
      })),
    }, services, { currentDateMs: 0 });

    expect(layout.pages.map((page) => page.layers.body.map((node) => node.source.path[0])))
      .toEqual([[0], [1, 2, 3]]);
  });

  it('balances the outgoing multi-column section before an incoming continuous section', () => {
    const balancedSection: SectionLayoutContext = {
      ...section,
      columns: [{ xPt: 10, wPt: 80 }, { xPt: 110, wPt: 80 }],
    };
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => {
          const layout = paragraph(`p${input.source.path[0]}`, input.source, 20);
          return {
            layout: {
              ...layout,
              flowBounds: { ...layout.flowBounds, widthPt: 80 },
              inkBounds: { ...layout.inkBounds, widthPt: 80 },
            },
            blockExtentPt: 20, fragmentation: { kind: 'indivisible' },
          };
        },
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: () => ({ fullExtentPt: 20, leadContentExtentPt: 20 }),
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => undefined,
      }),
    });
    const owner = (
      id: string,
      context: SectionLayoutContext,
      startType: 'continuous' | 'nextPage',
    ) => ({
      sectionOccurrenceId: id, source: source(7), startType, context,
      pageNumbering: { start: null, format: null }, titlePage: false, evenAndOddHeaders: false,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      pageBordersAuthored: false, pageBorders: null,
      pageLayout: {
        physicalGeometry: context.geometry, columns: context === balancedSection
          ? { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] }
          : null,
        textDirection: 'lrTb', gutterPt: 0, rtlGutter: false, mirrorMargins: false,
        gutterAtTop: false, bookFoldPrinting: false, bookFoldRevPrinting: false,
        printTwoOnOne: false,
      },
    });
    const first = owner('section:balanced', balancedSection, 'nextPage');
    const final = owner('section:final', section, 'continuous');
    const blocks = Array.from({ length: 6 }, (_, index) => ({
      kind: 'body-block' as const,
      block: {
        kind: 'paragraph' as const, source: source(index), pageBreakBefore: false,
        keepLines: false, keepNext: false, widowControl: true,
        spaceBeforePt: 0, spaceAfterPt: 0, contextualSpacing: false, styleId: null,
      },
    }));
    const layout = paginateBody({
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: first,
      sequence: [...blocks, { kind: 'begin-section', source: source(6), section: final }],
    }, services, { currentDateMs: 0 });
    const balancedNodes = layout.pages[0]!.layers.body;

    expect(balancedNodes.map((node) => node.flowBounds.xPt)).toEqual([10, 10, 10, 110, 110, 110]);
  });

  it.each([16, 64, 256])(
    'uses a constant number of full-document passes for %i balance boundaries',
    (boundaryCount) => {
      const contentHeightPt = boundaryCount + 20;
      const balancedSection: SectionLayoutContext = {
        ...section,
        geometry: {
          ...section.geometry,
          pageHeight: contentHeightPt + 20,
        },
        columns: [{ xPt: 10, wPt: 80 }, { xPt: 110, wPt: 80 }],
      };
      let openedSessions = 0;
      const services = Object.freeze({
        text: { fingerprint: 'text' },
        images: { fingerprint: 'images' },
        math: { fingerprint: 'math' },
      }) as LayoutServices;
      attachBodyLayoutKernel(services, {
        openBodyLayoutSession: () => {
          openedSessions += 1;
          return {
            hasPaginationFields: false,
            measureParagraph: ({ input }) => {
              const layout = paragraph(`p${input.source.path[0]}`, input.source, 1);
              return {
                layout: {
                  ...layout,
                  flowBounds: { ...layout.flowBounds, widthPt: 80 },
                  inkBounds: { ...layout.inkBounds, widthPt: 80 },
                },
                blockExtentPt: 1,
                fragmentation: { kind: 'indivisible' },
              };
            },
            measureTable: () => { throw new Error('unused'); },
            measureStoryExtent: () => 0,
            measureFootnoteReserve: () => 0,
            measureFollowingBlock: () => ({ fullExtentPt: 1, leadContentExtentPt: 1 }),
            measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
            resetPageAcquisition: () => undefined,
            moveAcquisitionCursor: () => undefined,
            flowRegistrySnapshot: emptyFlowRegistrySnapshot,
            commitFlowRegistryDelta: () => undefined,
          };
        },
      });
      const owner = (
        id: string,
        context: SectionLayoutContext,
        startType: 'continuous' | 'nextPage',
      ) => ({
        sectionOccurrenceId: id,
        source: source(boundaryCount + 1),
        startType,
        context,
        pageNumbering: { start: null, format: null },
        titlePage: false,
        evenAndOddHeaders: false,
        headers: { default: null, first: null, even: null },
        footers: { default: null, first: null, even: null },
        pageBordersAuthored: false,
        pageBorders: null,
        pageLayout: {
          physicalGeometry: context.geometry,
          columns: context === balancedSection
            ? { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] }
            : null,
          textDirection: 'lrTb' as const,
          gutterPt: 0,
          rtlGutter: false,
          mirrorMargins: false,
          gutterAtTop: false,
          bookFoldPrinting: false,
          bookFoldRevPrinting: false,
          printTwoOnOne: false,
        },
      });
      const incomingContext: SectionLayoutContext = {
        ...section,
        geometry: balancedSection.geometry,
      };
      const outgoing = owner('section:balanced', balancedSection, 'nextPage');
      const incoming = owner('section:final', incomingContext, 'continuous');
      const blocks = Array.from({ length: boundaryCount }, (_, index) => ({
        kind: 'body-block' as const,
        block: {
          kind: 'paragraph' as const,
          source: source(index),
          pageBreakBefore: false,
          keepLines: false,
          keepNext: false,
          widowControl: true,
          spaceBeforePt: 0,
          spaceAfterPt: 0,
          contextualSpacing: false,
          styleId: null,
        },
      }));

      const layout = paginateBody({
        source: { story: 'body', storyInstance: 'body', path: [] },
        initialSection: outgoing,
        sequence: [
          ...blocks,
          { kind: 'begin-section' as const, source: source(boundaryCount), section: incoming },
        ],
      }, services, { currentDateMs: 0 });

      expect(openedSessions).toBe(2);
      expect(layout.pages[0]!.layers.body.filter((node) => (
        node.flowDomainId === layout.pages[0]!.sectionRegions[0]!.flowDomainIds[0]
      ))).toHaveLength(Math.ceil(boundaryCount / 2));
    },
  );

  it('commits page-owned anchor exclusions before measuring earlier page content', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const events: string[] = [];
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => ({
        hasPaginationFields: false,
        measureParagraph: ({ input }) => {
          events.push(`measure:${input.source.path[0]}`);
          return {
            layout: paragraph(`p${input.source.path[0]}`, input.source, 20),
            blockExtentPt: 20, fragmentation: { kind: 'indivisible' },
          };
        },
        measureTable: () => { throw new Error('unused'); },
        measureStoryExtent: () => 0,
        measureFootnoteReserve: () => 0,
        measureFollowingBlock: () => ({ fullExtentPt: 20, leadContentExtentPt: 20 }),
        prescanPageAnchors: ({ anchors, location }) => {
          events.push(`prescan:${anchors.map((anchor) => anchor.paragraphSource.path[0]).join(',')}`);
          return {
            floats: {
              coordinateSpace: 'logical-page-points', flowDomainId: location.flowDomainId,
              baseEntries: [],
              baseNextParagraphId: 0, nextParagraphId: 1,
              entries: [{
                kind: 'shape', occurrenceId: anchors[0]!.occurrenceId, paragraphId: 0,
                bounds: { xPt: 100, yPt: 10, widthPt: 20, heightPt: 20 },
                exclusionBounds: { xPt: 100, yPt: 10, widthPt: 20, heightPt: 20 },
              }],
            },
          };
        },
        measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
        resetPageAcquisition: () => undefined,
        moveAcquisitionCursor: () => undefined,
        flowRegistrySnapshot: emptyFlowRegistrySnapshot,
        commitFlowRegistryDelta: () => { events.push('commit'); },
      }),
    });
    const owner = {
      sectionOccurrenceId: 'section:0', source: source(2), startType: 'nextPage' as const,
      context: section, pageNumbering: { start: null, format: null },
      titlePage: false, evenAndOddHeaders: false,
      headers: { default: null, first: null, even: null }, footers: { default: null, first: null, even: null },
      pageBordersAuthored: false, pageBorders: null,
      pageLayout: {
        physicalGeometry: section.geometry, columns: null, textDirection: 'lrTb', gutterPt: 0,
        rtlGutter: false, mirrorMargins: false, gutterAtTop: false, bookFoldPrinting: false,
        bookFoldRevPrinting: false, printTwoOnOne: false,
      },
    };
    const block = (index: number, anchors: readonly string[] = []) => ({
      kind: 'body-block' as const,
      block: {
        kind: 'paragraph' as const, source: source(index), pageBreakBefore: false,
        keepLines: false, keepNext: false, widowControl: true, spaceBeforePt: 0, spaceAfterPt: 0,
        contextualSpacing: false, styleId: null,
        ...(anchors.length === 0 ? {} : { pageOwnedAnchorOccurrenceIds: anchors }),
      },
    });

    paginateBody({
      source: { story: 'body', storyInstance: 'body', path: [] }, initialSection: owner,
      sequence: [block(0), block(1, ['anchor:1'])],
    }, services, { currentDateMs: 0 });

    expect(events.slice(0, 3)).toEqual(['prescan:1', 'commit', 'measure:0']);
  });

  it('prescans incoming nextColumn anchors before measuring earlier content in that flow domain', () => {
    const columns = [{ xPt: 10, wPt: 80 }, { xPt: 110, wPt: 80 }];
    const twoColumnSection: SectionLayoutContext = { ...section, columns };
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const commitCounts: number[] = [];
    const measuredWithIncomingAuthority: boolean[] = [];
    attachBodyLayoutKernel(services, {
      openBodyLayoutSession: () => {
        const sessionIndex = commitCounts.push(0) - 1;
        return {
          hasPaginationFields: false,
          measureParagraph: ({ input }) => {
            const index = input.source.path[0]!;
            const base = paragraph(`p${index}`, input.source, 20);
            const sized = {
              ...base,
              flowBounds: { ...base.flowBounds, widthPt: 80 },
              inkBounds: { ...base.inkBounds, widthPt: 80 },
            };
            if (index === 1) {
              const authoritative = commitCounts[sessionIndex] === 1;
              measuredWithIncomingAuthority.push(authoritative);
              return {
                layout: {
                  ...sized,
                  exclusions: authoritative ? [{
                    id: 'incoming-wrap',
                    wrap: 'square' as const,
                    bounds: { xPt: 120, yPt: 10, widthPt: 20, heightPt: 20 },
                    polygon: [{ xPt: 120, yPt: 10 }],
                    anchorOccurrenceId: 'anchor:incoming',
                  }] : [],
                },
                blockExtentPt: 20,
                fragmentation: { kind: 'indivisible' },
              };
            }
            if (index === 2) {
              const drawing = {
                kind: 'drawing' as const,
                id: 'incoming-anchor',
                source: input.source,
                flowDomainId: 'acquisition',
                flowBounds: { xPt: 120, yPt: 10, widthPt: 20, heightPt: 20 },
                inkBounds: { xPt: 120, yPt: 10, widthPt: 20, heightPt: 20 },
                advancePt: 0,
                ordinaryFlow: false,
                commands: [],
                anchorLayer: {
                  occurrenceId: 'anchor:incoming',
                  acquisitionOccurrenceId: 'anchor:incoming',
                  behindDoc: false,
                  relativeHeight: 0,
                  sourceOrder: 0,
                  horizontalOwnership: 'page' as const,
                  verticalOwnership: 'page' as const,
                },
              };
              return {
                layout: { ...sized, drawings: [drawing] },
                blockExtentPt: 20,
                fragmentation: { kind: 'indivisible' },
              };
            }
            return { layout: sized, blockExtentPt: 20, fragmentation: { kind: 'indivisible' } };
          },
          measureTable: () => { throw new Error('unused'); },
          measureStoryExtent: () => 0,
          measureFootnoteReserve: () => 0,
          measureFollowingBlock: () => ({ fullExtentPt: 20, leadContentExtentPt: 20 }),
          prescanPageAnchors: ({ anchors, location }) => ({
            floats: {
              coordinateSpace: 'logical-page-points',
              flowDomainId: location.flowDomainId,
              baseEntries: [],
              baseNextParagraphId: 0,
              nextParagraphId: 1,
              entries: [{
                kind: 'shape',
                occurrenceId: anchors[0]!.occurrenceId,
                paragraphId: 0,
                bounds: { xPt: 120, yPt: 10, widthPt: 20, heightPt: 20 },
                exclusionBounds: { xPt: 120, yPt: 10, widthPt: 20, heightPt: 20 },
              }],
            },
          }),
          measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
          resetPageAcquisition: () => undefined,
          moveAcquisitionCursor: () => undefined,
          flowRegistrySnapshot: emptyFlowRegistrySnapshot,
          commitFlowRegistryDelta: () => { commitCounts[sessionIndex]! += 1; },
        };
      },
    });
    const owner = (id: string, startType: 'nextPage' | 'nextColumn') => ({
      sectionOccurrenceId: id,
      source: source(id === 'section:outgoing' ? 10 : 11),
      startType,
      context: twoColumnSection,
      pageNumbering: { start: null, format: null },
      titlePage: false,
      evenAndOddHeaders: false,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      pageBordersAuthored: false,
      pageBorders: null,
      pageLayout: {
        physicalGeometry: twoColumnSection.geometry,
        columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] },
        textDirection: 'lrTb',
        gutterPt: 0,
        rtlGutter: false,
        mirrorMargins: false,
        gutterAtTop: false,
        bookFoldPrinting: false,
        bookFoldRevPrinting: false,
        printTwoOnOne: false,
      },
    });
    const block = (index: number, anchors: readonly string[] = []) => ({
      kind: 'body-block' as const,
      block: {
        kind: 'paragraph' as const,
        source: source(index),
        pageBreakBefore: false,
        keepLines: false,
        keepNext: false,
        widowControl: true,
        spaceBeforePt: 0,
        spaceAfterPt: 0,
        contextualSpacing: false,
        styleId: null,
        ...(anchors.length === 0 ? {} : { pageOwnedAnchorOccurrenceIds: anchors }),
      },
    });
    const layout = paginateBody({
      source: { story: 'body', storyInstance: 'body', path: [] },
      initialSection: owner('section:outgoing', 'nextPage'),
      sequence: [
        block(0),
        {
          kind: 'begin-section',
          source: source(9),
          section: owner('section:incoming', 'nextColumn'),
        },
        block(1),
        block(2, ['anchor:incoming']),
      ],
    }, services, { currentDateMs: 0 });
    const incomingLead = layout.pages[0]!.layers.body.find((node) => node.source.path[0] === 1);

    expect(measuredWithIncomingAuthority.length).toBeGreaterThan(0);
    expect(measuredWithIncomingAuthority.every(Boolean)).toBe(true);
    expect(incomingLead?.kind === 'paragraph' ? incomingLead.exclusions : []).toHaveLength(1);
    expect(commitCounts.length).toBeGreaterThan(0);
    expect(commitCounts.every((count) => count === 1)).toBe(true);
  });
});

import { describe, expect, it } from 'vitest';
import type { SectionLayoutContext } from '../layout-context.js';
import type {
  BodyLayoutInput,
  BodySectionLayoutInput,
  BodyStoryReferenceSet,
} from './body-layout-input.js';
import type { SyntheticBodyLayoutKernel } from './body-layout-kernel.test-support.js';
import { attachSyntheticBodyLayoutKernel } from './body-layout-kernel.test-support.js';
import { paginateBody } from './body-paginator.js';
import { LayoutVariantStore } from './variant-store.js';
import {
  fieldAcquisitionContextOf,
} from './runtime-state.js';
import type {
  LayoutOptions,
} from './options.js';
import type {
  LayoutServices,
  ParagraphLayout,
  SourceRef,
} from './types.js';

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

const bodySource = (index: number): SourceRef => ({
  story: 'body', storyInstance: 'body', path: [index],
});

const storySource = (storyInstance: string): SourceRef => ({
  story: 'header', storyInstance, path: [],
});

const noStories = (): BodyStoryReferenceSet => ({
  default: null, first: null, even: null,
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

function owner(
  sectionOccurrenceId: string,
  overrides: Partial<BodySectionLayoutInput> = {},
): BodySectionLayoutInput {
  return {
    sectionOccurrenceId,
    source: bodySource(100),
    startType: 'nextPage',
    context: section,
    pageNumbering: { start: null, format: null },
    titlePage: false,
    evenAndOddHeaders: false,
    headers: noStories(),
    footers: noStories(),
    pageBordersAuthored: false,
    pageBorders: null,
    pageLayout: {
      physicalGeometry: section.geometry,
      columns: null,
      textDirection: 'lrTb',
      gutterPt: 0,
      rtlGutter: false,
      mirrorMargins: false,
      gutterAtTop: false,
      bookFoldPrinting: false,
      bookFoldRevPrinting: false,
      printTwoOnOne: false,
    },
    ...overrides,
  };
}

const bodyBlock = (index: number) => ({
  kind: 'body-block' as const,
  block: {
    kind: 'paragraph' as const,
    source: bodySource(index),
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

function paragraph(index: number, heightPt: number): ParagraphLayout {
  return {
    kind: 'paragraph',
    id: `paragraph:${index}`,
    source: bodySource(index),
    flowDomainId: 'acquisition',
    flowBounds: { xPt: 0, yPt: 0, widthPt: 180, heightPt },
    inkBounds: { xPt: 0, yPt: 0, widthPt: 180, heightPt },
    advancePt: heightPt,
    ordinaryFlow: true,
    spacing: { beforePt: 0, afterPt: 0 },
    contextualSpacing: false,
    lines: [],
    borders: [],
    resources: [],
    drawings: [],
    textBoxes: [],
    events: [],
    exclusions: [],
  };
}

interface PageFieldObservation {
  readonly sourceIndex: number;
  readonly physicalPageIndex: number;
  readonly totalPages: number;
  readonly displayPageNumber: number | undefined;
  readonly pageNumberFormat: string | undefined;
}

interface StoryObservation {
  readonly storyInstance: string;
  readonly physicalPageIndex: number;
}

function createServices(input: Readonly<{
  heights: ReadonlyMap<number, number>;
  hasPaginationFields?: boolean;
  fieldObservations?: PageFieldObservation[];
  storyObservations?: StoryObservation[];
}>): LayoutServices {
  const services = Object.freeze({
    text: { fingerprint: 'text' },
    images: { fingerprint: 'images' },
    math: { fingerprint: 'math' },
  }) as LayoutServices;
  const kernel: SyntheticBodyLayoutKernel = {
    openBodyLayoutSession: (_sessionInput, iterationServices) => ({
      hasPaginationFields: input.hasPaginationFields ?? false,
      measureParagraph: ({ input: paragraphInput, location }) => {
        const sourceIndex = paragraphInput.source.path[0]!;
        const context = fieldAcquisitionContextOf(iterationServices);
        const page = context.resolveDestinationPage?.(location.pageIndex);
        input.fieldObservations?.push({
          sourceIndex,
          physicalPageIndex: location.pageIndex,
          totalPages: context.totalPages,
          displayPageNumber: page?.displayPageNumber,
          pageNumberFormat: page?.pageNumberFormat,
        });
        const heightPt = input.heights.get(sourceIndex) ?? 10;
        return {
          layout: paragraph(sourceIndex, heightPt),
          blockExtentPt: heightPt,
          fragmentation: { kind: 'indivisible' },
        };
      },
      measureTable: () => { throw new Error('unused'); },
      measureStoryExtent: ({ source, pageIndex }) => {
        input.storyObservations?.push({
          storyInstance: source.storyInstance,
          physicalPageIndex: pageIndex,
        });
        return 0;
      },
      measureFootnoteReserve: () => 0,
      measureFollowingBlock: ({ input: following }) => {
        const heightPt = input.heights.get(following.source.path[0]!) ?? 10;
        return { fullExtentPt: heightPt, leadContentExtentPt: heightPt };
      },
      measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
      resetPageAcquisition: () => undefined,
      moveAcquisitionCursor: () => undefined,
      flowRegistrySnapshot: emptyFlowRegistrySnapshot,
      commitFlowRegistryDelta: () => undefined,
    }),
  };
  attachSyntheticBodyLayoutKernel(services, kernel);
  return services;
}

function paginate(
  initialSection: BodySectionLayoutInput,
  sequence: BodyLayoutInput['sequence'],
  heights: ReadonlyMap<number, number>,
  extras: Readonly<{
    hasPaginationFields?: boolean;
    fieldObservations?: PageFieldObservation[];
    storyObservations?: StoryObservation[];
  }> = {},
) {
  const services = createServices({ heights, ...extras });
  return paginateBody({
    source: { story: 'body', storyInstance: 'body', path: [] },
    initialSection,
    sequence,
  }, services, { currentDateMs: 0 });
}

describe('canonical physical-page numbering', () => {
  it.each([
    {
      display: 'firstPage',
      parity: 'odd' as const,
      coverBlockCount: 1,
      blankVisible: true,
      followingVisible: false,
    },
    {
      display: 'notFirstPage',
      parity: 'odd' as const,
      coverBlockCount: 1,
      blankVisible: false,
      followingVisible: true,
    },
    {
      display: 'firstPage',
      parity: 'even' as const,
      coverBlockCount: 2,
      blankVisible: true,
      followingVisible: false,
    },
    {
      display: 'notFirstPage',
      parity: 'even' as const,
      coverBlockCount: 2,
      blankVisible: false,
      followingVisible: true,
    },
  ] as const)(
    'treats a forced $parity parity blank as the first page owned by a continuous section for $display',
    ({ display, parity, coverBlockCount, blankVisible, followingVisible }) => {
      const cover = owner('section:cover');
      const continued = owner('section:continued', {
        startType: 'continuous',
        pageBordersAuthored: true,
        pageBorders: {
          offsetFrom: 'page',
          display,
          zOrder: 'front',
          top: { style: 'single', width: 1, space: 4 },
        },
      });
      const coverBlocks = Array.from(
        { length: coverBlockCount },
        (_, index) => bodyBlock(index),
      );
      const continuedBlockIndex = coverBlockCount;
      const layout = paginate(cover, [
        ...coverBlocks,
        { kind: 'begin-section', source: bodySource(10), section: continued },
        {
          kind: 'authored-break',
          source: bodySource(11),
          break: 'page',
          parity,
        },
        bodyBlock(continuedBlockIndex),
      ], new Map([
        ...coverBlocks.map((_, index) => [index, coverBlockCount === 1 ? 20 : 60] as const),
        [continuedBlockIndex, 20],
      ]));

      expect(layout.pages.slice(-2).map((page) => ({
        owner: page.sectionOccurrenceId,
        parityBlank: page.parityBlank,
        border: page.pageBorder !== null,
      }))).toEqual([
        { owner: 'section:continued', parityBlank: true, border: blankVisible },
        { owner: 'section:continued', parityBlank: false, border: followingVisible },
      ]);
      expect(layout.pages.at(-3)?.sectionOccurrenceId).toBe('section:cover');
    },
  );

  it('numbers physical pages monotonically when no section declares pgNumType', () => {
    const layout = paginate(
      owner('section:default'),
      [bodyBlock(0), bodyBlock(1), bodyBlock(2)],
      new Map([[0, 60], [1, 60], [2, 60]]),
    );

    expect(layout.pages.map((page) => page.pageNumber)).toEqual([
      { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:default' },
      { displayNumber: 2, format: 'decimal', sectionOccurrenceId: 'section:default' },
      { displayNumber: 3, format: 'decimal', sectionOccurrenceId: 'section:default' },
    ]);
  });

  it.each([0, 1, 25])('starts the first section at %i and increments physical pages', (start) => {
    const layout = paginate(
      owner('section:numbered', {
        pageNumbering: { start, format: null },
      }),
      [bodyBlock(0), bodyBlock(1), bodyBlock(2)],
      new Map([[0, 60], [1, 60], [2, 60]]),
    );

    expect(layout.pages.map((page) => page.pageNumber.displayNumber))
      .toEqual([start, start + 1, start + 2]);
  });

  it('continues numbering when an incoming section omits w:start', () => {
    const front = owner('section:front', {
      pageNumbering: { start: 5, format: null },
    });
    const body = owner('section:body');
    const layout = paginate(front, [
      bodyBlock(0),
      bodyBlock(1),
      { kind: 'begin-section', source: bodySource(10), section: body },
      bodyBlock(2),
      bodyBlock(3),
    ], new Map([[0, 60], [1, 60], [2, 60], [3, 60]]));

    expect(layout.pages.map((page) => page.pageNumber.displayNumber))
      .toEqual([5, 6, 7, 8]);
  });

  it('defaults an omitted incoming format to decimal instead of inheriting it', () => {
    const front = owner('section:front', {
      pageNumbering: { start: 1, format: 'lowerRoman' },
    });
    const body = owner('section:body', {
      pageNumbering: { start: 1, format: null },
    });
    const layout = paginate(front, [
      bodyBlock(0),
      { kind: 'begin-section', source: bodySource(10), section: body },
      bodyBlock(1),
    ], new Map([[0, 20], [1, 20]]));

    expect(layout.pages.map((page) => page.pageNumber.format))
      .toEqual(['lowerRoman', 'decimal']);
  });

  it('restarts and reformats numbering on the nextPage owned by the incoming section', () => {
    const front = owner('section:front');
    const body = owner('section:body', {
      pageNumbering: { start: 5, format: 'lowerRoman' },
    });
    const layout = paginate(front, [
      bodyBlock(0),
      bodyBlock(1),
      { kind: 'begin-section', source: bodySource(10), section: body },
      bodyBlock(2),
      bodyBlock(3),
    ], new Map([[0, 60], [1, 60], [2, 60], [3, 60]]));

    expect(layout.pages.map(({ pageNumber }) => [
      pageNumber.displayNumber,
      pageNumber.format,
      pageNumber.sectionOccurrenceId,
    ])).toEqual([
      [1, 'decimal', 'section:front'],
      [2, 'decimal', 'section:front'],
      [5, 'lowerRoman', 'section:body'],
      [6, 'lowerRoman', 'section:body'],
    ]);
  });

  it('keeps shared-page display ownership at the page top and anchors a continuous restart there', () => {
    const outgoing = owner('section:outgoing');
    const incoming = owner('section:incoming', {
      startType: 'continuous',
      pageNumbering: { start: 50, format: 'upperRoman' },
    });
    const layout = paginate(outgoing, [
      bodyBlock(0),
      { kind: 'begin-section', source: bodySource(10), section: incoming },
      bodyBlock(1),
      bodyBlock(2),
    ], new Map([[0, 20], [1, 60], [2, 30]]));

    expect(layout.pages[0]!.sectionRegions.map((region) => region.sectionOccurrenceId))
      .toEqual(['section:outgoing', 'section:incoming']);
    expect(layout.pages.map((page) => page.sectionOccurrenceId))
      .toEqual(['section:outgoing', 'section:incoming']);
    expect(layout.pages.map((page) => page.pageNumber)).toEqual([
      { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:outgoing' },
      { displayNumber: 51, format: 'upperRoman', sectionOccurrenceId: 'section:incoming' },
    ]);
  });

  it('does not consume a continuous restart on an exhausted shared page', () => {
    const outgoing = owner('section:outgoing');
    const incoming = owner('section:incoming', {
      startType: 'continuous',
      pageNumbering: { start: 2, format: null },
    });
    const layout = paginate(outgoing, [
      bodyBlock(0),
      { kind: 'begin-section', source: bodySource(10), section: incoming },
      bodyBlock(1),
    ], new Map([[0, 80], [1, 20]]));

    expect(layout.pages[0]!.sectionRegions.map((region) => ({
      owner: region.sectionOccurrenceId,
      height: region.blockEndPt - region.blockStartPt,
    }))).toEqual([
      { owner: 'section:outgoing', height: 80 },
      { owner: 'section:incoming', height: 0 },
    ]);
    expect(layout.pages.map((page) => page.pageNumber.displayNumber)).toEqual([1, 2]);
  });

  it('does not consume a continuous restart on an empty shared region before pageBreakBefore', () => {
    const outgoing = owner('section:outgoing');
    const incoming = owner('section:incoming', {
      startType: 'continuous',
      pageNumbering: { start: 2, format: null },
    });
    const forced = bodyBlock(1);
    const layout = paginate(outgoing, [
      bodyBlock(0),
      { kind: 'begin-section', source: bodySource(10), section: incoming },
      { ...forced, block: { ...forced.block, pageBreakBefore: true } },
    ], new Map([[0, 20], [1, 20]]));

    expect(layout.pages[0]!.sectionRegions.map((region) => ({
      owner: region.sectionOccurrenceId,
      height: region.blockEndPt - region.blockStartPt,
    }))).toEqual([
      { owner: 'section:outgoing', height: 20 },
      { owner: 'section:incoming', height: 60 },
    ]);
    expect(layout.pages.map((page) => page.pageNumber.displayNumber)).toEqual([1, 2]);
  });

  it.each([
    {
      startType: 'oddPage' as const,
      outgoingBlocks: [bodyBlock(0)],
      heights: new Map([[0, 20], [2, 20]]),
      expectedNumbers: [10, 11, 1],
      blankIndex: 1,
      incomingIndex: 2,
    },
    {
      startType: 'evenPage' as const,
      outgoingBlocks: [bodyBlock(0), bodyBlock(1)],
      heights: new Map([[0, 60], [1, 60], [2, 20]]),
      expectedNumbers: [10, 11, 12, 1],
      blankIndex: 2,
      incomingIndex: 3,
    },
  ])('keeps the inserted $startType parity blank in the outgoing numbering series', ({
    startType,
    outgoingBlocks,
    heights,
    expectedNumbers,
    blankIndex,
    incomingIndex,
  }) => {
    const outgoing = owner('section:outgoing', {
      pageNumbering: { start: 10, format: 'upperRoman' },
    });
    const incoming = owner('section:incoming', {
      startType,
      pageNumbering: { start: 1, format: 'lowerRoman' },
    });
    const layout = paginate(outgoing, [
      ...outgoingBlocks,
      { kind: 'begin-section', source: bodySource(10), section: incoming },
      bodyBlock(2),
    ], heights);

    expect(layout.pages.map((page) => page.pageNumber.displayNumber)).toEqual(expectedNumbers);
    expect(layout.pages[blankIndex]).toMatchObject({
      parityBlank: true,
      sectionOccurrenceId: 'section:outgoing',
      pageNumber: {
        sectionOccurrenceId: 'section:outgoing',
        format: 'upperRoman',
      },
      sectionRegions: [],
    });
    expect(layout.pages[incomingIndex]).toMatchObject({
      parityBlank: false,
      sectionOccurrenceId: 'section:incoming',
      pageNumber: {
        displayNumber: 1,
        sectionOccurrenceId: 'section:incoming',
        format: 'lowerRoman',
      },
    });
  });

  it('converges PAGE and NUMPAGES acquisition against final canonical page metadata', () => {
    const observations: PageFieldObservation[] = [];
    const layout = paginate(
      owner('section:fields'),
      [bodyBlock(0), bodyBlock(1)],
      new Map([[0, 60], [1, 60]]),
      { hasPaginationFields: true, fieldObservations: observations },
    );
    const finalPass = observations.filter(({ totalPages }) => totalPages === layout.pages.length);

    expect(layout.pages.map((page) => page.pageNumber.displayNumber)).toEqual([1, 2]);
    expect(finalPass).toEqual(expect.arrayContaining([
      expect.objectContaining({
        physicalPageIndex: 0,
        totalPages: 2,
        displayPageNumber: 1,
        pageNumberFormat: 'decimal',
      }),
      expect.objectContaining({
        physicalPageIndex: 1,
        totalPages: 2,
        displayPageNumber: 2,
        pageNumberFormat: 'decimal',
      }),
    ]));
  });

  it('retains an authored parity blank in physical order and NUMPAGES structure', () => {
    const observations: PageFieldObservation[] = [];
    const numberedOwner = owner('section:authored-parity', {
      pageNumbering: { start: 50, format: 'decimal' },
    });
    const layout = paginate(numberedOwner, [
      bodyBlock(0),
      {
        kind: 'authored-break',
        source: bodySource(1),
        break: 'page',
        parity: 'odd',
      },
      bodyBlock(2),
    ], new Map([[0, 20], [2, 20]]), {
      hasPaginationFields: true,
      fieldObservations: observations,
    });
    const finalPass = observations.filter(({ totalPages }) => totalPages === layout.pages.length);

    expect(layout.pages.map((page) => ({
      pageIndex: page.pageIndex,
      displayNumber: page.pageNumber.displayNumber,
      owner: page.sectionOccurrenceId,
      parityBlank: page.parityBlank,
    }))).toEqual([
      { pageIndex: 0, displayNumber: 50, owner: 'section:authored-parity', parityBlank: false },
      { pageIndex: 1, displayNumber: 51, owner: 'section:authored-parity', parityBlank: true },
      { pageIndex: 2, displayNumber: 52, owner: 'section:authored-parity', parityBlank: false },
    ]);
    expect(layout.pages[1]).toMatchObject({
      sectionRegions: [],
      layers: {
        roots: [],
        background: [],
        behindText: [],
        header: [],
        body: [],
        notes: [],
        front: [],
        footer: [],
      },
      readingOrder: [],
    });
    expect(layout.pages.map((page) => page.layers.body.map((node) => node.source.path[0])))
      .toEqual([[0], [], [2]]);
    expect(finalPass).toEqual(expect.arrayContaining([
      expect.objectContaining({ sourceIndex: 0, physicalPageIndex: 0, totalPages: 3 }),
      expect.objectContaining({ sourceIndex: 2, physicalPageIndex: 2, totalPages: 3 }),
    ]));
  });

  it('keeps selected page-number metadata immutable across keyed layout variants', () => {
    const heights = new Map([[0, 20]]);
    const services = createServices({ heights });
    const store = new LayoutVariantStore(
      services,
      { currentDateMs: 100 },
      (options: LayoutOptions) => paginateBody({
        source: { story: 'body', storyInstance: 'body', path: [] },
        initialSection: owner('section:variant', {
          pageNumbering: {
            start: options.currentDateMs === 100 ? 1 : 20,
            format: null,
          },
        }),
        sequence: [bodyBlock(0)],
      }, services, options),
    );
    const defaultPageNumber = store.defaultLayout.pages[0]!.pageNumber;
    const selected = store.selectPage({ currentDateMs: 200 }, 0);

    expect(defaultPageNumber.displayNumber).toBe(1);
    expect(selected.page.pageNumber.displayNumber).toBe(20);
    expect(store.defaultLayout.pages[0]!.pageNumber).toBe(defaultPageNumber);
    expect(Object.isFrozen(defaultPageNumber)).toBe(true);
    expect(Object.isFrozen(selected.page.pageNumber)).toBe(true);
  });

  it('separates continuous content numbering from the first owned story page', () => {
    const observations: StoryObservation[] = [];
    const outgoing = owner('section:outgoing', {
      headers: { default: storySource('outgoing-default'), first: null, even: null },
    });
    const incoming = owner('section:incoming', {
      startType: 'continuous',
      pageNumbering: { start: 50, format: null },
      titlePage: true,
      evenAndOddHeaders: true,
      headers: {
        default: storySource('incoming-default'),
        first: storySource('incoming-first'),
        even: storySource('incoming-even'),
      },
    });
    const layout = paginate(outgoing, [
      bodyBlock(0),
      { kind: 'begin-section', source: bodySource(10), section: incoming },
      bodyBlock(1),
      bodyBlock(2),
    ], new Map([[0, 20], [1, 60], [2, 30]]), { storyObservations: observations });

    expect(layout.pages.map((page) => page.pageNumber.displayNumber)).toEqual([1, 51]);
    expect(new Set(observations.map(({ physicalPageIndex, storyInstance }) => (
      `${physicalPageIndex}:${storyInstance}`
    )))).toEqual(new Set([
      '0:outgoing-default',
      '1:incoming-first',
    ]));
  });

  it('selects the even story from the resolved displayed page number', () => {
    const observations: StoryObservation[] = [];
    const layout = paginate(owner('section:even-story', {
      pageNumbering: { start: 2, format: null },
      evenAndOddHeaders: true,
      headers: {
        default: storySource('default'),
        first: null,
        even: storySource('even'),
      },
    }), [bodyBlock(0)], new Map([[0, 20]]), { storyObservations: observations });

    expect(layout.pages[0]!.pageNumber.displayNumber).toBe(2);
    expect(new Set(observations.map(({ storyInstance }) => storyInstance)))
      .toEqual(new Set(['even']));
  });
});

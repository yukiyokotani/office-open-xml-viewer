import { describe, expect, it } from 'vitest';
import type { SectionLayoutContext } from '../layout-context.js';
import type {
  BodyLayoutInput,
  BodySectionLayoutInput,
  BodyStoryReferenceSet,
} from './body-layout-input.js';
import type { BodyLayoutKernel } from './body-layout-kernel.js';
import { paginateBody } from './body-paginator.js';
import { attachBodyLayoutKernel } from './runtime-state.js';
import type { LayoutServices, ParagraphLayout, SourceRef, StoryLayout, TextPlacement } from './types.js';

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

const storySource = (
  story: 'header' | 'footer',
  storyInstance: 'default' | 'first' | 'even',
): SourceRef => ({ story, storyInstance, path: [] });

const noStories = (): BodyStoryReferenceSet => ({
  default: null, first: null, even: null,
});

function section(verticalAlignment = 'top'): SectionLayoutContext {
  return {
    geometry: {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5,
    },
    columns: [{ xPt: 10, wPt: 180 }],
    columnSeparator: false,
    grid: { kind: 'none', linePitchPt: null, charSpacePt: null },
    textDirection: 'lrTb',
    verticalAlignment,
  };
}

function owner(
  sectionOccurrenceId: string,
  context: SectionLayoutContext,
  overrides: Partial<BodySectionLayoutInput> = {},
): BodySectionLayoutInput {
  return {
    sectionOccurrenceId,
    source: bodySource(100),
    startType: 'nextPage',
    context,
    pageNumbering: { start: null, format: null },
    titlePage: false,
    evenAndOddHeaders: false,
    headers: noStories(),
    footers: noStories(),
    pageBordersAuthored: false,
    pageBorders: null,
    pageLayout: {
      physicalGeometry: context.geometry,
      columns: null,
      textDirection: 'lrTb',
      gutterPt: 0,
      rtlGutter: false,
      sheetOrientation: 'portrait' as const,
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
  const source = bodySource(index);
  return {
    kind: 'paragraph',
    id: `paragraph:${index}`,
    source,
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

function emptyStoryLayout(source: SourceRef, heightPt: number, pageIndex = 0): StoryLayout {
  const bounds = { xPt: 10, yPt: 0, widthPt: 180, heightPt };
  return {
    story: source.story,
    flowBounds: bounds,
    inkBounds: bounds,
    blocks: heightPt > 0 ? [{
      ...paragraph(0, heightPt),
      id: `story:${source.story}:${source.storyInstance}:page:${pageIndex}:paragraph`,
      source,
      flowDomainId: `story:${source.story}:page:${pageIndex}`,
      flowBounds: bounds,
      inkBounds: bounds,
      borders: [{
        from: { xPt: 10, yPt: 0 }, to: { xPt: 190, yPt: 0 },
        color: '#000000', widthPt: 0.5, authoredStyle: 'single', style: 'solid' as const,
      }],
    }] : [],
    advancePt: heightPt,
    diagnostics: [],
  };
}

function plainTextStoryLayout(source: SourceRef, heightPt: number, text: string, pageIndex: number): StoryLayout {
  const base = emptyStoryLayout(source, heightPt, pageIndex);
  const bounds = { xPt: 10, yPt: 0, widthPt: 180, heightPt };
  const placement = {
    kind: 'text', text, range: { start: 0, end: text.length },
    origin: { xPt: 10, yPt: 8 }, bounds,
    advancePt: 5, clusters: [], decorations: [],
    paintOps: [{ text, range: { start: 0, end: text.length },
      offset: { xPt: 0, yPt: 0 }, letterSpacingPt: 0, scaleX: 1,
      direction: 'ltr', kerning: 'auto', writingMode: 'horizontal-tb' }],
    color: { kind: 'explicit', color: '#000000' },
    fontRoute: { familyList: 'serif', scope: 'native', fingerprint: 'serif' },
    fontSizePt: 36, fontWeight: 400, fontStyle: 'normal', direction: 'ltr',
  } satisfies TextPlacement;
  return {
    ...base,
    blocks: [{
      ...paragraph(0, heightPt),
      id: `story:${source.story}:${source.storyInstance}:page:${pageIndex}:paragraph`,
      source,
      flowDomainId: `story:${source.story}:page:${pageIndex}`,
      flowBounds: bounds,
      inkBounds: bounds,
      lines: [{ range: { start: 0, end: text.length }, bounds, baselinePt: 8,
        advancePt: heightPt, placements: [placement] }],
    }],
  };
}

function paginate(input: Readonly<{
  initialSection: BodySectionLayoutInput;
  sequence: BodyLayoutInput['sequence'];
  heightPt?: number | ((source: SourceRef) => number);
  storyExtentPt?: (source: SourceRef, pageIndex: number) => number;
  storyLayout?: (source: SourceRef, pageIndex: number) => StoryLayout;
  measuredStories?: Array<Readonly<{ source: SourceRef; pageIndex: number }>>;
  measuredBodyPagesByPass?: number[][];
}>) {
  const height = input.heightPt ?? 10;
  const session: ReturnType<BodyLayoutKernel['openBodyLayoutSession']> = {
    hasPaginationFields: false,
    measureParagraph: ({ input: paragraphInput, location }) => {
      input.measuredBodyPagesByPass?.at(-1)?.push(location.pageIndex);
      const heightPt = typeof height === 'function' ? height(paragraphInput.source) : height;
      return {
        layout: paragraph(paragraphInput.source.path[0]!, heightPt),
        blockExtentPt: heightPt,
        fragmentation: { kind: 'indivisible' },
      };
    },
    measureTable: () => { throw new Error('unused'); },
    layoutStory: ({ source, pageIndex }) => {
      input.measuredStories?.push({ source, pageIndex });
      return input.storyLayout?.(source, pageIndex)
        ?? emptyStoryLayout(source, input.storyExtentPt?.(source, pageIndex) ?? 0, pageIndex);
    },
    measureFollowingBlock: ({ input: following }) => {
      const heightPt = typeof height === 'function' ? height(following.source) : height;
      return { fullExtentPt: heightPt, leadContentExtentPt: heightPt };
    },
    measureLineNumberGlyph: () => ({ widthPt: 0, ascentPt: 0, descentPt: 0 }),
    resetPageAcquisition: () => undefined,
    moveAcquisitionCursor: () => undefined,
    flowRegistrySnapshot: emptyFlowRegistrySnapshot,
    commitFlowRegistryDelta: () => undefined,
  };
  const services = Object.freeze({
    text: { fingerprint: 'text' },
    images: { fingerprint: 'images' },
    math: { fingerprint: 'math' },
  }) as LayoutServices;
  attachBodyLayoutKernel(services, {
    openBodyLayoutSession: () => {
      input.measuredBodyPagesByPass?.push([]);
      return session;
    },
  });
  return paginateBody({
    source: { story: 'body', storyInstance: 'body', path: [] },
    initialSection: input.initialSection,
    sequence: input.sequence,
  }, services, { currentDateMs: 0 });
}

describe('canonical header/footer reservation', () => {
  const header = storySource('header', 'default');
  const footer = storySource('footer', 'default');

  // Windows Word 16.0.20430 controls: empty/space paragraphs reserve their
  // acquired extent just like visible text, across header slots and styles.
  it.each([' ', ''])('reserves body space for a blank header (%j)', (text) => {
    const sectionOwner = owner('section:space-header', section(), {
      headers: { default: header, first: null, even: null },
    });
    const layout = paginate({
      initialSection: sectionOwner,
      sequence: [bodyBlock(0)],
      storyLayout: (source, pageIndex) => plainTextStoryLayout(source, 36, text, pageIndex),
    });

    expect(layout.pages[0]?.geometry.contentTopPt).toBe(41);
  });

  it('reserves body space for a visible header glyph', () => {
    const sectionOwner = owner('section:visible-header', section(), {
      headers: { default: header, first: null, even: null },
    });
    const layout = paginate({
      initialSection: sectionOwner,
      sequence: [bodyBlock(0)],
      storyLayout: (source, pageIndex) => plainTextStoryLayout(source, 36, 'X', pageIndex),
    });

    expect(layout.pages[0]?.geometry.contentTopPt).toBe(41);
  });

  it('reserves the full extent of multiple blank header paragraphs', () => {
    const sectionOwner = owner('section:multi-space-header', section(), {
      headers: { default: header, first: null, even: null },
    });
    const layout = paginate({
      initialSection: sectionOwner,
      sequence: [bodyBlock(0)],
      storyLayout: (source, pageIndex) => {
        const single = plainTextStoryLayout(source, 36, ' ', pageIndex);
        const secondBounds = { ...single.blocks[0].flowBounds, yPt: 36 };
        return { ...single, advancePt: 72, blocks: [single.blocks[0], {
          ...single.blocks[0],
          id: `${single.blocks[0].id}:second`,
          source: { ...source, path: [1] },
          flowBounds: secondBounds,
          inkBounds: secondBounds,
        }] };
      },
    });

    expect(layout.pages[0]?.geometry.contentTopPt).toBe(77);
  });

  it('paginates and constructs pages from the same reduced body interval', () => {
    const sectionOwner = owner('section:reserved', section(), {
      headers: { default: header, first: null, even: null },
      footers: { default: footer, first: null, even: null },
    });
    const layout = paginate({
      initialSection: sectionOwner,
      sequence: [bodyBlock(0), bodyBlock(1)],
      heightPt: 30,
      // §17.6.11: the 5pt margin-to-distance allowance leaves 20pt/10pt overflow.
      storyExtentPt: (source) => source.story === 'header' ? 25 : 15,
    });

    expect(layout.pages).toHaveLength(2);
    expect(layout.pages.map((page) => page.geometry)).toEqual([
      expect.objectContaining({ contentTopPt: 30, contentBottomPt: 80 }),
      expect.objectContaining({ contentTopPt: 30, contentBottomPt: 80 }),
    ]);
    expect(layout.pages.map((page) => page.sectionRegions[0])).toEqual([
      expect.objectContaining({ blockStartPt: 30, blockEndPt: 80 }),
      expect.objectContaining({ blockStartPt: 30, blockEndPt: 80 }),
    ]);
    expect(layout.pages.map((page) => page.layers.body.map((node) => node.source.path[0])))
      .toEqual([[0], [1]]);
  });

  it.each([
    ['top', 30],
    ['center', 50],
    ['bottom', 70],
  ] as const)('aligns %s body flow inside the reserved interval', (alignment, expectedTopPt) => {
    const sectionOwner = owner(`section:${alignment}`, section(alignment), {
      headers: { default: header, first: null, even: null },
      footers: { default: footer, first: null, even: null },
    });
    const layout = paginate({
      initialSection: sectionOwner,
      sequence: [bodyBlock(0)],
      heightPt: 10,
      storyExtentPt: (source) => source.story === 'header' ? 25 : 15,
    });

    expect(layout.pages[0]?.geometry).toMatchObject({ contentTopPt: 30, contentBottomPt: 80 });
    expect(layout.pages[0]?.layers.body[0]?.flowBounds.yPt).toBe(expectedTopPt);
  });

  it.each([
    {
      name: 'first',
      titlePage: true,
      evenAndOddHeaders: false,
      pageNumberStart: null,
      expected: 'first',
    },
    {
      name: 'even',
      titlePage: false,
      evenAndOddHeaders: true,
      pageNumberStart: 2,
      expected: 'even',
    },
    {
      name: 'default',
      titlePage: false,
      evenAndOddHeaders: true,
      pageNumberStart: 1,
      expected: 'default',
    },
  ])('measures the $name story selected by section occurrence and displayed parity', ({
    titlePage,
    evenAndOddHeaders,
    pageNumberStart,
    expected,
  }) => {
    const measuredStories: Array<Readonly<{ source: SourceRef; pageIndex: number }>> = [];
    const sectionOwner = owner('section:stories', section(), {
      titlePage,
      evenAndOddHeaders,
      pageNumbering: { start: pageNumberStart, format: null },
      headers: {
        default: storySource('header', 'default'),
        first: storySource('header', 'first'),
        even: storySource('header', 'even'),
      },
    });
    paginate({ initialSection: sectionOwner, sequence: [bodyBlock(0)], measuredStories });

    expect(new Set(measuredStories.map(({ source }) => source.storyInstance))).toEqual(new Set([expected]));
  });

  it('keeps an absent selected story blank instead of borrowing the default slot', () => {
    const measuredStories: Array<Readonly<{ source: SourceRef; pageIndex: number }>> = [];
    const sectionOwner = owner('section:blank-even', section(), {
      evenAndOddHeaders: true,
      pageNumbering: { start: 2, format: null },
      headers: { default: header, first: null, even: null },
    });
    const layout = paginate({
      initialSection: sectionOwner,
      sequence: [bodyBlock(0)],
      measuredStories,
      storyExtentPt: () => 25,
    });

    expect(measuredStories).toEqual([]);
    expect(layout.pages[0]?.geometry).toMatchObject({ contentTopPt: 10, contentBottomPt: 90 });
  });

  it('leaves a horizontal page with no selected story at its margin interval', () => {
    const layout = paginate({
      initialSection: owner('section:no-story', section()),
      sequence: [bodyBlock(0)],
      storyExtentPt: () => { throw new Error('No story may be measured'); },
    });

    expect(layout.pages[0]?.geometry).toMatchObject({ contentTopPt: 10, contentBottomPt: 90 });
    expect(layout.pages[0]?.sectionRegions[0]).toMatchObject({ blockStartPt: 10, blockEndPt: 90 });
    expect(layout.pages[0]?.layers.body[0]?.flowBounds.yPt).toBe(10);
  });

  it('does not reserve an empty selected footer whose anchor lies inside the body margin', () => {
    const base = section();
    const context: SectionLayoutContext = {
      ...base,
      geometry: { ...base.geometry, footerDistance: 15 },
    };
    const sectionOwner = owner('section:empty-footer', context, {
      footers: { default: footer, first: null, even: null },
    });
    const layout = paginate({
      initialSection: sectionOwner,
      sequence: [bodyBlock(0), bodyBlock(1)],
      heightPt: 40,
      storyExtentPt: () => 0,
    });

    expect(layout.pages).toHaveLength(1);
    expect(layout.pages[0]?.geometry).toMatchObject({ contentTopPt: 10, contentBottomPt: 90 });
    expect(layout.pages[0]?.layers.body).toHaveLength(2);
  });

  it('uses the section occurrence that owns the physical page across a continuous boundary', () => {
    const measuredStories: Array<Readonly<{ source: SourceRef; pageIndex: number }>> = [];
    const outgoingHeader = storySource('header', 'default');
    const incomingHeader: SourceRef = {
      story: 'header', storyInstance: 'section:incoming:default', path: [],
    };
    const outgoing = owner('section:outgoing', section(), {
      headers: { default: outgoingHeader, first: null, even: null },
    });
    const incoming = owner('section:incoming', section(), {
      startType: 'continuous',
      headers: { default: incomingHeader, first: null, even: null },
    });
    const heights = new Map([[0, 20], [1, 60], [2, 30]]);
    const layout = paginate({
      initialSection: outgoing,
      sequence: [
        bodyBlock(0),
        { kind: 'begin-section', source: bodySource(3), section: incoming },
        bodyBlock(1),
        bodyBlock(2),
      ],
      heightPt: (source) => heights.get(source.path[0]!) ?? 0,
      measuredStories,
    });

    expect(layout.pages.map((page) => page.sectionOccurrenceId)).toEqual([
      'section:outgoing', 'section:incoming',
    ]);
    expect(new Set(measuredStories.map(({ source, pageIndex }) => (
      `${pageIndex}:${source.storyInstance}`
    )))).toEqual(new Set([
      '0:default',
      '1:section:incoming:default',
    ]));
  });
});

it('issue #1809 applies an unchanged selected header reserve on newly opened pages', () => {
  const measuredBodyPagesByPass: number[][] = [];
  const layout = paginate({
    initialSection: owner('section:constant-header', section(), {
      headers: { default: storySource('header', 'default'), first: null, even: null },
    }),
    sequence: Array.from({ length: 24 }, (_, index) => bodyBlock(index)),
    heightPt: 10,
    // 45pt story - (10pt margin - 5pt distance) = 40pt body reserve.
    storyExtentPt: () => 45,
    measuredBodyPagesByPass,
  });
  expect(layout.pages).toHaveLength(6);
  // One discovery pass, one pass with the known constant story geometry,
  // then one equality check. New pages should not start another discovery wave.
  expect(measuredBodyPagesByPass.map(pages => Math.max(...pages) + 1)).toEqual([3, 6, 6]);
});

it.each(['nextPage', 'evenPage'] as const)('discovers destination first/even slots across a %s boundary', startType => {
  const measuredBodyPagesByPass: number[][] = [];
  const first = storySource('header', 'first');
  const even = storySource('header', 'even');
  const normal = storySource('header', 'default');
  const incoming = owner('section:incoming-slots', section(), {
    startType,
    titlePage: true, evenAndOddHeaders: true,
    pageNumbering: { start: 2, format: 'decimal' },
    headers: { default: normal, first, even },
  });
  const layout = paginate({
    initialSection: owner('section:preceding', section(), {
      headers: { default: normal, first: null, even: null },
    }),
    sequence: [...Array.from({ length: 8 }, (_, i) => bodyBlock(i)),
      { kind: 'begin-section', source: bodySource(100), section: incoming },
      ...Array.from({ length: 24 }, (_, i) => bodyBlock(i + 8))],
    heightPt: 10,
    storyExtentPt: source => source.storyInstance === 'first' ? 65 : source.storyInstance === 'even' ? 25 : 45,
    measuredBodyPagesByPass,
  });
  // 80pt nominal body, with 60/40/20pt reserves for first/default/even.
  // Outgoing: 4+4. Incoming: 2+4+6+4+6+2.
  const blank = startType === 'evenPage' ? [0] : [];
  expect(layout.pages.map(page => page.layers.body.length)).toEqual([4, 4, ...blank, 2, 4, 6, 4, 6, 2]);
  expect(layout.pages.slice(2 + blank.length).map(page => page.pageNumber.displayNumber)).toEqual([2, 3, 4, 5, 6, 7]);
  expect(measuredBodyPagesByPass.map(pages => Math.max(...pages) + 1)).toEqual([4, 8 + blank.length, 8 + blank.length]);
});

// An authored blank precedes the section's first retained content. Its exact
// numbering is learned by finalization; reopening must preserve that authority.
it('reopens an authored empty page with the observed restart anchor', () => {
  const layout = paginate({
    initialSection: owner('section:empty-opening', section(), {
      pageNumbering: { start: 2, format: 'decimal' },
      evenAndOddHeaders: true,
      headers: { default: storySource('header', 'default'), first: null,
        even: storySource('header', 'even') },
    }),
    sequence: [{ kind: 'authored-break', source: bodySource(99), break: 'page' },
      ...Array.from({ length: 12 }, (_, index) => bodyBlock(index))],
    heightPt: 10,
    storyExtentPt: source => source.storyInstance === 'even' ? 25 : 45,
  });
  expect(layout.pages.map(page => page.pageNumber.displayNumber)).toEqual([1, 2, 3, 4]);
  expect(layout.pages.map(page => page.layers.body.length)).toEqual([0, 6, 4, 2]);
});

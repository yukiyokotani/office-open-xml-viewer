import { describe, expect, it } from 'vitest';
import type { CanvasFontRoute } from '@silurus/ooxml-core';
import { buildPageLayers } from './layout/page-layers.js';
import { textRunGeometryForPage } from './layout/text-index.js';
import { textRunsForPage } from './text-run-projection.js';
import type {
  DocumentLayout,
  DrawingLayout,
  LayoutPage,
  LayoutRect,
  Matrix2DData,
  NoteLayout,
  PageLayers,
  ParagraphPlacement,
  ParagraphLayout,
  ResolvedFloatingTablePlacementLayout,
  SourceRef,
  TableLayout,
  TextBoxLayout,
  TextPlacement,
} from './layout/types.js';

const fontRoute = Object.freeze({
  familyList: '"Index Sans"',
  scope: 'native',
  fingerprint: 'text-index-test',
}) satisfies CanvasFontRoute;

const identity = Object.freeze({
  a: 1, b: 0, c: 0, d: 1, e: 0, f: 0,
}) satisfies Matrix2DData;

function rect(xPt = 0, yPt = 0, widthPt = 100, heightPt = 20): LayoutRect {
  return Object.freeze({ xPt, yPt, widthPt, heightPt });
}

function source(story: SourceRef['story'], path: readonly number[]): SourceRef {
  return Object.freeze({
    story,
    storyInstance: `${story}:test`,
    path: Object.freeze([...path]),
  });
}

function placement(
  text: string,
  xPt: number,
  yPt: number,
  options: Readonly<{
    rangeStart?: number;
    direction?: 'ltr' | 'rtl';
    hyperlink?: TextPlacement['hyperlink'];
    letterSpacingPt?: number;
    tateChuYoko?: boolean;
    trailingSpaceCompressionPt?: number;
  }> = {},
): TextPlacement {
  const rangeStart = options.rangeStart ?? 0;
  const range = Object.freeze({ start: rangeStart, end: rangeStart + text.length });
  const direction = options.direction ?? 'ltr';
  return Object.freeze({
    kind: 'text',
    text,
    range,
    origin: Object.freeze({ xPt, yPt: yPt + 8 }),
    bounds: rect(xPt, yPt, text.length * 5, 10),
    advancePt: text.length * 5,
    clusters: Object.freeze([Object.freeze({
      range,
      offset: Object.freeze({ xPt: 0, yPt: 0 }),
      advancePt: text.length * 5,
    })]),
    paintOps: Object.freeze([Object.freeze({
      text,
      range,
      offset: Object.freeze({ xPt: 0, yPt: 0 }),
      letterSpacingPt: options.letterSpacingPt ?? 0,
      scaleX: 1,
      direction,
      kerning: 'auto',
      writingMode: 'horizontal-tb',
    })]),
    color: Object.freeze({ kind: 'explicit', color: '#000000' }),
    fontRoute,
    fontSizePt: 10,
    fontWeight: 400,
    fontStyle: 'normal',
    direction,
    decorations: Object.freeze([]),
    ...(options.hyperlink ? { hyperlink: options.hyperlink } : {}),
    ...(options.tateChuYoko ? { tateChuYoko: true } : {}),
    ...(options.trailingSpaceCompressionPt
      ? { trailingSpaceCompressionPt: options.trailingSpaceCompressionPt }
      : {}),
  });
}

function paragraph(
  id: string,
  story: SourceRef['story'],
  placements: readonly ParagraphPlacement[],
  options: Readonly<{
    drawings?: readonly DrawingLayout[];
    textBoxes?: readonly TextBoxLayout[];
    flowDomainId?: string;
    paragraphId?: string;
  }> = {},
): ParagraphLayout {
  const bounds = rect();
  return Object.freeze({
    kind: 'paragraph',
    id,
    source: source(story, [0]),
    flowDomainId: options.flowDomainId ?? `${story}:domain`,
    ...(options.paragraphId ? { paragraphId: options.paragraphId } : {}),
    flowBounds: bounds,
    inkBounds: bounds,
    advancePt: 20,
    ordinaryFlow: story === 'body',
    spacing: Object.freeze({ beforePt: 0, afterPt: 0 }),
    contextualSpacing: false,
    lines: Object.freeze([Object.freeze({
      range: Object.freeze({ start: 0, end: placements.length }),
      bounds,
      baselinePt: 8,
      advancePt: 20,
      placements: Object.freeze([...placements]),
    })]),
    borders: Object.freeze([]),
    resources: Object.freeze([]),
    drawings: Object.freeze([...(options.drawings ?? [])]),
    textBoxes: Object.freeze([...(options.textBoxes ?? [])]),
    events: Object.freeze([]),
    exclusions: Object.freeze([]),
  });
}

function table(id: string, child: ParagraphLayout | TableLayout): TableLayout {
  const bounds = rect();
  return Object.freeze({
    kind: 'table',
    id,
    source: source('body', [1]),
    flowDomainId: 'body:domain',
    flowBounds: bounds,
    inkBounds: bounds,
    advancePt: 40,
    ordinaryFlow: true,
    columnWidthsPt: Object.freeze([100]),
    rows: Object.freeze([Object.freeze({
      kind: 'table-row',
      id: `${id}:row`,
      source: source('body', [1, 0]),
      flowDomainId: 'body:domain',
      flowBounds: rect(10, 20, 100, 40),
      inkBounds: bounds,
      advancePt: 40,
      ordinaryFlow: true,
      heightPt: 40,
      contentHeightPt: 40,
      cells: Object.freeze([Object.freeze({
        kind: 'table-cell',
        id: `${id}:cell`,
        source: source('body', [1, 0, 0]),
        flowDomainId: 'body:domain',
        flowBounds: rect(10, 20, 100, 40),
        inkBounds: bounds,
        advancePt: 40,
        ordinaryFlow: true,
        contentBounds: rect(15, 20, 90, 40),
        verticalMerge: 'none',
        vAlign: 'top',
        blocks: Object.freeze([Object.freeze({
          layout: child,
          offsetPt: 7,
          advancePt: child.advancePt,
        })]),
      })]),
    })]),
    borders: Object.freeze([]),
  });
}

function textBox(id: string, text: string): TextBoxLayout {
  const bounds = rect();
  const child = paragraph(`${id}:paragraph`, 'textbox', [placement(text, 1, 2)], {
    flowDomainId: `${id}:domain`,
  });
  return Object.freeze({
    kind: 'textbox',
    id,
    source: source('textbox', [0]),
    flowDomainId: `${id}:domain`,
    flowBounds: bounds,
    inkBounds: bounds,
    advancePt: 20,
    ordinaryFlow: false,
    story: Object.freeze({
      story: 'textbox',
      flowBounds: bounds,
      inkBounds: bounds,
      blocks: Object.freeze([child]),
      advancePt: 20,
      diagnostics: Object.freeze([]),
    }),
    transform: Object.freeze({ a: 1, b: 0, c: 0, d: 1, e: 5, f: 6 }),
    writingMode: 'horizontal-tb',
    insets: Object.freeze({ topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 }),
  });
}

function anchoredDrawing(
  id: string,
  textBoxId: string,
  options: Readonly<{
    behindDoc: boolean;
    relativeHeight: number;
    sourceOrder: number;
    horizontalOwnership?: 'page' | 'host';
    verticalOwnership?: 'page' | 'host';
    orientation?: DrawingLayout['orientation'];
    transform?: Matrix2DData;
  }>,
): DrawingLayout {
  return Object.freeze({
    kind: 'drawing',
    id,
    source: source('body', [options.sourceOrder]),
    flowDomainId: 'body:domain',
    flowBounds: rect(),
    inkBounds: rect(),
    advancePt: 0,
    ordinaryFlow: false,
    commands: Object.freeze([]),
    ...(options.orientation ? { orientation: options.orientation } : {}),
    ...(options.transform ? { transform: options.transform } : {}),
    anchorLayer: Object.freeze({
      occurrenceId: `anchor:${id}`,
      behindDoc: options.behindDoc,
      relativeHeight: options.relativeHeight,
      sourceOrder: options.sourceOrder,
      horizontalOwnership: options.horizontalOwnership ?? 'host',
      verticalOwnership: options.verticalOwnership ?? 'host',
    }),
    textBoxIds: Object.freeze([textBoxId]),
  });
}

function resolvedFloatingTable(
  child: TableLayout,
  xPt: number,
  yPt: number,
): ResolvedFloatingTablePlacementLayout {
  const bounds = rect(xPt, yPt, child.flowBounds.widthPt, child.flowBounds.heightPt);
  return Object.freeze({
    kind: 'resolved-floating-table-placement',
    occurrenceId: `resolved:${child.id}`,
    xPt,
    yPt,
    bounds,
    exclusionBounds: bounds,
    overlap: 'overlap',
    child,
    source: Object.freeze({
      kind: 'floating-table-placement',
      occurrenceId: `source:${child.id}`,
      ownership: 'source',
      physicalPageIndex: 0,
      displayPageNumber: 1,
      hostCellId: 'host-cell',
      sourceBlockIndex: 0,
      anchorBlockIndex: 0,
      tableId: child.id,
      overlap: 'overlap',
      positioning: Object.freeze({
        leftFromTextPt: 0,
        rightFromTextPt: 0,
        topFromTextPt: 0,
        bottomFromTextPt: 0,
        horzAnchor: 'page',
        horzSpecified: true,
        vertAnchor: 'page',
        xPt,
        yPt,
      }),
      anchorBounds: rect(),
      child,
    }),
  });
}

function note(id: string, text: string): NoteLayout {
  const bounds = rect();
  return Object.freeze({
    kind: 'note',
    id,
    source: source('footnote', [0]),
    flowDomainId: 'footnote:domain',
    flowBounds: bounds,
    inkBounds: bounds,
    advancePt: 20,
    ordinaryFlow: false,
    separator: Object.freeze([]),
    story: Object.freeze({
      story: 'footnote',
      flowBounds: bounds,
      inkBounds: bounds,
      blocks: Object.freeze([
        paragraph(`${id}:paragraph`, 'footnote', [placement(text, 0, 0)], {
          flowDomainId: 'footnote:domain',
        }),
      ]),
      advancePt: 20,
      diagnostics: Object.freeze([]),
    }),
  });
}

function page(layers: PageLayers, readingOrder: readonly string[]): LayoutPage {
  const geometry = Object.freeze({
    ...rect(0, 0, 200, 300),
    contentTopPt: 20,
    contentBottomPt: 280,
  });
  return {
    pageIndex: 0,
    geometry,
    flowDomains: [
      {
        id: 'header:domain', kind: 'header',
        logicalBounds: geometry, physicalBounds: geometry,
      },
      {
        id: 'body:domain', kind: 'body',
        sectionRegionId: 'region',
        logicalBounds: geometry, physicalBounds: geometry,
      },
      {
        id: 'footnote:domain', kind: 'footnote',
        sectionRegionId: 'region',
        logicalBounds: geometry, physicalBounds: geometry,
      },
      {
        id: 'footer:domain', kind: 'footer',
        logicalBounds: geometry, physicalBounds: geometry,
      },
    ],
    section: {} as LayoutPage['section'],
    sectionOccurrenceId: 'section',
    parityBlank: false,
    bookmarkStarts: [],
    pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section' },
    sectionRegions: [{
      id: 'region',
      sectionOccurrenceId: 'section',
      coordinateSpace: {
        writingMode: 'horizontal-tb',
        logicalToPhysical: identity,
        physicalToLogical: identity,
      },
      blockStartPt: 0,
      blockEndPt: 300,
      columnFlowDirection: 'ltr',
      columnIndexes: [0],
      flowDomainIds: ['body:domain'],
      section: {} as LayoutPage['section'],
    }],
    columnSeparators: [],
    pageBorder: null,
    layers,
    readingOrder,
  };
}

function documentLayout(layoutPage: LayoutPage): DocumentLayout {
  return Object.freeze({
    pages: Object.freeze([layoutPage]),
    diagnostics: Object.freeze([]),
  });
}

describe('textRunsForPage', () => {
  it('projects retained terminal-space compression into CSS pixels', () => {
    const compressed = paragraph('compressed', 'body', [
      placement('A ', 0, 0, { trailingSpaceCompressionPt: 2 }),
    ]);
    const layers = buildPageLayers([
      { layer: 'body', node: compressed, coordinateSpace: 'section-logical' },
    ]);
    expect(textRunsForPage(documentLayout(page(layers, [compressed.id])), 0, { scale: 2 })[0])
      .toMatchObject({ text: 'A ', trailingSpaceCompressionPx: 4 });
  });

  it('projects the source w14:paraId onto every run owned by that paragraph', () => {
    const identified = paragraph('identified', 'body', [
      placement('first', 0, 0),
      placement('second', 25, 0),
    ], { paragraphId: '1A2B3C4D' });
    const positionalOnly = paragraph('positional', 'body', [placement('third', 0, 20)]);
    const layers = buildPageLayers([
      { layer: 'body', node: identified, coordinateSpace: 'section-logical' },
      { layer: 'body', node: positionalOnly, coordinateSpace: 'section-logical' },
    ]);

    const runs = textRunsForPage(
      documentLayout(page(layers, [identified.id, positionalOnly.id])),
      0,
      { scale: 1 },
    );

    expect(runs.map(({ text, paragraphId, source }) => ({ text, paragraphId, source }))).toEqual([
      {
        text: 'first', paragraphId: '1A2B3C4D',
        source: { story: 'body', storyInstance: 'body:test', path: [0] },
      },
      {
        text: 'second', paragraphId: '1A2B3C4D',
        source: { story: 'body', storyInstance: 'body:test', path: [0] },
      },
      {
        text: 'third', paragraphId: undefined,
        source: { story: 'body', storyInstance: 'body:test', path: [0] },
      },
    ]);
  });

  it('projects retained visual placements in semantic story order without reshaping', () => {
    const header = paragraph('header', 'header', [placement('H', 1, 1)], {
      flowDomainId: 'header:domain',
    });
    const body = paragraph('body', 'body', [
      placement('RTL', 20, 10, {
        rangeStart: 3,
        direction: 'rtl',
        hyperlink: { kind: 'internal', ref: 'destination' },
        letterSpacingPt: 0.5,
      }),
      Object.freeze({
        kind: 'tab',
        range: Object.freeze({ start: 2, end: 3 }),
        bounds: rect(15, 10, 5, 10),
        advancePt: 5,
        leader: 'none',
      }),
      placement('fi', 5, 10, { rangeStart: 0 }),
    ]);
    const footnote = note('note', 'N');
    const footer = paragraph('footer', 'footer', [placement('F', 2, 250)], {
      flowDomainId: 'footer:domain',
    });
    const layers = buildPageLayers([
      { layer: 'header', node: header, coordinateSpace: 'section-logical' },
      { layer: 'body', node: body, coordinateSpace: 'section-logical' },
      { layer: 'notes', node: footnote, coordinateSpace: 'section-logical' },
      { layer: 'footer', node: footer, coordinateSpace: 'section-logical' },
    ]);

    const runs = textRunsForPage(
      documentLayout(page(layers, [header.id, body.id, footnote.id, footer.id])),
      0,
      { scale: 2 },
    );

    expect(runs.map((run) => run.text)).toEqual(['H', 'RTL', 'fi', 'N', 'F']);
    expect(runs[1]).toMatchObject({
      x: 40,
      y: 20,
      w: 30,
      h: 20,
      fontSize: 20,
      letterSpacingPx: 1,
      hyperlink: { kind: 'internal', ref: 'destination' },
    });
    expect(runs[2]?.font).toBe('normal 400 20px "Index Sans"');
  });

  it('projects table-cell and vertical page transforms into physical CSS geometry', () => {
    const child = paragraph('cell-paragraph', 'body', [
      placement('V', 2, 3, { tateChuYoko: true }),
    ]);
    const root = table('table', child);
    const layers = buildPageLayers([
      { layer: 'body', node: root, coordinateSpace: 'section-logical' },
    ]);
    const verticalPage = page(layers, [root.id]);
    const logicalToPhysical = Object.freeze({
      a: 0, b: 1, c: -1, d: 0, e: 200, f: 0,
    });
    const physicalToLogical = Object.freeze({
      a: 0, b: -1, c: 1, d: 0, e: 0, f: 200,
    });
    const transformedPage: LayoutPage = {
      ...verticalPage,
      sectionRegions: verticalPage.sectionRegions.map((region) => ({
        ...region,
        coordinateSpace: {
          writingMode: 'vertical-rl',
          logicalToPhysical,
          physicalToLogical,
        },
      })),
    };

    expect(textRunsForPage(documentLayout(transformedPage), 0, { scale: 2 }))
      .toEqual([expect.objectContaining({
        text: 'V',
        x: 340,
        y: 34,
        w: 10,
        h: 20,
        fontSize: 20,
        transform: 'rotate(90deg)',
        eastAsianVert: true,
      })]);
  });

  it('uses the first section region for an unbound note domain, matching paint', () => {
    const footnote = note('fallback-note', 'N');
    const layers = buildPageLayers([
      { layer: 'notes', node: footnote, coordinateSpace: 'section-logical' },
    ]);
    const fallbackPage = page(layers, [footnote.id]);
    const logicalToPhysical = Object.freeze({
      a: 0, b: 1, c: -1, d: 0, e: 200, f: 0,
    });
    const physicalToLogical = Object.freeze({
      a: 0, b: -1, c: 1, d: 0, e: 0, f: 200,
    });
    const transformedPage: LayoutPage = {
      ...fallbackPage,
      flowDomains: fallbackPage.flowDomains.map((domain) => (
        domain.id === 'footnote:domain'
          ? { ...domain, sectionRegionId: undefined }
          : domain
      )),
      sectionRegions: fallbackPage.sectionRegions.map((region) => ({
        ...region,
        coordinateSpace: {
          writingMode: 'vertical-rl',
          logicalToPhysical,
          physicalToLogical,
        },
      })),
    };

    expect(textRunsForPage(documentLayout(transformedPage), 0, { scale: 1 }))
      .toEqual([expect.objectContaining({
        text: 'N',
        x: 200,
        y: 0,
        transform: 'rotate(90deg)',
      })]);
  });

  it('rejects a note domain whose retained section region is missing, matching paint', () => {
    const footnote = note('dangling-note', 'N');
    const layers = buildPageLayers([
      { layer: 'notes', node: footnote, coordinateSpace: 'section-logical' },
    ]);
    const danglingPage = page(layers, [footnote.id]);
    const invalidPage: LayoutPage = {
      ...danglingPage,
      flowDomains: danglingPage.flowDomains.map((domain) => (
        domain.id === 'footnote:domain'
          ? { ...domain, kind: 'endnote', sectionRegionId: 'missing-region' }
          : domain
      )),
    };

    expect(() => textRunsForPage(documentLayout(invalidPage), 0, { scale: 1 }))
      .toThrow(
        'footnote:domain references missing page story region missing-region',
      );
  });

  it('composes nested table-cell placements without paint callbacks', () => {
    const child = paragraph('nested-cell-paragraph', 'body', [
      placement('nested', 2, 3),
    ]);
    const nested = table('nested-table', child);
    const outer = table('outer-table', nested);
    const layers = buildPageLayers([
      { layer: 'body', node: outer, coordinateSpace: 'section-logical' },
    ]);

    expect(textRunsForPage(
      documentLayout(page(layers, [outer.id])),
      0,
      { scale: 1 },
    )).toEqual([expect.objectContaining({
      text: 'nested',
      x: 32,
      y: 57,
    })]);
  });

  it('projects resolved floating tables from absolute retained page coordinates', () => {
    const floatingText = paragraph('floating-paragraph', 'body', [
      placement('FLOAT', 2, 3),
    ]);
    const floating = table('floating-table', floatingText);
    const hostBase = table(
      'host-table',
      paragraph('host-paragraph', 'body', []),
    );
    const host: TableLayout = Object.freeze({
      ...hostBase,
      resolvedFloatingTables: Object.freeze([
        resolvedFloatingTable(floating, 70, 90),
      ]),
    });
    const outer = table('outer-table', host);
    const layers = buildPageLayers([
      { layer: 'body', node: outer, coordinateSpace: 'section-logical' },
    ]);
    const layout = documentLayout(page(layers, [outer.id]));

    const geometry = textRunGeometryForPage(layout, 0);
    const floatingGeometry = geometry.find(({ placement: run }) => run.text === 'FLOAT');
    expect(floatingGeometry).toMatchObject({
      pointToPage: { a: 1, b: 0, c: 0, d: 1, e: 85, f: 117 },
    });
    expect(textRunsForPage(layout, 0, { scale: 2 }))
      .toEqual([expect.objectContaining({
        text: 'FLOAT',
        x: 174,
        y: 240,
      })]);
  });

  it('recursively visits anchors inside an owned text-box local stacking context', () => {
    const innerBox = textBox('inner-box', 'I');
    const innerDrawing = anchoredDrawing('inner-drawing', innerBox.id, {
      behindDoc: true,
      relativeHeight: 10,
      sourceOrder: 0,
    });
    const outerParagraph = paragraph('outer-box:paragraph', 'textbox', [
      placement('O', 1, 2),
    ], {
      flowDomainId: 'outer-box:domain',
      drawings: [innerDrawing],
      textBoxes: [innerBox],
    });
    const outerBoxBase = textBox('outer-box', 'unused');
    const outerBox: TextBoxLayout = Object.freeze({
      ...outerBoxBase,
      story: Object.freeze({
        ...outerBoxBase.story,
        blocks: Object.freeze([outerParagraph]),
      }),
    });
    const outerDrawing = anchoredDrawing('outer-drawing', outerBox.id, {
      behindDoc: false,
      relativeHeight: 20,
      sourceOrder: 0,
    });
    const body = paragraph('body', 'body', [placement('B', 0, 0)], {
      drawings: [outerDrawing],
      textBoxes: [outerBox],
    });
    const layers = buildPageLayers([
      { layer: 'body', node: body, coordinateSpace: 'section-logical' },
    ]);

    const runs = textRunsForPage(
      documentLayout(page(layers, [body.id])),
      0,
      { scale: 1 },
    );
    expect(runs.map((run) => run.text)).toEqual(['B', 'O', 'I']);
    expect(runs[2]).toMatchObject({ x: 11, y: 14 });
  });

  it('projects an upright drawing-owned text box through the retained inverse page turn', () => {
    const owned = textBox('upright-box', 'U');
    const drawing = anchoredDrawing('upright-drawing', owned.id, {
      behindDoc: false,
      relativeHeight: 1,
      sourceOrder: 0,
      orientation: 'upright-physical',
      transform: Object.freeze({ a: 0, b: -1, c: 1, d: 0, e: 65, f: 120 }),
    });
    const body = paragraph('body', 'body', [], {
      drawings: [drawing],
      textBoxes: [owned],
    });
    const layers = buildPageLayers([
      { layer: 'body', node: body, coordinateSpace: 'section-logical' },
    ]);
    const verticalPage = page(layers, [body.id]);
    const transformedPage: LayoutPage = {
      ...verticalPage,
      sectionRegions: verticalPage.sectionRegions.map((region) => ({
        ...region,
        coordinateSpace: {
          writingMode: 'vertical-rl',
          logicalToPhysical: { a: 0, b: 1, c: -1, d: 0, e: 200, f: 0 },
          physicalToLogical: { a: 0, b: -1, c: 1, d: 0, e: 0, f: 200 },
        },
      })),
    };

    expect(textRunGeometryForPage(documentLayout(transformedPage), 0)[0])
      .toMatchObject({
        placement: { text: 'U' },
        pointToPage: { a: 1, b: 0, c: 0, d: 1, e: 85, f: 71 },
      });
    expect(textRunsForPage(documentLayout(transformedPage), 0, { scale: 1 }))
      .toEqual([expect.objectContaining({ text: 'U', x: 86, y: 73 })]);
  });

  it('keeps anchored text-box sequence stable under z-order edits and uses retained frames for geometry', () => {
    const firstBox = textBox('first-box', 'T1');
    const secondBox = textBox('second-box', 'T2');
    const firstDrawing = anchoredDrawing('first-drawing', firstBox.id, {
      behindDoc: true,
      relativeHeight: 1,
      sourceOrder: 0,
      horizontalOwnership: 'page',
      verticalOwnership: 'host',
    });
    const secondDrawing = anchoredDrawing('second-drawing', secondBox.id, {
      behindDoc: false,
      relativeHeight: 500,
      sourceOrder: 1,
      horizontalOwnership: 'page',
      verticalOwnership: 'host',
    });
    const body = paragraph('body', 'body', [placement('B', 0, 0)], {
      drawings: [firstDrawing, secondDrawing],
      textBoxes: [secondBox, firstBox],
    });
    const built = buildPageLayers([
      { layer: 'body', node: body, coordinateSpace: 'section-logical' },
    ]);
    const layers: PageLayers = Object.freeze({
      ...built,
      paintOrder: Object.freeze(built.paintOrder.map((entry) => (
        entry.kind === 'drawing' ? Object.freeze({
          ...entry,
          frames: Object.freeze([Object.freeze({
            kind: 'transform' as const,
            transform: Object.freeze({ a: 1, b: 0, c: 0, d: 1, e: 30, f: 40 }),
          })]),
          layoutTranslationPt: Object.freeze({ xPt: 30, yPt: 40 }),
        }) : entry
      ))),
    });

    const original = textRunsForPage(
      documentLayout(page(layers, [body.id])),
      0,
      { scale: 2 },
    );

    const raisedFirstDrawing = anchoredDrawing('first-drawing', firstBox.id, {
      behindDoc: false,
      relativeHeight: 999,
      sourceOrder: 0,
      horizontalOwnership: 'page',
      verticalOwnership: 'host',
    });
    const raisedSecondDrawing = anchoredDrawing('second-drawing', secondBox.id, {
      behindDoc: true,
      relativeHeight: 0,
      sourceOrder: 1,
      horizontalOwnership: 'page',
      verticalOwnership: 'host',
    });
    const raisedBody = paragraph('body', 'body', [placement('B', 0, 0)], {
      drawings: [raisedFirstDrawing, raisedSecondDrawing],
      textBoxes: [firstBox, secondBox],
    });
    const raisedBuilt = buildPageLayers([
      { layer: 'body', node: raisedBody, coordinateSpace: 'section-logical' },
    ]);
    const raisedLayers: PageLayers = Object.freeze({
      ...raisedBuilt,
      paintOrder: Object.freeze(raisedBuilt.paintOrder.map((entry) => (
        entry.kind === 'drawing' ? Object.freeze({
          ...entry,
          frames: Object.freeze([Object.freeze({
            kind: 'transform' as const,
            transform: Object.freeze({ a: 1, b: 0, c: 0, d: 1, e: 30, f: 40 }),
          })]),
          layoutTranslationPt: Object.freeze({ xPt: 30, yPt: 40 }),
        }) : entry
      ))),
    });
    const raised = textRunsForPage(
      documentLayout(page(raisedLayers, [raisedBody.id])),
      0,
      { scale: 2 },
    );

    expect(original.map((run) => run.text)).toEqual(['B', 'T1', 'T2']);
    expect(raised).toEqual(original);
    expect(original[1]).toMatchObject({
      x: 12,
      y: 96,
      w: 20,
      h: 20,
    });
  });

  it('uses the retained paragraph run ordinal when anchor metadata is mixed', () => {
    const anchoredBox = textBox('anchored-box', 'A');
    const inlineBox = textBox('inline-box', 'I');
    const anchored = anchoredDrawing('anchored-drawing', anchoredBox.id, {
      behindDoc: false,
      relativeHeight: 1,
      sourceOrder: 50,
    });
    const { anchorLayer: _anchorLayer, ...inlineDrawing } = anchoredDrawing(
      'inline-drawing',
      inlineBox.id,
      {
        behindDoc: false,
        relativeHeight: 1,
        sourceOrder: 0,
      },
    );
    const body = paragraph('body', 'body', [placement('B', 0, 0)], {
      drawings: [anchored, inlineDrawing],
      textBoxes: [anchoredBox, inlineBox],
    });
    const layers = buildPageLayers([
      { layer: 'body', node: body, coordinateSpace: 'section-logical' },
    ]);

    expect(textRunsForPage(
      documentLayout(page(layers, [body.id])),
      0,
      { scale: 1 },
    ).map((run) => run.text)).toEqual(['B', 'I', 'A']);
  });

  it.each([0, -1, Number.NaN])('rejects a non-positive display scale %s', (scale) => {
    const body = paragraph('body', 'body', [placement('B', 0, 0)]);
    const layers = buildPageLayers([
      { layer: 'body', node: body, coordinateSpace: 'section-logical' },
    ]);

    expect(() => textRunsForPage(
      documentLayout(page(layers, [body.id])),
      0,
      { scale },
    )).toThrow(`Text projection scale must be positive: ${scale}`);
  });
});

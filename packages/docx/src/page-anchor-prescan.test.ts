import { describe, expect, it } from 'vitest';
import { __test_preRegisterPageFloats } from './test-support/renderer-internals.test-support.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { isPageLevelAnchorY } from './layout/anchor-classification.js';
import type { AnchorFloatRegistrationState } from './layout/acquisition-context.js';
import type { AnchorAcquisitionInput, AnchorEdgesInput } from './layout/anchor-input.js';
import type { ParagraphLayout } from './layout/types.js';
import type {
  BodyElement,
  DocParagraph,
  DocxDocumentModel,
  DocxTextRun,
  ImageRun,
  SectionProps,
  ShapeRun,
} from './types.js';

// Pins ECMA-376 §20.4.3.2 / §20.4.3.5 behaviour: a `<wp:anchor>` whose
// `<wp:positionV>@relativeFrom` resolves Y independently of the source-order
// anchoring paragraph (page / margin / *Margin / column) is laid out by Word
// as soon as the page is opened, so paragraphs that PRECEDE the anchor's
// paragraph in source order still wrap around the float. Renderer mirrors
// this by pre-scanning page-level floats at every page-start
// (renderBodyElements + paginator's float-reset call sites) and recording
// the source paragraph in `state.pageAnchorPrescanned` so the per-paragraph
// `registerAnchorFloats` skips the duplicate registration when the flow
// reaches it. These three tests cover (1) the wrap-window narrowing seen by
// an EARLIER paragraph, (2) the carve-out for paragraph-local Y (which must
// NOT pre-register), and (3) idempotency (no double FloatRect).

// ===== Stub canvas + helpers (mirrors pagination.test) =====

function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, textAlign: 'left' as CanvasTextAlign,
    direction: 'ltr' as CanvasDirection,
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

function section(overrides: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: 200, pageHeight: 200,
    marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    ...overrides,
  };
}

type DocRun = DocParagraph['runs'][number];

function textRun(text: string, fontSize: number): DocRun {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  };
  return { type: 'text', ...run } as DocRun;
}

function para(opts: { text?: string; fontSize?: number } = {}): BodyElement {
  const fontSize = opts.fontSize ?? 20;
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: opts.text ? [textRun(opts.text, fontSize)] : [],
    defaultFontSize: fontSize, defaultFontFamily: 'NotInMetrics',
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

function paraWith(runs: DocRun[], opts: { fontSize?: number } = {}): BodyElement {
  const fontSize = opts.fontSize ?? 20;
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs,
    defaultFontSize: fontSize, defaultFontFamily: 'NotInMetrics',
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

const missingAnchorEdges = (): AnchorEdgesInput => ({
  topPt: null, topStatus: 'missing',
  rightPt: null, rightStatus: 'missing',
  bottomPt: null, bottomStatus: 'missing',
  leftPt: null, leftStatus: 'missing',
});

const validAnchorEdges = (
  topPt: number,
  rightPt: number,
  bottomPt: number,
  leftPt: number,
): AnchorEdgesInput => ({
  topPt, topStatus: 'valid',
  rightPt, rightStatus: 'valid',
  bottomPt, bottomStatus: 'valid',
  leftPt, leftStatus: 'valid',
});

function parserAnchor(
  occurrenceId: string,
  overrides: Readonly<{
    xPt?: number;
    yPt?: number;
    widthPt?: number;
    heightPt?: number;
    horizontalRelativeFrom?: 'margin' | 'page';
    verticalRelativeFrom?: 'margin' | 'paragraph';
    wrap?: 'square' | 'topAndBottom' | 'none';
    allowOverlap?: boolean;
    relativeHeight?: number;
  }> = {},
): AnchorAcquisitionInput {
  const wrap = overrides.wrap ?? 'square';
  return {
    occurrenceId,
    simplePosition: {
      enabled: false, status: 'valid',
      xPt: 0, xStatus: 'valid', yPt: 0, yStatus: 'valid',
    },
    horizontal: {
      relativeFrom: overrides.horizontalRelativeFrom ?? 'margin',
      relativeFromStatus: 'valid',
      choice: { kind: 'offset', valuePt: overrides.xPt ?? 80 },
    },
    vertical: {
      relativeFrom: overrides.verticalRelativeFrom ?? 'margin',
      relativeFromStatus: 'valid',
      choice: { kind: 'offset', valuePt: overrides.yPt ?? 0 },
    },
    extent: {
      widthPt: overrides.widthPt ?? 80,
      heightPt: overrides.heightPt ?? 60,
      widthStatus: 'valid', heightStatus: 'valid',
    },
    parentEffectExtent: missingAnchorEdges(),
    anchorDistances: missingAnchorEdges(),
    relativeSize: { horizontal: null, vertical: null },
    wrap: {
      kind: wrap,
      authoredKinds: [wrap === 'square'
        ? 'wrapSquare'
        : wrap === 'topAndBottom' ? 'wrapTopAndBottom' : 'wrapNone'],
      side: wrap === 'square' ? 'bothSides' : null,
      distances: missingAnchorEdges(), effectExtent: null, polygon: null,
    },
    behavior: {
      behindDoc: false, behindDocStatus: 'valid',
      relativeHeight: overrides.relativeHeight ?? 1, relativeHeightStatus: 'valid',
      locked: false, lockedStatus: 'valid',
      allowOverlap: overrides.allowOverlap ?? true, allowOverlapStatus: 'valid',
      layoutInCell: true, layoutInCellStatus: 'valid',
    },
    group: null,
  };
}

function parserAnchoredImage(acquisition: AnchorAcquisitionInput): DocRun[] {
  return [
    {
      type: 'anchorHost', fontSize: 20,
      __anchorOccurrenceId: acquisition.occurrenceId,
    } as unknown as DocRun,
    {
      type: 'image', imagePath: 'word/media/canonical.png', mimeType: 'image/png',
      widthPt: acquisition.extent.widthPt ?? 1,
      heightPt: acquisition.extent.heightPt ?? 1,
      anchor: true,
      __anchorAcquisition: acquisition,
    } as unknown as DocRun,
  ];
}

function canonicalModel(body: BodyElement[]): DocxDocumentModel {
  return {
    section: section({ sectionStart: 'nextPage', columns: null }),
    body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    footnotes: [], endnotes: [], fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
}

function canonicalLayout(body: BodyElement[], sectionOverrides: Partial<SectionProps> = {}) {
  const model = {
    ...canonicalModel(body),
    section: section({ sectionStart: 'nextPage', columns: null, ...sectionOverrides }),
  };
  return layoutDocument(
    model,
    createLayoutServices(model, { measureContext: makeCtx() }),
    { currentDateMs: 0 },
  );
}

function sourceParagraphs(
  layout: ReturnType<typeof canonicalLayout>,
  bodyIndex: number,
): ParagraphLayout[] {
  return layout.pages.flatMap((page) => page.layers.body.filter(
    (node): node is ParagraphLayout => node.kind === 'paragraph'
      && node.source.story === 'body'
      && node.source.storyInstance === 'body'
      && node.source.path[0] === bodyIndex,
  ));
}

function retainedText(paragraphs: readonly ParagraphLayout[]): string {
  return paragraphs.flatMap((paragraph) => paragraph.lines)
    .flatMap((line) => line.placements)
    .filter((placement) => placement.kind === 'text')
    .map((placement) => placement.text)
    .join('');
}

// Minimal page-level square-anchored image float (positionV relativeFrom=margin
// + align=top, square wrap). 100×40 pt, pinned at the right edge of the top
// margin so anything on the page from y∈[20,60], x∈[120,220] wraps around it.
function pageImageRun(opts: {
  widthPt?: number;
  heightPt?: number;
  anchorXPt?: number;
  anchorYPt?: number;
  anchorXRelativeFrom?: string | null;
  anchorYRelativeFrom?: string | null;
  anchorXFromMargin?: boolean;
  anchorYFromPara?: boolean;
  wrapMode?: string;
  wrapSide?: string;
} = {}): DocRun {
  const img: ImageRun = {
    imagePath: 'word/media/test1.png',
    mimeType: 'image/png',
    widthPt: opts.widthPt ?? 60,
    heightPt: opts.heightPt ?? 40,
    anchor: true,
    anchorXPt: opts.anchorXPt ?? 100,
    anchorYPt: opts.anchorYPt ?? 0,
    anchorXFromMargin: opts.anchorXFromMargin ?? true,
    anchorYFromPara: opts.anchorYFromPara ?? false,
    wrapMode: opts.wrapMode ?? 'square',
    wrapSide: opts.wrapSide ?? 'bothSides',
    anchorXRelativeFrom: opts.anchorXRelativeFrom ?? 'margin',
    anchorYRelativeFrom: opts.anchorYRelativeFrom ?? 'margin',
  };
  return { type: 'image', ...img } as DocRun;
}

function pageChartRun(opts: {
  anchorXPt?: number;
  anchorYPt?: number;
  widthPt?: number;
  heightPt?: number;
} = {}): DocRun {
  return {
    type: 'chart', chart: {}, anchor: true,
    widthPt: opts.widthPt ?? 40, heightPt: opts.heightPt ?? 40,
    anchorXPt: opts.anchorXPt ?? 0, anchorYPt: opts.anchorYPt ?? 0,
    anchorXRelativeFrom: 'margin', anchorYRelativeFrom: 'margin',
    wrapMode: 'square', wrapSide: 'bothSides',
  } as unknown as DocRun;
}

// Minimal page-level wrap shape (positionV relativeFrom=page + topAndBottom).
function pageShapeRun(opts: {
  widthPt?: number;
  heightPt?: number;
  anchorYPt?: number;
  anchorYRelativeFrom?: string | null;
  anchorYFromPara?: boolean;
  wrapMode?: string | null;
} = {}): DocRun {
  const s: ShapeRun = {
    widthPt: opts.widthPt ?? 160,
    heightPt: opts.heightPt ?? 30,
    anchorXPt: 0,
    anchorYPt: opts.anchorYPt ?? 20,
    anchorXFromMargin: true,
    anchorYFromPara: opts.anchorYFromPara ?? false,
    zOrder: 0,
    subpaths: [],
    presetGeometry: 'rect',
    fill: { fillType: 'solid', color: 'FFFFFF' },
    stroke: null,
    wrapMode: opts.wrapMode === undefined ? 'topAndBottom' : opts.wrapMode,
    wrapSide: null,
    distTop: 0, distBottom: 0, distLeft: 0, distRight: 0,
    anchorYRelativeFrom: opts.anchorYRelativeFrom ?? 'page',
  } as ShapeRun;
  return { type: 'shape', ...s } as DocRun;
}

// ===== Tests =====

describe('preRegisterPageFloats — isPageLevelAnchorY classifier (§20.4.3.5)', () => {
  it('paragraph/line/character ⇒ paragraph-local (NOT page-level)', () => {
    expect(isPageLevelAnchorY('paragraph', false)).toBe(false);
    expect(isPageLevelAnchorY('line', false)).toBe(false);
    expect(isPageLevelAnchorY('character', false)).toBe(false);
  });

  it('page / margin / *Margin / column ⇒ page-level', () => {
    for (const rf of ['page', 'margin', 'topMargin', 'bottomMargin', 'leftMargin', 'rightMargin', 'insideMargin', 'outsideMargin', 'column']) {
      expect(isPageLevelAnchorY(rf, false)).toBe(true);
    }
  });

  it('absent relativeFrom defers to anchorYFromPara (page-level only when NOT from-para)', () => {
    expect(isPageLevelAnchorY(null, false)).toBe(true);
    expect(isPageLevelAnchorY(null, true)).toBe(false);
    expect(isPageLevelAnchorY(undefined, false)).toBe(true);
  });
});

describe('canonical page-owned anchor prescan (§20.4.2.3/.17/.20)', () => {
  it('preserves page-owned wrapping for a hand-built public image anchor', () => {
    const layout = canonicalLayout([
      para({ text: 'あ'.repeat(40), fontSize: 20 }),
      paraWith([pageImageRun({
        widthPt: 80,
        heightPt: 60,
        anchorXPt: 80,
        anchorYPt: 0,
        anchorXRelativeFrom: 'margin',
        anchorYRelativeFrom: 'margin',
        anchorYFromPara: false,
        wrapMode: 'square',
      })]),
    ]);
    const earlier = sourceParagraphs(layout, 0);

    expect(earlier[0]!.exclusions).toEqual([
      expect.objectContaining({
        wrap: 'square',
        anchorOccurrenceId: expect.stringContaining(
          encodeURIComponent('public-anchor:body:body:1:0'),
        ),
      }),
    ]);
  });

  it('prescans public image, chart, and shape occurrences without collapsing identity', () => {
    const layout = canonicalLayout([
      para({ text: 'あ'.repeat(8), fontSize: 20 }),
      paraWith([
        pageImageRun({ widthPt: 40, heightPt: 40, anchorXPt: 0, anchorYPt: 0 }),
        pageChartRun({ widthPt: 40, heightPt: 40, anchorXPt: 50, anchorYPt: 0 }),
        pageShapeRun({ widthPt: 40, heightPt: 40, anchorYPt: 0 }),
      ]),
    ]);
    const occurrenceIds = sourceParagraphs(layout, 0)[0]!.exclusions
      .map((exclusion) => exclusion.anchorOccurrenceId ?? '');

    expect(occurrenceIds).toHaveLength(3);
    expect(occurrenceIds).toEqual(expect.arrayContaining([
      expect.stringContaining(encodeURIComponent('public-anchor:body:body:1:0')),
      expect.stringContaining(encodeURIComponent('public-anchor:body:body:1:1')),
      expect.stringContaining(encodeURIComponent('public-shape:body:body:1:2')),
    ]));
  });

  it('acquires a later page-owned occurrence from parser facts before paragraph zero', () => {
    const text = 'あ'.repeat(40);
    const pageOwned = parserAnchor('page-owned');
    const paragraphLocal = parserAnchor('paragraph-local', {
      verticalRelativeFrom: 'paragraph',
    });
    const pageLayout = canonicalLayout([
      para({ text, fontSize: 20 }),
      paraWith(parserAnchoredImage(pageOwned)),
    ]);
    const localLayout = canonicalLayout([
      para({ text, fontSize: 20 }),
      paraWith(parserAnchoredImage(paragraphLocal)),
    ]);
    const pageParagraphs = sourceParagraphs(pageLayout, 0);
    const localParagraphs = sourceParagraphs(localLayout, 0);

    expect(pageParagraphs.flatMap((paragraph) => paragraph.lines).length)
      .toBeGreaterThan(localParagraphs.flatMap((paragraph) => paragraph.lines).length);
    expect(retainedText(pageParagraphs)).toBe(text);
    expect(retainedText(localParagraphs)).toBe(text);
    expect(pageParagraphs[0]!.exclusions).toHaveLength(1);
    expect(pageParagraphs[0]!.exclusions[0]!.wrap).toBe('square');
    expect(localParagraphs[0]!.exclusions).toHaveLength(0);
  });

  it('projects a tbRl page-owned anchor into the section-logical wrap frame', () => {
    const physicalAnchor = parserAnchor('vertical-page-owned', {
      xPt: 220,
      yPt: 50,
      widthPt: 80,
      heightPt: 30,
      horizontalRelativeFrom: 'page',
      verticalRelativeFrom: 'margin',
      wrap: 'square',
    });
    const layout = canonicalLayout([
      para({ text: 'あ'.repeat(20), fontSize: 20 }),
      paraWith(parserAnchoredImage(physicalAnchor)),
    ], {
      pageWidth: 300,
      pageHeight: 200,
      marginTop: 0,
      marginRight: 0,
      marginBottom: 0,
      marginLeft: 0,
      textDirection: 'tbRl',
    });
    const control = canonicalLayout([
      para({ text: 'あ'.repeat(20), fontSize: 20 }),
      para(),
    ], {
      pageWidth: 300,
      pageHeight: 200,
      marginTop: 0,
      marginRight: 0,
      marginBottom: 0,
      marginLeft: 0,
      textDirection: 'tbRl',
    });
    const earlier = sourceParagraphs(layout, 0)[0]!;
    const controlEarlier = sourceParagraphs(control, 0)[0]!;
    const owner = sourceParagraphs(layout, 1)[0]!;
    const drawing = owner.drawings[0]!;

    // ECMA-376 §17.6.20 + §§20.4.2.3/.7/.10/.11: wp:positionH/V and
    // wp:extent resolve in the upright physical 300×200 page. Retained vertical
    // body flow uses the inverse quarter-turn:
    //   physical = (pageWidth - logical.y, logical.x)
    // so physical {220,50,80×30} becomes logical {50,0,30×80}.
    expect(earlier.exclusions).toEqual([
      expect.objectContaining({
        wrap: 'square',
        bounds: { xPt: 50, yPt: 0, widthPt: 30, heightPt: 80 },
      }),
    ]);
    expect(drawing.flowBounds).toEqual({ xPt: 50, yPt: 0, widthPt: 30, heightPt: 80 });
    expect(drawing.transform).toEqual({ a: 0, b: -1, c: 1, d: 0, e: 65, f: 40 });
    expect(owner.exclusions[0]?.bounds).toEqual(drawing.flowBounds);
    expect(earlier.lines.length).toBeGreaterThan(controlEarlier.lines.length);
  });

  it('projects a tbLrV page-owned anchor with the vertical-lr inverse', () => {
    const physicalAnchor = parserAnchor('vertical-lr-page-owned', {
      xPt: 220,
      yPt: 50,
      widthPt: 80,
      heightPt: 30,
      horizontalRelativeFrom: 'page',
      verticalRelativeFrom: 'margin',
      wrap: 'square',
    });
    const layout = canonicalLayout([
      para({ text: 'あ'.repeat(20), fontSize: 20 }),
      paraWith(parserAnchoredImage(physicalAnchor)),
    ], {
      pageWidth: 300,
      pageHeight: 200,
      marginTop: 0,
      marginRight: 0,
      marginBottom: 0,
      marginLeft: 0,
      textDirection: 'tbLrV',
    });
    const earlier = sourceParagraphs(layout, 0)[0]!;
    const owner = sourceParagraphs(layout, 1)[0]!;
    const drawing = owner.drawings[0]!;

    // ECMA-376 Part 4 §14.11.7 makes tbLrV semantically equivalent to lrV.
    // Its physical-to-logical mapping is (x, y) -> (y, x), so the physical
    // {220,50,80×30} box becomes logical {50,220,30×80}.
    expect(earlier.exclusions).toEqual([
      expect.objectContaining({
        wrap: 'square',
        bounds: { xPt: 50, yPt: 220, widthPt: 30, heightPt: 80 },
      }),
    ]);
    expect(drawing.flowBounds).toEqual({ xPt: 50, yPt: 220, widthPt: 30, heightPt: 80 });
    expect(drawing.transform).toEqual({ a: 0, b: 1, c: 1, d: 0, e: 65, f: 260 });
    expect(owner.exclusions[0]?.bounds).toEqual(drawing.flowBounds);
  });

  it('projects tbLrV prescan edges and polygons through the same affine transform', () => {
    const base = parserAnchor('vertical-lr-polygon', {
      xPt: 220,
      yPt: 50,
      widthPt: 80,
      heightPt: 30,
      horizontalRelativeFrom: 'page',
      verticalRelativeFrom: 'margin',
      wrap: 'square',
    });
    const physicalAnchor: AnchorAcquisitionInput = {
      ...base,
      anchorDistances: validAnchorEdges(1, 2, 3, 4),
      wrap: {
        ...base.wrap,
        kind: 'through',
        authoredKinds: ['wrapThrough'],
        polygon: {
          edited: false,
          coordinateSpace: { width: 21600, height: 21600 },
          points: [
            { x: 0, y: 0, rawX: '0', rawY: '0' },
            { x: 21600, y: 0, rawX: '21600', rawY: '0' },
            { x: 0, y: 21600, rawX: '0', rawY: '21600' },
          ],
          invalidPointCount: 0,
        },
      },
    };
    const layout = canonicalLayout([
      para({ text: 'あ'.repeat(20), fontSize: 20 }),
      paraWith(parserAnchoredImage(physicalAnchor)),
    ], {
      pageWidth: 300,
      pageHeight: 200,
      marginTop: 0,
      marginRight: 0,
      marginBottom: 0,
      marginLeft: 0,
      textDirection: 'tbLrV',
    });
    const exclusion = sourceParagraphs(layout, 0)[0]!.exclusions[0]!;

    expect(exclusion.bounds).toEqual({
      xPt: 49,
      yPt: 216,
      widthPt: 34,
      heightPt: 86,
    });
    expect(exclusion.polygon).toEqual([
      { xPt: 50, yPt: 220 },
      { xPt: 50, yPt: 300 },
      { xPt: 80, yPt: 220 },
    ]);
  });

  it('keeps two occurrences in one paragraph distinct and preserves authored wrap kinds', () => {
    const square = parserAnchor('square', {
      xPt: 80, yPt: 0, widthPt: 80, heightPt: 40, wrap: 'square',
    });
    const topAndBottom = parserAnchor('top-bottom', {
      xPt: 0, yPt: 80, widthPt: 160, heightPt: 20, wrap: 'topAndBottom',
    });
    const layout = canonicalLayout([
      para({ text: 'あ'.repeat(24), fontSize: 20 }),
      paraWith([
        ...parserAnchoredImage(square),
        ...parserAnchoredImage(topAndBottom),
      ]),
    ]);
    const first = sourceParagraphs(layout, 0)[0]!;
    const canonical = first.exclusions;

    expect(canonical.map((exclusion) => ({
      wrap: exclusion.wrap,
    }))).toEqual([
      { wrap: 'square' },
      { wrap: 'topAndBottom' },
    ]);
    expect(new Set(canonical.map((exclusion) => exclusion.anchorOccurrenceId)).size).toBe(2);
  });

  it('preserves an authored four-anchor composition within one source paragraph', () => {
    // Word preserves this authored composition as one title/metadata block.
    // The leading host-owned object overlaps the full-width page-owned object
    // by 0.30pt; the two page-owned objects otherwise abut exactly. Resolving
    // all four as independent CT_Anchor collisions cascades the full-width
    // object below the second page-owned object instead of preserving the
    // authored composition.
    const leadingHostOwned = parserAnchor('composition-leading-host', {
      xPt: 0.3,
      yPt: -28.2,
      widthPt: 285.85,
      heightPt: 28.5,
      verticalRelativeFrom: 'paragraph',
      wrap: 'none',
      allowOverlap: false,
      relativeHeight: 251658240,
    });
    const trailingPageOwned = parserAnchor('composition-trailing-page', {
      xPt: 36.1,
      yPt: 102.05,
      widthPt: 402.95,
      heightPt: 75.2,
      wrap: 'topAndBottom',
      allowOverlap: false,
      relativeHeight: 251657216,
    });
    const trailingHostOwned = parserAnchor('composition-trailing-host', {
      xPt: 246.15,
      yPt: 177.85,
      widthPt: 237.4,
      heightPt: 121.45,
      verticalRelativeFrom: 'paragraph',
      wrap: 'square',
      allowOverlap: true,
      relativeHeight: 251659264,
    });
    const leadingPageOwned = parserAnchor('composition-leading-page', {
      xPt: 0,
      yPt: 0,
      widthPt: 481.9,
      heightPt: 102.05,
      wrap: 'topAndBottom',
      allowOverlap: false,
      relativeHeight: 251656192,
    });
    const layout = canonicalLayout([
      paraWith([
        ...parserAnchoredImage(leadingHostOwned),
        ...parserAnchoredImage(trailingPageOwned),
        ...parserAnchoredImage(trailingHostOwned),
        ...parserAnchoredImage(leadingPageOwned),
      ]),
    ], {
      pageWidth: 612,
      pageHeight: 792,
      marginTop: 56.7,
      marginRight: 73.4,
      marginBottom: 56.7,
      marginLeft: 56.7,
    });
    const drawings = sourceParagraphs(layout, 0).flatMap((paragraph) => paragraph.drawings);
    const bounds = (suffix: string) => drawings.find((drawing) =>
      drawing.anchorLayer?.occurrenceId.endsWith(suffix))?.flowBounds;
    const expectVerticalBounds = (suffix: string, yPt: number, heightPt: number) => {
      const actual = bounds(suffix);
      expect(actual).toBeDefined();
      expect(actual!.yPt).toBeCloseTo(yPt, 8);
      expect(actual!.heightPt).toBeCloseTo(heightPt, 8);
    };

    expectVerticalBounds('composition-leading-host', 28.5, 28.5);
    expectVerticalBounds('composition-leading-page', 56.7, 102.05);
    expectVerticalBounds('composition-trailing-page', 158.75, 75.2);
    expectVerticalBounds('composition-trailing-host', 234.55, 121.45);
  });

  it('does not displace an overlap-allowed anchor by its own paragraph\'s prescanned siblings', () => {
    // Both page-owned square anchors sit in one paragraph and allow overlap
    // (§20.4.2.3 allowOverlap=true). Prescan registers each one on the page
    // before the paragraph lays out; that registration must not turn a
    // same-paragraph sibling into a different-paragraph wrap blocker.
    const upper = parserAnchor('same-paragraph-upper', {
      xPt: 0, yPt: 0, widthPt: 160, heightPt: 60, wrap: 'square',
    });
    const lower = parserAnchor('same-paragraph-lower', {
      xPt: 0, yPt: 40, widthPt: 160, heightPt: 40, wrap: 'square',
    });
    const layout = canonicalLayout([
      paraWith([...parserAnchoredImage(upper), ...parserAnchoredImage(lower)]),
    ]);
    const drawings = sourceParagraphs(layout, 0).flatMap((paragraph) => paragraph.drawings);
    const bounds = (suffix: string) => drawings.find((drawing) =>
      drawing.anchorLayer?.occurrenceId.endsWith(suffix))?.flowBounds;
    const upperBounds = bounds('same-paragraph-upper');
    const lowerBounds = bounds('same-paragraph-lower');

    expect(upperBounds).toBeDefined();
    expect(lowerBounds).toBeDefined();
    expect(lowerBounds!.yPt - upperBounds!.yPt).toBe(40);
  });

  it('does not let a future prescanned text float displace an earlier object', () => {
    const earlier = parserAnchor('earlier-object', {
      xPt: 80, yPt: 0, widthPt: 80, heightPt: 60,
      verticalRelativeFrom: 'paragraph', wrap: 'none', allowOverlap: false,
    });
    const future = parserAnchor('future-wrapped-object', {
      xPt: 80, yPt: 0, widthPt: 80, heightPt: 60,
      wrap: 'square',
    });
    const text = 'あ'.repeat(32);
    const baseline = canonicalLayout([
      paraWith([...parserAnchoredImage(earlier), textRun(text, 20)]),
    ]);
    const withFuture = canonicalLayout([
      paraWith([...parserAnchoredImage(earlier), textRun(text, 20)]),
      paraWith(parserAnchoredImage(future)),
    ]);
    const baselineEarlier = sourceParagraphs(baseline, 0)[0]!;
    const prescannedEarlier = sourceParagraphs(withFuture, 0)[0]!;

    expect(prescannedEarlier.drawings[0]?.flowBounds)
      .toEqual(baselineEarlier.drawings[0]?.flowBounds);
    expect(prescannedEarlier.lines.length).toBeGreaterThan(baselineEarlier.lines.length);
    expect(prescannedEarlier.exclusions).toEqual([
      expect.objectContaining({
        anchorOccurrenceId: expect.stringContaining('future-wrapped-object'),
      }),
    ]);
  });

  it('starts page-owned object collision authority only after source acceptance', () => {
    const future = parserAnchor('accepted-page-object', {
      xPt: 80, yPt: 0, widthPt: 80, heightPt: 100,
      wrap: 'square',
    });
    const later = parserAnchor('later-object', {
      xPt: 80, yPt: 0, widthPt: 80, heightPt: 20,
      verticalRelativeFrom: 'paragraph', wrap: 'none', allowOverlap: false,
    });
    const layout = canonicalLayout([
      para({ text: '前', fontSize: 20 }),
      paraWith(parserAnchoredImage(future)),
      paraWith(parserAnchoredImage(later)),
    ]);
    const blocker = sourceParagraphs(layout, 1)
      .flatMap((paragraph) => paragraph.drawings)
      .find((drawing) => drawing.anchorLayer?.occurrenceId.endsWith('accepted-page-object'));
    const moving = sourceParagraphs(layout, 2)
      .flatMap((paragraph) => paragraph.drawings)
      .find((drawing) => drawing.anchorLayer?.occurrenceId.endsWith('later-object'));

    expect(blocker).toBeDefined();
    expect(moving).toBeDefined();
    expect(moving!.flowBounds.yPt).toBe(
      blocker!.flowBounds.yPt + blocker!.flowBounds.heightPt,
    );
  });
});

describe('canonical paragraph-owned anchor registry (§20.4.2.3/.20 + §20.4.3.5)', () => {
  it('clears accepted object collisions at a physical page transition', () => {
    const firstPage = parserAnchor('first-page-wrap-none', {
      xPt: 0, yPt: 0, widthPt: 80, heightPt: 60,
      verticalRelativeFrom: 'paragraph', wrap: 'none',
    });
    const secondPage = parserAnchor('second-page-moving', {
      xPt: 0, yPt: 0, widthPt: 80, heightPt: 40,
      verticalRelativeFrom: 'paragraph', wrap: 'none', allowOverlap: false,
    });
    const baseline = canonicalLayout([
      paraWith(parserAnchoredImage(secondPage)),
    ]);
    const paginated = canonicalLayout([
      paraWith(parserAnchoredImage(firstPage)),
      { type: 'pageBreak' } as BodyElement,
      paraWith(parserAnchoredImage(secondPage)),
    ]);
    const baselineBounds = sourceParagraphs(baseline, 0)[0]!.drawings[0]!.flowBounds;
    const secondPageBounds = sourceParagraphs(paginated, 2)[0]!.drawings[0]!.flowBounds;

    expect(secondPageBounds.xPt).toBe(baselineBounds.xPt);
    expect(secondPageBounds.yPt).toBe(baselineBounds.yPt);
  });

  it('paginates parser-owned square-anchor text from the final reflowed line boundaries', () => {
    const occurrenceId = 'same-paragraph-square-pagination';
    const square = parserAnchor(occurrenceId, {
      xPt: 100, yPt: 0, widthPt: 40, heightPt: 60,
      verticalRelativeFrom: 'paragraph', wrap: 'square',
    });
    const none = parserAnchor(occurrenceId, {
      xPt: 100, yPt: 0, widthPt: 40, heightPt: 60,
      verticalRelativeFrom: 'paragraph', wrap: 'none',
    });
    const text = 'あ'.repeat(80);
    const acquire = (anchor: AnchorAcquisitionInput) => canonicalLayout([
      paraWith([...parserAnchoredImage(anchor), textRun(text, 20)]),
    ]);

    const wrappedLayout = acquire(square);
    const noneLayout = acquire(none);
    const wrapped = sourceParagraphs(wrappedLayout, 0);
    const unwrapped = sourceParagraphs(noneLayout, 0);

    expect(wrapped.flatMap((paragraph) => paragraph.lines).length)
      .toBeGreaterThan(unwrapped.flatMap((paragraph) => paragraph.lines).length);
    expect(retainedText(wrapped)).toBe(text);
    expect(wrapped.flatMap((paragraph) => paragraph.drawings)).toHaveLength(1);
    expect(wrapped.flatMap((paragraph) => paragraph.exclusions)).toHaveLength(1);
    expect(wrapped.every((paragraph, index) => index === 0
      ? paragraph.continuation?.continuesFromPrevious !== true
      : paragraph.continuation?.continuesFromPrevious === true)).toBe(true);
    expect(wrapped.length).toBeGreaterThan(1);
  });

  it('commits a parser-owned topAndBottom exclusion for following paragraphs', () => {
    const paragraphOwned = parserAnchor('paragraph-owned-top-bottom', {
      xPt: 0,
      yPt: 0,
      widthPt: 160,
      heightPt: 60,
      verticalRelativeFrom: 'paragraph',
      wrap: 'topAndBottom',
    });
    const layout = canonicalLayout([
      paraWith(parserAnchoredImage(paragraphOwned)),
      para({ text: 'あ'.repeat(48), fontSize: 20 }),
    ]);
    const following = sourceParagraphs(layout, 1);

    expect(layout.pages).toHaveLength(2);
    expect(following).toHaveLength(2);
    expect(following[0]!.exclusions).toEqual([
      expect.objectContaining({
        wrap: 'topAndBottom',
        anchorOccurrenceId: expect.stringContaining(
          encodeURIComponent('anchor:body:body:0:paragraph-owned-top-bottom'),
        ),
      }),
    ]);
  });

  it('commits a parser-owned anchor only when its continuation is accepted', () => {
    const occurrenceId = 'later-continuation-top-bottom';
    const paragraphOwned = parserAnchor(occurrenceId, {
      xPt: 0,
      yPt: 0,
      widthPt: 160,
      heightPt: 60,
      verticalRelativeFrom: 'paragraph',
      wrap: 'topAndBottom',
    });
    const layout = canonicalLayout([
      paraWith([
        textRun('あ'.repeat(80), 20),
        ...parserAnchoredImage(paragraphOwned),
      ]),
      para({ text: 'い'.repeat(48), fontSize: 20 }),
    ]);
    const anchored = sourceParagraphs(layout, 0);
    const following = sourceParagraphs(layout, 1);
    const ownsOccurrence = (paragraph: ParagraphLayout) => paragraph.drawings.some((drawing) =>
      drawing.anchorLayer?.acquisitionOccurrenceId?.endsWith(occurrenceId));
    const carriesOccurrenceExclusion = (paragraph: ParagraphLayout) => paragraph.exclusions.some(
      (exclusion) => exclusion.anchorOccurrenceId?.endsWith(occurrenceId),
    );

    expect(anchored).toHaveLength(2);
    expect(ownsOccurrence(anchored[0]!)).toBe(false);
    expect(carriesOccurrenceExclusion(anchored[0]!)).toBe(false);
    expect(ownsOccurrence(anchored[1]!)).toBe(true);
    expect(carriesOccurrenceExclusion(anchored[1]!)).toBe(true);
    expect(following).toHaveLength(2);
    expect(layout.pages).toHaveLength(3);
  });

  it('measures an anchored keepNext successor through the same reflow acquisition', () => {
    const paragraphOwned = parserAnchor('keep-next-successor-square', {
      xPt: 0,
      yPt: 0,
      widthPt: 160,
      heightPt: 60,
      verticalRelativeFrom: 'paragraph',
      wrap: 'topAndBottom',
    });
    const kept = {
      ...para({ text: 'keep', fontSize: 20 }),
      keepNext: true,
    } as BodyElement;
    const anchoredSuccessor = {
      ...paraWith([
        ...parserAnchoredImage(paragraphOwned),
        textRun('後'.repeat(8), 20),
      ]),
      keepNext: true,
    } as BodyElement;
    const layout = canonicalLayout([
      para({ text: '前'.repeat(40), fontSize: 20 }),
      kept,
      anchoredSuccessor,
      para({ text: '終', fontSize: 20 }),
    ]);
    const sourcePages = (bodyIndex: number) => layout.pages.flatMap((page, pageIndex) =>
      page.layers.body.some((node) => node.kind === 'paragraph'
        && node.source.story === 'body'
        && node.source.path[0] === bodyIndex)
        ? [pageIndex]
        : []);

    expect(sourcePages(1)).toEqual([1]);
    expect(sourcePages(2)).toEqual([1]);
    expect(sourcePages(3)).toEqual([1]);
  });

  it('includes anchor displacement in the terminal successor lead extent', () => {
    const paragraphOwned = parserAnchor('keep-next-terminal-top-bottom', {
      xPt: 0,
      yPt: 0,
      widthPt: 160,
      heightPt: 60,
      verticalRelativeFrom: 'paragraph',
      wrap: 'topAndBottom',
    });
    const kept = {
      ...para({ text: 'keep', fontSize: 20 }),
      keepNext: true,
    } as BodyElement;
    const layout = canonicalLayout([
      para({ text: '前'.repeat(40), fontSize: 20 }),
      kept,
      paraWith([
        ...parserAnchoredImage(paragraphOwned),
        textRun('後'.repeat(8), 20),
      ]),
    ]);
    const sourcePages = (bodyIndex: number) => layout.pages.flatMap((page, pageIndex) =>
      page.layers.body.some((node) => node.kind === 'paragraph'
        && node.source.story === 'body'
        && node.source.path[0] === bodyIndex)
        ? [pageIndex]
        : []);

    expect(sourcePages(1)).toEqual([1]);
    expect(sourcePages(2)).toEqual([1]);
  });

});

describe('preRegisterPageFloats — paragraph-local Y is NOT pre-registered', () => {
  it('pre-scan ignores a wrap float with positionV relativeFrom="paragraph"', () => {
    const body = [
      para({ text: 'ABC' }),
      paraWith([
        // Paragraph-local: must NOT be pre-registered (paraId stays at the
        // paragraph's own registration in registerAnchorFloats).
        pageShapeRun({ widthPt: 100, heightPt: 30, anchorYRelativeFrom: 'paragraph', anchorYFromPara: true, wrapMode: 'topAndBottom' }),
      ]),
    ];
    // Minimal acquisition-state stub matching what registerImageFloat/registerShapeFloat
    // touch point-space margins/page/column geometry and the float registry.
    const state = {
      marginLeft: 20, marginRight: 20, marginTop: 20, marginBottom: 20,
      pageWidth: 200, pageH: 200,
      contentX: 20, contentW: 160,
      floats: [],
      floatParaSeq: 0,
      pageAnchorPrescanned: new Set(),
    } as unknown as AnchorFloatRegistrationState;
    __test_preRegisterPageFloats(body, 0, state);
    expect(state.floats.length).toBe(0);
    expect(state.pageAnchorPrescanned?.size ?? 0).toBe(0);
  });

  it('pre-scan REGISTERS a page-level (relativeFrom="margin") wrap float on an earlier-scanned paragraph', () => {
    const body = [
      para({ text: 'ABC' }), // paragraph A, no floats
      paraWith([
        // Paragraph B carries a page-level square anchor
        pageImageRun({ widthPt: 60, heightPt: 40, anchorYRelativeFrom: 'margin', anchorYFromPara: false, wrapMode: 'square' }),
      ]),
    ];
    const state = {
      marginLeft: 20, marginRight: 20, marginTop: 20, marginBottom: 20,
      pageWidth: 200, pageH: 200,
      contentX: 20, contentW: 160,
      floats: [],
      floatParaSeq: 0,
      pageAnchorPrescanned: new Set(),
    } as unknown as AnchorFloatRegistrationState;
    __test_preRegisterPageFloats(body, 0, state);
    expect(state.floats.length).toBe(1);
    expect(state.pageAnchorPrescanned?.size).toBe(1);
  });
});

describe('preRegisterPageFloats — idempotent on repeated pre-scan', () => {
  it('a paragraph already in pageAnchorPrescanned is not re-registered', () => {
    const paraB = paraWith([
      pageImageRun({ widthPt: 60, heightPt: 40, anchorYRelativeFrom: 'margin', wrapMode: 'square' }),
    ]);
    const body = [para({ text: 'A' }), paraB];
    const state = {
      marginLeft: 20, marginRight: 20, marginTop: 20, marginBottom: 20,
      pageWidth: 200, pageH: 200,
      contentX: 20, contentW: 160,
      floats: [],
      floatParaSeq: 0,
      pageAnchorPrescanned: new Set(),
    } as unknown as AnchorFloatRegistrationState;
    __test_preRegisterPageFloats(body, 0, state);
    const after1 = state.floats.length;
    expect(after1).toBe(1);
    // Second call: same body, same start index. The dedupe key is the
    // already-pre-registered paragraph in pageAnchorPrescanned, so floats
    // must NOT grow.
    __test_preRegisterPageFloats(body, 0, state);
    expect(state.floats.length).toBe(after1);
  });

  it('stops at the next pageBreak — items after the break are NOT pre-scanned for this page', () => {
    const body = [
      paraWith([pageImageRun({ widthPt: 60, heightPt: 40, anchorYRelativeFrom: 'margin', wrapMode: 'square' })]), // would register
      { type: 'pageBreak' } as BodyElement,                                                                       // page boundary
      paraWith([pageImageRun({ widthPt: 60, heightPt: 40, anchorYRelativeFrom: 'margin', wrapMode: 'square' })]), // must be SKIPPED — belongs to the next page
    ];
    const state = {
      marginLeft: 20, marginRight: 20, marginTop: 20, marginBottom: 20,
      pageWidth: 200, pageH: 200,
      contentX: 20, contentW: 160,
      floats: [],
      floatParaSeq: 0,
      pageAnchorPrescanned: new Set(),
    } as unknown as AnchorFloatRegistrationState;
    __test_preRegisterPageFloats(body, 0, state);
    expect(state.floats.length).toBe(1);
    expect(state.pageAnchorPrescanned?.size).toBe(1);
  });
});

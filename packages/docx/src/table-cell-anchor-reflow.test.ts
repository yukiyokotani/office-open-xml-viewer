import { describe, expect, it } from 'vitest';
import type { AnchorAcquisitionInput, AnchorEdgesInput } from './layout/anchor-input.js';
import type { ParagraphLayout } from './layout/types.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import type {
  BodyElement,
  DocParagraph,
  DocTable,
  DocTableCell,
  DocTableRow,
  DocxDocumentModel,
  DocxTextRun,
  FieldRun,
  SectionProps,
} from './types.js';

function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText(text: string) {
      const size = px();
      return {
        width: [...text].length * size,
        fontBoundingBoxAscent: size * 0.8,
        fontBoundingBoxDescent: size * 0.2,
        actualBoundingBoxAscent: size * 0.8,
        actualBoundingBoxDescent: size * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign,
    direction: 'ltr' as CanvasDirection,
  } as unknown as CanvasRenderingContext2D;
}

function textRun(text: string): DocParagraph['runs'][number] {
  const run: DocxTextRun = {
    text,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    fontSize: 20,
    color: null,
    fontFamily: 'NotInMetrics',
    isLink: false,
    background: null,
    vertAlign: null,
    hyperlink: null,
  };
  return { type: 'text', ...run } as DocParagraph['runs'][number];
}

function pageFieldRun(): DocParagraph['runs'][number] {
  const run: FieldRun = {
    fieldType: 'page',
    instruction: 'PAGE',
    fallbackText: '?',
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    fontSize: 20,
    color: null,
    fontFamily: 'NotInMetrics',
    background: null,
    vertAlign: null,
  };
  return { type: 'field', ...run } as DocParagraph['runs'][number];
}

function vmlRectRun(horizontalOffsetPt: number): DocParagraph['runs'][number] {
  return {
    type: 'shape', inline: false,
    widthPt: 45, heightPt: 14,
    anchorXPt: horizontalOffsetPt, anchorYPt: 0,
    anchorXFromMargin: true, anchorYFromPara: true,
    anchorXRelativeFrom: 'column', anchorYRelativeFrom: 'paragraph',
    presetGeometry: 'rect', adjValues: [], subpaths: [],
    fill: null, stroke: '000000', strokeWidth: 0.75,
    rotation: 0, flipH: false, flipV: false,
    textBlocks: [], textBoxContent: [],
  } as unknown as DocParagraph['runs'][number];
}

const missingEdges = (): AnchorEdgesInput => ({
  topPt: null, topStatus: 'missing',
  rightPt: null, rightStatus: 'missing',
  bottomPt: null, bottomStatus: 'missing',
  leftPt: null, leftStatus: 'missing',
});

function anchoredImageRuns(options: Readonly<{
  occurrenceId?: string;
  verticalRelativeFrom?: 'line' | 'paragraph' | 'page' | 'margin';
  horizontalRelativeFrom?: 'column' | 'margin';
  horizontalOffsetPt?: number;
  verticalOffsetPt?: number;
  simplePosition?: Readonly<{ xPt: number; yPt: number }>;
  widthPt?: number;
  heightPt?: number;
  allowOverlap?: boolean;
  wrapKind?: 'none' | 'square';
}> = {}): DocParagraph['runs'] {
  const occurrenceId = options.occurrenceId ?? 'cell-later';
  const verticalRelativeFrom = options.verticalRelativeFrom ?? 'line';
  const horizontalRelativeFrom = options.horizontalRelativeFrom ?? 'column';
  const horizontalOffsetPt = options.horizontalOffsetPt ?? 80;
  const verticalOffsetPt = options.verticalOffsetPt ?? 0;
  const widthPt = options.widthPt ?? 80;
  const heightPt = options.heightPt ?? 40;
  const allowOverlap = options.allowOverlap ?? true;
  const wrapKind = options.wrapKind ?? 'square';
  const acquisition: AnchorAcquisitionInput = {
    occurrenceId,
    simplePosition: {
      enabled: options.simplePosition !== undefined, status: 'valid',
      xPt: options.simplePosition?.xPt ?? 0, xStatus: 'valid',
      yPt: options.simplePosition?.yPt ?? 0, yStatus: 'valid',
    },
    horizontal: {
      relativeFrom: horizontalRelativeFrom, relativeFromStatus: 'valid',
      choice: { kind: 'offset', valuePt: horizontalOffsetPt },
    },
    vertical: {
      relativeFrom: verticalRelativeFrom, relativeFromStatus: 'valid',
      choice: { kind: 'offset', valuePt: verticalOffsetPt },
    },
    extent: {
      widthPt, heightPt,
      widthStatus: 'valid', heightStatus: 'valid',
    },
    parentEffectExtent: missingEdges(),
    anchorDistances: missingEdges(),
    relativeSize: { horizontal: null, vertical: null },
    wrap: {
      kind: wrapKind,
      authoredKinds: [wrapKind === 'none' ? 'wrapNone' : 'wrapSquare'],
      side: 'bothSides',
      distances: missingEdges(),
      effectExtent: null,
      polygon: null,
    },
    behavior: {
      behindDoc: false, behindDocStatus: 'valid',
      relativeHeight: 1, relativeHeightStatus: 'valid',
      locked: false, lockedStatus: 'valid',
      allowOverlap, allowOverlapStatus: 'valid',
      layoutInCell: true, layoutInCellStatus: 'valid',
    },
    group: null,
  };
  return [
    {
      type: 'anchorHost',
      fontSize: 20,
      __anchorOccurrenceId: occurrenceId,
    } as unknown as DocParagraph['runs'][number],
    {
      type: 'image',
      imagePath: 'word/media/cell-anchor.png',
      mimeType: 'image/png',
      widthPt,
      heightPt,
      anchor: true,
      anchorXPt: horizontalOffsetPt,
      anchorXRelativeFrom: horizontalRelativeFrom,
      anchorYPt: verticalOffsetPt,
      anchorYRelativeFrom: verticalRelativeFrom,
      anchorYFromPara: true,
      wrapMode: wrapKind,
      wrapSide: 'bothSides',
      layoutInCell: true,
      __anchorAcquisition: acquisition,
    } as unknown as DocParagraph['runs'][number],
  ];
}

function noBorders() {
  return {
    top: null, right: null, bottom: null, left: null,
    insideH: null, insideV: null,
  };
}

function paragraph(runs: DocParagraph['runs']): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [], runs,
    defaultFontSize: 20, defaultFontFamily: 'NotInMetrics',
  } as unknown as DocParagraph;
}

function model(paragraphs: DocParagraph[] = [
  paragraph([
    textRun('A'.repeat(9)),
    ...anchoredImageRuns(),
    textRun('B'.repeat(8)),
  ]),
], cellWidthPt = 160): DocxDocumentModel {
  const cell = {
    content: paragraphs,
    colSpan: 1,
    vMerge: null,
    borders: noBorders(),
    background: null,
    vAlign: 'top',
    widthPt: cellWidthPt,
  } as unknown as DocTableCell;
  const row = {
    cells: [cell],
    rowHeight: null,
    rowHeightRule: 'auto',
    isHeader: false,
  } as unknown as DocTableRow;
  const table = {
    type: 'table',
    colWidths: [cellWidthPt],
    rows: [row],
    borders: noBorders(),
    cellMarginTop: 0,
    cellMarginRight: 0,
    cellMarginBottom: 0,
    cellMarginLeft: 0,
    jc: 'left',
    layout: 'fixed',
  } as unknown as DocTable;
  const section: SectionProps = {
    pageWidth: 200, pageHeight: 200,
    marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
    headerDistance: 0, footerDistance: 0,
    titlePage: false, evenAndOddHeaders: false,
    sectionStart: 'nextPage', columns: null,
  };
  return {
    section,
    body: [table as unknown as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    footnotes: [],
    endnotes: [],
    fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
}

function cellParagraphs(document = model()): ParagraphLayout[] {
  const layout = layoutDocument(
    document,
    createLayoutServices(document, { measureContext: makeCtx() }),
    { currentDateMs: 0 },
  );
  const table = layout.pages[0]!.layers.body.find((node) => node.kind === 'table');
  if (!table || table.kind !== 'table') throw new Error('Expected retained table');
  return table.rows[0]?.cells[0]?.blocks.flatMap((block) =>
    block.layout.kind === 'paragraph' ? [block.layout] : []) ?? [];
}

function bodyParagraphs(paragraphs: DocParagraph[]): ParagraphLayout[] {
  const document = model();
  document.body = paragraphs as unknown as BodyElement[];
  const layout = layoutDocument(
    document,
    createLayoutServices(document, { measureContext: makeCtx() }),
    { currentDateMs: 0 },
  );
  return layout.pages.flatMap((page) => page.layers.body.flatMap((node) =>
    node.kind === 'paragraph' ? [node] : []));
}

describe('table-cell parser-owned anchor reflow', () => {
  it('translates a public VML text-relative rectangle with its owning cell', () => {
    const retainedParagraph = cellParagraphs(model([
      paragraph([textRun('Postcode'), vmlRectRun(10)]),
    ]))[0]!;
    const drawing = retainedParagraph.drawings.find((candidate) =>
      candidate.anchorLayer?.acquisitionOccurrenceId?.startsWith('public-shape:'));

    expect(drawing).toBeDefined();
    expect(drawing!.flowBounds.xPt).toBe(10);
    // A VML `mso-position-horizontal-relative:text` rectangle maps to the
    // containing column. Inside a table that column is the cell-local text
    // frame, so the final page projection must translate it with the host.
    expect(drawing!.anchorLayer?.horizontalOwnership).toBe('host');
  });

  it('keeps a layoutInCell column anchor in the owning cell coordinate space', () => {
    const retainedParagraph = cellParagraphs(model([
      paragraph([
        ...anchoredImageRuns({
          occurrenceId: 'cell-column-owned',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: -3,
          widthPt: 80,
          heightPt: 40,
          wrapKind: 'none',
        }),
      ]),
    ]))[0]!;
    const drawing = retainedParagraph.drawings.find((candidate) =>
      candidate.anchorLayer?.occurrenceId.endsWith('cell-column-owned'));

    expect(drawing).toBeDefined();
    // ECMA-376 Part 1 §20.4.2.3: with layoutInCell=true, positioning is
    // relative to the existing cell. The page paint plan must therefore retain
    // the cell placement instead of undoing it as a page-owned coordinate.
    expect(drawing!.flowBounds.xPt).toBe(-3);
    expect(drawing!.anchorLayer?.horizontalOwnership).toBe('host');
  });

  it('resolves a layoutInCell margin anchor from the owning cell', () => {
    const retainedParagraph = cellParagraphs(model([
      paragraph([
        ...anchoredImageRuns({
          occurrenceId: 'cell-margin-owned',
          horizontalRelativeFrom: 'margin',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: 2,
          widthPt: 80,
          heightPt: 40,
          wrapKind: 'none',
        }),
      ]),
    ]))[0]!;
    const drawing = retainedParagraph.drawings.find((candidate) =>
      candidate.anchorLayer?.occurrenceId.endsWith('cell-margin-owned'));

    expect(drawing).toBeDefined();
    expect(drawing!.flowBounds.xPt).toBe(2);
    expect(drawing!.anchorLayer?.horizontalOwnership).toBe('host');
  });

  it.each(['page', 'margin'] as const)(
    'resolves a layoutInCell %s vertical offset from the owning cell',
    (verticalRelativeFrom) => {
      const retainedParagraph = cellParagraphs(model([
        paragraph([
          ...anchoredImageRuns({
            occurrenceId: `cell-${verticalRelativeFrom}-vertical-owned`,
            verticalRelativeFrom,
            horizontalOffsetPt: 0,
            verticalOffsetPt: 6,
            widthPt: 80,
            heightPt: 40,
            wrapKind: 'none',
          }),
        ]),
      ]))[0]!;
      const drawing = retainedParagraph.drawings.find((candidate) =>
        candidate.anchorLayer?.occurrenceId.endsWith(
          `cell-${verticalRelativeFrom}-vertical-owned`,
        ));

      expect(drawing).toBeDefined();
      expect(drawing!.flowBounds.yPt).toBe(6);
      expect(drawing!.anchorLayer?.verticalOwnership).toBe('host');
    },
  );

  it('resolves layoutInCell simplePos coordinates from the owning cell', () => {
    const retainedParagraph = cellParagraphs(model([
      paragraph([
        ...anchoredImageRuns({
          occurrenceId: 'cell-simple-position-owned',
          simplePosition: { xPt: 3, yPt: 7 },
          horizontalOffsetPt: 0,
          widthPt: 80,
          heightPt: 40,
          wrapKind: 'none',
        }),
      ]),
    ]))[0]!;
    const drawing = retainedParagraph.drawings.find((candidate) =>
      candidate.anchorLayer?.occurrenceId.endsWith('cell-simple-position-owned'));

    expect(drawing).toBeDefined();
    expect(drawing!.flowBounds).toMatchObject({ xPt: 3, yPt: 7 });
    expect(drawing!.anchorLayer).toMatchObject({
      horizontalOwnership: 'host',
      verticalOwnership: 'host',
    });
  });

  it('retains both cell-owned axes through the final page paint translation', () => {
    const document = model([
      paragraph([
        ...anchoredImageRuns({
          occurrenceId: 'cell-final-paint-owned',
          verticalRelativeFrom: 'page',
          horizontalOffsetPt: 2,
          verticalOffsetPt: 5,
          widthPt: 80,
          heightPt: 40,
          wrapKind: 'none',
        }),
      ]),
    ]);
    const result = layoutDocument(
      document,
      createLayoutServices(document, { measureContext: makeCtx() }),
      { currentDateMs: 0 },
    );
    const table = result.pages[0]!.layers.body.find((node) => node.kind === 'table');
    if (!table || table.kind !== 'table') throw new Error('Expected retained table');
    const cell = table.rows[0]!.cells[0]!;
    const block = cell.blocks[0]!;
    const drawingEntry = result.pages[0]!.layers.paintOrder.find((entry) =>
      entry.kind === 'drawing'
      && entry.node.anchorLayer?.occurrenceId.includes('cell-final-paint-owned'));

    expect(drawingEntry?.kind).toBe('drawing');
    if (!drawingEntry || drawingEntry.kind !== 'drawing') {
      throw new Error('Expected retained page drawing entry');
    }
    expect(drawingEntry.node.anchorLayer).toMatchObject({
      horizontalOwnership: 'host',
      verticalOwnership: 'host',
    });
    expect(drawingEntry.layoutTranslationPt).toEqual({
      xPt: cell.contentBounds.xPt,
      yPt: cell.flowBounds.yPt + block.offsetPt,
    });
    expect({
      xPt: drawingEntry.node.flowBounds.xPt + drawingEntry.layoutTranslationPt.xPt,
      yPt: drawingEntry.node.flowBounds.yPt + drawingEntry.layoutTranslationPt.yPt,
    }).toEqual({
      xPt: cell.contentBounds.xPt + 2,
      yPt: cell.flowBounds.yPt + block.offsetPt + 5,
    });
  });

  it('measures and retains the same occurrence-keyed exclusion fixed point', () => {
    const paragraph = cellParagraphs()[0]!;
    const lineTexts = paragraph.lines.map((line) => line.placements
      .filter((placement) => placement.kind === 'text')
      .map((placement) => placement.text)
      .join(''));
    const exclusion = paragraph.exclusions.find((candidate) =>
      candidate.anchorOccurrenceId?.endsWith('cell-later'));

    expect(lineTexts).toEqual(['AAAAAAAA', 'A', 'BBBB', 'BBBB']);
    expect(exclusion).toBeDefined();
    expect(exclusion!.bounds).toMatchObject({
      xPt: 80,
      yPt: paragraph.flowBounds.yPt + 20,
      widthPt: 80,
      heightPt: 40,
    });
  });

  it('carries a retained host exclusion into the following cell paragraph', () => {
    const [anchored, following] = cellParagraphs(model([
      paragraph([
        ...anchoredImageRuns({
          occurrenceId: 'cell-following',
          verticalRelativeFrom: 'paragraph',
          heightPt: 60,
        }),
        textRun('A'),
      ]),
      paragraph([textRun('B'.repeat(8))]),
    ]));
    const followingLines = following!.lines.map((line) => line.placements
      .filter((placement) => placement.kind === 'text')
      .map((placement) => placement.text)
      .join(''));

    expect(anchored!.exclusions).toHaveLength(1);
    expect(following!.flowBounds.yPt).toBe(20);
    expect(followingLines).toEqual(['BBBB', 'BBBB']);
    expect(following!.exclusions.map((candidate) =>
      candidate.anchorOccurrenceId)).toEqual([
      anchored!.exclusions[0]!.anchorOccurrenceId,
    ]);
  });

  it('preserves inherited anchor authority during page-dependent cell reacquisition', () => {
    const [anchored, following] = cellParagraphs(model([
      paragraph([
        ...anchoredImageRuns({
          occurrenceId: 'cell-page-dependent-prior',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: 0,
          widthPt: 80,
          heightPt: 60,
          wrapKind: 'none',
        }),
        textRun('A'),
      ]),
      paragraph([
        pageFieldRun(),
        ...anchoredImageRuns({
          occurrenceId: 'cell-page-dependent-moving',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: 0,
          widthPt: 80,
          heightPt: 40,
          wrapKind: 'none',
          allowOverlap: false,
        }),
        textRun('B'.repeat(8)),
      ]),
    ]));
    const blocker = anchored!.drawings.find((drawing) =>
      drawing.anchorLayer?.occurrenceId.endsWith('cell-page-dependent-prior'));
    const moving = following!.drawings.find((drawing) =>
      drawing.anchorLayer?.occurrenceId.endsWith('cell-page-dependent-moving'));

    expect(blocker).toBeDefined();
    expect(moving).toBeDefined();
    expect(moving!.flowBounds.xPt).toBe(
      blocker!.flowBounds.xPt + blocker!.flowBounds.widthPt,
    );
    expect(following!.anchorCollisions?.map((entry) => entry.occurrenceId))
      .toEqual(expect.arrayContaining([
        anchored!.anchorCollisions![0]!.occurrenceId,
        expect.stringContaining('cell-page-dependent-moving'),
      ]));
  });

  it('repositions a later parser-owned anchor when allowOverlap is false', () => {
    const [first, second] = cellParagraphs(model([
      paragraph([
        ...anchoredImageRuns({
          occurrenceId: 'cell-blocker',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: 0,
          widthPt: 80,
          heightPt: 60,
        }),
        textRun('A'),
      ]),
      paragraph([
        ...anchoredImageRuns({
          occurrenceId: 'cell-moving',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: 0,
          widthPt: 80,
          heightPt: 40,
          allowOverlap: false,
        }),
        textRun('B'),
      ]),
    ]));
    const firstExclusion = first!.exclusions.find((candidate) =>
      candidate.anchorOccurrenceId?.endsWith('cell-blocker'));
    const secondExclusion = second!.exclusions.find((candidate) =>
      candidate.anchorOccurrenceId?.endsWith('cell-moving'));

    expect(firstExclusion).toBeDefined();
    expect(secondExclusion).toBeDefined();
    expect(secondExclusion!.bounds.xPt).toBe(
      firstExclusion!.bounds.xPt + firstExclusion!.bounds.widthPt,
    );
    expect(secondExclusion!.bounds.yPt).toBe(second!.flowBounds.yPt);
  });

  it('preserves prior-paragraph collision avoidance when allowOverlap permits overlap', () => {
    const [first, second] = cellParagraphs(model([
      paragraph([
        ...anchoredImageRuns({
          occurrenceId: 'cell-observed-blocker',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: 0,
          widthPt: 80,
          heightPt: 60,
        }),
        textRun('A'),
      ]),
      paragraph([
        ...anchoredImageRuns({
          occurrenceId: 'cell-observed-moving',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: 0,
          widthPt: 80,
          heightPt: 40,
          allowOverlap: true,
        }),
        textRun('B'),
      ]),
    ]));
    const firstExclusion = first!.exclusions.find((candidate) =>
      candidate.anchorOccurrenceId?.endsWith('cell-observed-blocker'));
    const secondExclusion = second!.exclusions.find((candidate) =>
      candidate.anchorOccurrenceId?.endsWith('cell-observed-moving'));

    expect(firstExclusion).toBeDefined();
    expect(secondExclusion).toBeDefined();
    expect(secondExclusion!.bounds.xPt).toBe(
      firstExclusion!.bounds.xPt + firstExclusion!.bounds.widthPt,
    );
  });

  it('applies allowOverlap=false between anchors hosted by the same paragraph', () => {
    const [layout] = cellParagraphs(model([
      paragraph([
        ...anchoredImageRuns({
          occurrenceId: 'cell-same-blocker',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: 0,
          widthPt: 80,
          heightPt: 40,
        }),
        ...anchoredImageRuns({
          occurrenceId: 'cell-same-moving',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: 0,
          widthPt: 80,
          heightPt: 40,
          allowOverlap: false,
        }),
        textRun('A'),
      ]),
    ]));
    const blocker = layout!.exclusions.find((candidate) =>
      candidate.anchorOccurrenceId?.endsWith('cell-same-blocker'));
    const moving = layout!.exclusions.find((candidate) =>
      candidate.anchorOccurrenceId?.endsWith('cell-same-moving'));

    expect(blocker).toBeDefined();
    expect(moving).toBeDefined();
    expect(moving!.bounds.xPt).toBe(blocker!.bounds.xPt + blocker!.bounds.widthPt);
  });

  it('allows anchors hosted by the same paragraph to overlap when permitted', () => {
    const [layout] = cellParagraphs(model([
      paragraph([
        ...anchoredImageRuns({
          occurrenceId: 'cell-same-permitted-blocker',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: 0,
          widthPt: 80,
          heightPt: 40,
        }),
        ...anchoredImageRuns({
          occurrenceId: 'cell-same-permitted-moving',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: 0,
          widthPt: 80,
          heightPt: 40,
          allowOverlap: true,
        }),
        textRun('A'),
      ]),
    ]));
    const blocker = layout!.exclusions.find((candidate) =>
      candidate.anchorOccurrenceId?.endsWith('cell-same-permitted-blocker'));
    const moving = layout!.exclusions.find((candidate) =>
      candidate.anchorOccurrenceId?.endsWith('cell-same-permitted-moving'));

    expect(blocker).toBeDefined();
    expect(moving).toBeDefined();
    expect(moving!.bounds).toEqual(blocker!.bounds);
  });

  it('uses a wrapNone DrawingML object as an allowOverlap=false blocker', () => {
    const [layout] = cellParagraphs(model([
      paragraph([
        ...anchoredImageRuns({
          occurrenceId: 'cell-wrap-none-blocker',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: 0,
          widthPt: 80,
          heightPt: 40,
          wrapKind: 'none',
        }),
        ...anchoredImageRuns({
          occurrenceId: 'cell-wrapped-moving',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: 0,
          widthPt: 80,
          heightPt: 40,
          allowOverlap: false,
        }),
        textRun('A'),
      ]),
    ]));
    const blocker = layout!.drawings.find((candidate) =>
      candidate.anchorLayer?.occurrenceId.endsWith('cell-wrap-none-blocker'));
    const moving = layout!.drawings.find((candidate) =>
      candidate.anchorLayer?.occurrenceId.endsWith('cell-wrapped-moving'));

    expect(blocker).toBeDefined();
    expect(moving).toBeDefined();
    expect(moving!.flowBounds.xPt).toBe(
      blocker!.flowBounds.xPt + blocker!.flowBounds.widthPt,
    );
  });

  it('repositions a wrapNone allowOverlap=false object around a wrapped blocker', () => {
    const [layout] = cellParagraphs(model([
      paragraph([
        ...anchoredImageRuns({
          occurrenceId: 'cell-wrapped-blocker',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: 0,
          widthPt: 80,
          heightPt: 40,
        }),
        ...anchoredImageRuns({
          occurrenceId: 'cell-wrap-none-moving',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: 0,
          widthPt: 80,
          heightPt: 40,
          wrapKind: 'none',
          allowOverlap: false,
        }),
        textRun('A'),
      ]),
    ]));
    const blocker = layout!.drawings.find((candidate) =>
      candidate.anchorLayer?.occurrenceId.endsWith('cell-wrapped-blocker'));
    const moving = layout!.drawings.find((candidate) =>
      candidate.anchorLayer?.occurrenceId.endsWith('cell-wrap-none-moving'));

    expect(blocker).toBeDefined();
    expect(moving).toBeDefined();
    expect(moving!.flowBounds.xPt).toBe(
      blocker!.flowBounds.xPt + blocker!.flowBounds.widthPt,
    );
  });

  it('carries a wrapNone collision blocker into the following cell paragraph', () => {
    const [first, second] = cellParagraphs(model([
      paragraph([
        ...anchoredImageRuns({
          occurrenceId: 'cell-prior-wrap-none',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: 0,
          widthPt: 80,
          heightPt: 60,
          wrapKind: 'none',
        }),
        textRun('A'),
      ]),
      paragraph([
        ...anchoredImageRuns({
          occurrenceId: 'cell-later-wrapped',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: 0,
          widthPt: 80,
          heightPt: 40,
          allowOverlap: false,
        }),
        textRun('B'),
      ]),
    ]));
    const blocker = first!.drawings.find((candidate) =>
      candidate.anchorLayer?.occurrenceId.endsWith('cell-prior-wrap-none'));
    const moving = second!.drawings.find((candidate) =>
      candidate.anchorLayer?.occurrenceId.endsWith('cell-later-wrapped'));

    expect(blocker).toBeDefined();
    expect(moving).toBeDefined();
    expect(moving!.flowBounds.xPt).toBe(
      blocker!.flowBounds.xPt + blocker!.flowBounds.widthPt,
    );
  });

  it('keeps layoutInCell collision displacement inside the owning cell', () => {
    const [layout] = cellParagraphs(model([
      paragraph([
        ...anchoredImageRuns({
          occurrenceId: 'cell-narrow-blocker',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: 0,
          widthPt: 80,
          heightPt: 40,
        }),
        ...anchoredImageRuns({
          occurrenceId: 'cell-narrow-moving',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: 0,
          widthPt: 80,
          heightPt: 40,
          allowOverlap: false,
        }),
        textRun('A'),
      ]),
    ], 100));
    const blocker = layout!.drawings.find((candidate) =>
      candidate.anchorLayer?.occurrenceId.endsWith('cell-narrow-blocker'));
    const moving = layout!.drawings.find((candidate) =>
      candidate.anchorLayer?.occurrenceId.endsWith('cell-narrow-moving'));

    expect(blocker).toBeDefined();
    expect(moving).toBeDefined();
    expect(moving!.flowBounds.xPt).toBe(0);
    expect(moving!.flowBounds.yPt).toBe(
      blocker!.flowBounds.yPt + blocker!.flowBounds.heightPt,
    );
    expect(moving!.flowBounds.xPt + moving!.flowBounds.widthPt).toBeLessThanOrEqual(100);
  });

  it('grows an automatic row to contain a displaced layoutInCell wrapNone object', () => {
    const document = model([
      paragraph([
        ...anchoredImageRuns({
          occurrenceId: 'cell-containment-blocker',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: 0,
          widthPt: 80,
          heightPt: 40,
          wrapKind: 'none',
        }),
        ...anchoredImageRuns({
          occurrenceId: 'cell-containment-moving',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: 0,
          widthPt: 80,
          heightPt: 40,
          allowOverlap: false,
          wrapKind: 'none',
        }),
        textRun('A'),
      ]),
    ], 100);
    const result = layoutDocument(
      document,
      createLayoutServices(document, { measureContext: makeCtx() }),
      { currentDateMs: 0 },
    );
    const table = result.pages[0]!.layers.body.find((node) => node.kind === 'table');
    if (!table || table.kind !== 'table') throw new Error('Expected retained table');
    const row = table.rows[0]!;
    const cell = row.cells[0]!;
    const paragraphBlock = cell.blocks[0]!;
    if (paragraphBlock.layout.kind !== 'paragraph') {
      throw new Error('Expected retained cell paragraph');
    }
    const paragraphLayout = paragraphBlock.layout;
    const moving = paragraphLayout.drawings.find((candidate) =>
      candidate.anchorLayer?.occurrenceId.endsWith('cell-containment-moving'));
    expect(moving).toBeDefined();
    const placedMovingBottomPt = cell.flowBounds.yPt
      + paragraphBlock.offsetPt
      + moving!.flowBounds.yPt
      - paragraphLayout.flowBounds.yPt
      + moving!.flowBounds.heightPt;

    expect(cell.flowBounds.yPt + cell.flowBounds.heightPt).toBeGreaterThanOrEqual(
      placedMovingBottomPt,
    );
    expect(row.flowBounds.yPt + row.flowBounds.heightPt).toBeGreaterThanOrEqual(
      placedMovingBottomPt,
    );
  });

  it('does not grow an automatic row for an overlapping layoutInCell wrapNone object', () => {
    const baseline = model([
      paragraph([textRun('A')]),
    ], 100);
    const withOverlay = model([
      paragraph([
        ...anchoredImageRuns({
          occurrenceId: 'cell-overlay',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: 0,
          verticalOffsetPt: 10,
          widthPt: 80,
          heightPt: 80,
          allowOverlap: true,
          wrapKind: 'none',
        }),
        textRun('A'),
      ]),
    ], 100);
    const rowHeight = (document: DocxDocumentModel): number => {
      const result = layoutDocument(
        document,
        createLayoutServices(document, { measureContext: makeCtx() }),
        { currentDateMs: 0 },
      );
      const table = result.pages[0]!.layers.body.find((node) => node.kind === 'table');
      if (!table || table.kind !== 'table') throw new Error('Expected retained table');
      return table.rows[0]!.heightPt;
    };

    expect(rowHeight(withOverlay)).toBe(rowHeight(baseline));
  });

  it('moves a cell-owned overlay row when its image would leave the remaining page band', () => {
    const document = model([
      paragraph(anchoredImageRuns({
        occurrenceId: 'page-edge-cell-image',
        verticalRelativeFrom: 'paragraph',
        verticalOffsetPt: 10,
        widthPt: 20,
        heightPt: 70,
        allowOverlap: true,
        wrapKind: 'none',
      })),
      paragraph([]), paragraph([]), paragraph([]), paragraph([]),
    ]);
    const table = document.body[0] as DocTable;
    table.rows[0]!.rowHeight = 49;
    table.rows[0]!.rowHeightRule = 'atLeast';
    const leading = paragraph([textRun('leading')]);
    leading.spaceAfter = 80;
    document.body = [leading, table] as unknown as BodyElement[];

    const result = layoutDocument(
      document,
      createLayoutServices(document, { measureContext: makeCtx() }),
      { currentDateMs: 0 },
    );
    expect(result.pages[0]?.layers.body.some((node) => node.kind === 'table')).toBe(false);
    expect(result.pages[1]?.layers.body.find((node) => node.kind === 'table'))
      .toMatchObject({ rows: [{ fragmentIndex: 0 }] });
  });
});

describe('body parser-owned anchor collision carry', () => {
  it('retains wrapNone collision authority without creating a text exclusion', () => {
    const baseline = bodyParagraphs([
      paragraph([textRun('following text')]),
    ])[0]!;
    const [, following] = bodyParagraphs([
      paragraph([
        ...anchoredImageRuns({
          occurrenceId: 'body-wrap-none-text-neutral',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: 0,
          widthPt: 80,
          heightPt: 60,
          wrapKind: 'none',
        }),
        textRun('A'),
      ]),
      paragraph([textRun('following text')]),
    ]);
    const relativeLines = (layout: ParagraphLayout) => layout.lines.map((line) => ({
      xPt: line.bounds.xPt,
      yPt: line.bounds.yPt - layout.flowBounds.yPt,
      widthPt: line.bounds.widthPt,
      advancePt: line.advancePt,
      text: line.placements.flatMap((placement) =>
        placement.kind === 'text' ? [placement.text] : []).join(''),
    }));

    expect(following?.anchorCollisions).toEqual([
      expect.objectContaining({
        occurrenceId: expect.stringContaining('body-wrap-none-text-neutral'),
      }),
    ]);
    expect(following?.exclusions).toEqual([]);
    expect(relativeLines(following!)).toEqual(relativeLines(baseline));
  });

  it('carries a wrapNone blocker into the following body paragraph', () => {
    const [first, second] = bodyParagraphs([
      paragraph([
        ...anchoredImageRuns({
          occurrenceId: 'body-prior-wrap-none',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: 0,
          widthPt: 80,
          heightPt: 60,
          wrapKind: 'none',
        }),
        textRun('A'),
      ]),
      paragraph([
        ...anchoredImageRuns({
          occurrenceId: 'body-later-wrapped',
          verticalRelativeFrom: 'paragraph',
          horizontalOffsetPt: 0,
          widthPt: 80,
          heightPt: 40,
          allowOverlap: false,
        }),
        textRun('B'),
      ]),
    ]);
    const blocker = first!.drawings.find((candidate) =>
      candidate.anchorLayer?.occurrenceId.endsWith('body-prior-wrap-none'));
    const moving = second!.drawings.find((candidate) =>
      candidate.anchorLayer?.occurrenceId.endsWith('body-later-wrapped'));

    expect(blocker).toBeDefined();
    expect(moving).toBeDefined();
    expect(moving!.flowBounds.xPt).toBe(
      blocker!.flowBounds.xPt + blocker!.flowBounds.widthPt,
    );
  });
});

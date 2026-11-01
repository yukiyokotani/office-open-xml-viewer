import { describe, expect, it, vi } from 'vitest';
import type { ParagraphLayoutContext } from '../layout-context.js';
import type {
  BodyElement,
  DocParagraph,
  DocRun,
  DocxDocumentModel,
  ShapeRun,
} from '../types.js';
import { createLayoutServices } from '../layout-runtime.js';
import { layoutDocument } from '../document-layout.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import { acquireShapeTextBoxLayout } from './paragraph.js';
import type {
  FloatingTablePlacementLayout,
  ParagraphLayout,
  ResolvedFloatingTablePlacementLayout,
  StoryLayout,
  TableLayout,
} from './types.js';

const paragraphContext: ParagraphLayoutContext = {
  lineGrid: { active: false, pitchPt: null },
  characterGrid: { active: false, kind: null, pitchPt: null, deltaPt: 0 },
  rightIndentGrid: { pitchPt: null, paragraphAllowsAdjustment: true },
  physicalIndentLeftPt: 0,
  physicalIndentRightPt: 0,
  firstIndentPt: 0,
  lineSpacing: null,
  spaceBeforePt: 0,
  spaceAfterPt: 0,
  baseRtl: false,
  isJustified: false,
  stretchLastLine: false,
  tabStops: [],
  hasRuby: false,
  hasEastAsianText: false,
  kinsoku: { enabled: true, lineStartForbidden: new Set(), lineEndForbidden: new Set() },
  defaultTabPt: 36,
};

function canvas(): CanvasRenderingContext2D {
  return {
    font: '10px sans-serif',
    measureText: () => ({
      width: 10,
      fontBoundingBoxAscent: 8,
      fontBoundingBoxDescent: 2,
      actualBoundingBoxAscent: 8,
      actualBoundingBoxDescent: 2,
    }),
  } as unknown as CanvasRenderingContext2D;
}

function paragraph(): ParagraphLayout {
  const bounds = { xPt: 7, yPt: 11, widthPt: 80, heightPt: 10 };
  return {
    kind: 'paragraph',
    id: 'textbox-paragraph',
    source: { story: 'textbox', storyInstance: 'shape:9', path: [0] },
    flowDomainId: 'textbox:shape:9',
    ordinaryFlow: true,
    flowBounds: bounds,
    inkBounds: bounds,
    advancePt: 10,
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

function table(): TableLayout {
  const bounds = { xPt: 7, yPt: 21, widthPt: 80, heightPt: 15 };
  return {
    kind: 'table',
    id: 'textbox-table',
    source: { story: 'textbox', storyInstance: 'shape:9', path: [1] },
    flowDomainId: 'textbox:shape:9',
    ordinaryFlow: true,
    flowBounds: bounds,
    inkBounds: bounds,
    advancePt: 15,
    columnWidthsPt: [80],
    rows: [],
    borders: [],
  };
}

function paragraphAt(
  yPt: number,
  advancePt: number,
  options: Readonly<{
    visible?: boolean;
    afterPt?: number;
    path?: number;
  }> = {},
): ParagraphLayout {
  const afterPt = options.afterPt ?? 0;
  const visible = options.visible ?? false;
  const base = paragraph();
  const bounds = { xPt: 7, yPt, widthPt: 80, heightPt: advancePt };
  return {
    ...base,
    id: `textbox-paragraph:${options.path ?? 0}`,
    source: {
      story: 'textbox',
      storyInstance: 'shape:anchor',
      path: [options.path ?? 0],
    },
    flowBounds: bounds,
    inkBounds: bounds,
    advancePt,
    spacing: { beforePt: 0, afterPt },
    lines: visible ? [{
      range: { start: 0, end: 1 },
      bounds: { xPt: 7, yPt, widthPt: 10, heightPt: advancePt - afterPt },
      baselinePt: yPt + Math.max(0, advancePt - afterPt),
      advancePt: Math.max(0, advancePt - afterPt),
      placements: [{
        kind: 'text',
        text: 'x',
        range: { start: 0, end: 1 },
        origin: { xPt: 7, yPt: yPt + Math.max(0, advancePt - afterPt) },
        bounds: { xPt: 7, yPt, widthPt: 10, heightPt: Math.max(0, advancePt - afterPt) },
        advancePt: 10,
        clusters: [],
        paintOps: [],
        color: { kind: 'default' },
        fontRoute: {},
        fontSizePt: 10,
        fontWeight: 400,
        fontStyle: 'normal',
        direction: 'ltr',
        decorations: [],
      } as unknown as ParagraphLayout['lines'][number]['placements'][number]],
    }] : [],
  };
}

function anchoredTextBox(
  retainedStory: StoryLayout,
  anchor: 'ctr' | 'b' = 'ctr',
  options: Readonly<{
    textVert?: 'eaVert' | 'mongolianVert';
    rect?: Readonly<{ xPt: number; yPt: number; widthPt: number; heightPt: number }>;
  }> = {},
): ReturnType<typeof acquireShapeTextBoxLayout> {
  return acquireShapeTextBoxLayout(
    {
      type: 'shape',
      textAnchor: anchor,
      ...(options.textVert ? { textVert: options.textVert } : {}),
      textInsetL: 0,
      textInsetT: 0,
      textInsetR: 0,
      textInsetB: 0,
      textAutofit: 'none',
    } as unknown as ShapeRun,
    options.rect ?? { xPt: 0, yPt: 0, widthPt: 100, heightPt: 100 },
    {
      id: 'shape:anchor',
      source: { story: 'textbox', storyInstance: 'shape:anchor', path: [] },
      flowDomainId: 'shape-owner',
      context: paragraphContext,
      measurer: { context: canvas(), fontFamilyClasses: {} },
      environment: {
        pageIndex: 0,
        totalPages: 1,
        pageWritingMode: 'horizontal-tb',
        documentHasEastAsianText: false,
      },
      input: {
        kind: 'complete',
        source: { story: 'textbox', storyInstance: 'shape:anchor', path: [] },
        blockCount: retainedStory.blocks.length,
      },
      acquireCompleteStory: () => retainedStory,
    },
  );
}

describe('rich text-box story acquisition', () => {
  it('centers the smallest visible text bounds without deleting a trailing empty paragraph', () => {
    const title = paragraphAt(0, 50, { visible: true, path: 0 });
    const trailingEmpty = paragraphAt(50, 18, { afterPt: 6, path: 1 });
    const retainedStory: StoryLayout = {
      story: 'textbox',
      flowBounds: { xPt: 7, yPt: 0, widthPt: 80, heightPt: 68 },
      inkBounds: { xPt: 7, yPt: 0, widthPt: 80, heightPt: 68 },
      blocks: [title, trailingEmpty],
      advancePt: 68,
      diagnostics: [],
    };

    const layout = anchoredTextBox(retainedStory);
    const bottomLayout = anchoredTextBox(retainedStory, 'b');

    expect(layout?.story.advancePt).toBe(68);
    expect(layout?.story.blocks).toHaveLength(2);
    expect(layout?.story.blocks[0]?.flowBounds.yPt).toBe(25);
    expect(layout?.story.blocks[1]?.flowBounds.yPt).toBe(75);
    expect(layout?.story.blocks[1]).toMatchObject({
      kind: 'paragraph',
      spacing: { afterPt: 6 },
      lines: [],
    });
    expect(bottomLayout?.story.blocks.map((block) => block.flowBounds.yPt))
      .toEqual([50, 100]);
  });

  it('derives the anchor extent before projecting a complete story into vertical paint geometry', () => {
    // A 40pt physical width becomes the vertical text body's 40pt logical
    // block-axis capacity, with its local origin at -20pt.
    const title = paragraphAt(-20, 20, { visible: true, path: 0 });
    const trailingEmpty = paragraphAt(0, 10, { path: 1 });
    const retainedStory: StoryLayout = {
      story: 'textbox',
      flowBounds: { xPt: -50, yPt: -20, widthPt: 100, heightPt: 30 },
      inkBounds: { xPt: -50, yPt: -20, widthPt: 100, heightPt: 30 },
      blocks: [title, trailingEmpty],
      advancePt: 30,
      diagnostics: [],
    };

    const layout = anchoredTextBox(retainedStory, 'ctr', {
      textVert: 'mongolianVert',
      rect: { xPt: 0, yPt: 0, widthPt: 40, heightPt: 100 },
    });

    expect(layout).toMatchObject({
      verticalMode: 'mongolianVert',
      story: {
        advancePt: 30,
        blocks: [
          { kind: 'paragraph', flowBounds: { yPt: -10 } },
          { kind: 'paragraph', flowBounds: { yPt: 10 }, lines: [] },
        ],
      },
    });
  });

  it('retains meaningful inter-paragraph spacing and a terminal table in the anchor extent', () => {
    const first = paragraphAt(0, 20, { visible: true, afterPt: 10, path: 0 });
    const second = paragraphAt(30, 20, { visible: true, afterPt: 5, path: 1 });
    const terminalTable = {
      ...table(),
      source: { story: 'textbox', storyInstance: 'shape:anchor', path: [2] },
      flowBounds: { xPt: 7, yPt: 50, widthPt: 80, heightPt: 15 },
      inkBounds: { xPt: 7, yPt: 50, widthPt: 80, heightPt: 15 },
      advancePt: 15,
    } as TableLayout;
    const structuralEmpty = paragraphAt(65, 12, { path: 3 });
    const retainedStory: StoryLayout = {
      story: 'textbox',
      flowBounds: { xPt: 7, yPt: 0, widthPt: 80, heightPt: 77 },
      inkBounds: { xPt: 7, yPt: 0, widthPt: 80, heightPt: 77 },
      blocks: [first, second, terminalTable, structuralEmpty],
      advancePt: 77,
      diagnostics: [],
    };

    const layout = anchoredTextBox(retainedStory);

    // The visible extent is 65pt: the 10pt inter-paragraph gap and terminal
    // table remain authoritative, while only the structural empty tail is
    // outside the smallest visible text bounds.
    expect(layout?.story.blocks.map((block) => block.flowBounds.yPt))
      .toEqual([17.5, 47.5, 67.5, 82.5]);
    expect(layout?.story.blocks[0]).toMatchObject({
      kind: 'paragraph',
      spacing: { afterPt: 10 },
    });
    expect(layout?.story.blocks[1]).toMatchObject({
      kind: 'paragraph',
      spacing: { afterPt: 5 },
    });
    expect(layout?.story.blocks[2]?.kind).toBe('table');
    expect(layout?.story.advancePt).toBe(77);
  });

  it('prefers complete parser blocks and retains one nested StoryLayout', () => {
    const richBlocks = [
      { type: 'paragraph', runs: [{ type: 'text', text: 'before' }] },
      {
        type: 'table',
        rows: [{ cells: [{ content: [
          { type: 'paragraph', runs: [{ type: 'text', text: 'cell' }] },
        ] }] }],
      },
      { type: 'unsupportedTextBoxBlock', qName: 'w:altChunk', sourcePath: [2] },
    ];
    const retainedStory: StoryLayout = {
      story: 'textbox',
      flowBounds: { xPt: 7, yPt: 11, widthPt: 80, heightPt: 25 },
      inkBounds: { xPt: 7, yPt: 11, widthPt: 80, heightPt: 25 },
      clipBounds: { xPt: 7, yPt: 11, widthPt: 80, heightPt: 44 },
      blocks: [paragraph(), table()],
      advancePt: 25,
      diagnostics: [{
        code: 'UNSUPPORTED_FEATURE',
        severity: 'warning',
        source: { story: 'textbox', storyInstance: 'shape:9', path: [2] },
        message: 'Unsupported text-box block w:altChunk',
      }],
    };
    const acquireCompleteStory = vi.fn(() => retainedStory);
    const shape = {
      type: 'shape',
      textInsetL: 2,
      textInsetT: 3,
      textInsetR: 4,
      textInsetB: 5,
      textAutofit: 'none',
    } as unknown as ShapeRun;

    const layout = acquireShapeTextBoxLayout(
      shape,
      { xPt: 5, yPt: 8, widthPt: 86, heightPt: 52 },
      {
        id: 'shape:9',
        source: { story: 'textbox', storyInstance: 'shape:9', path: [] },
        flowDomainId: 'shape-owner',
        context: paragraphContext,
        measurer: { context: canvas(), fontFamilyClasses: {} },
        environment: {
          pageIndex: 0,
          totalPages: 1,
          pageWritingMode: 'horizontal-tb',
          documentHasEastAsianText: false,
        },
        input: {
          kind: 'complete',
          source: { story: 'textbox', storyInstance: 'shape:9', path: [] },
          blockCount: richBlocks.length,
        },
        acquireCompleteStory,
      },
    );

    expect(acquireCompleteStory).toHaveBeenCalledOnce();
    expect(acquireCompleteStory).toHaveBeenCalledWith(expect.objectContaining({
      source: { story: 'textbox', storyInstance: 'shape:9', path: [] },
      container: {
        id: 'shape:9:story',
        kind: 'textbox',
        bounds: { xPt: 7, yPt: 11, widthPt: 80, heightPt: 44 },
        capacity: 'unbounded',
      },
    }));
    expect(layout).toMatchObject({
      kind: 'textbox',
      story: retainedStory,
      flowBounds: { xPt: 5, yPt: 8, widthPt: 86, heightPt: 52 },
      clipBounds: { xPt: 7, yPt: 11, widthPt: 80, heightPt: 44 },
      transform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
    });
    expect(layout).not.toHaveProperty('paragraphs');
  });

  it.each([
    // WORD_TEXT_BOX_PHYSICAL_INSETS: lIns/tIns/rIns/bIns stay on the physical
    // edges of a quarter-turned story. Lines (logical x) run from tIns to bIns
    // for vert/eaVert and from bIns up to tIns for vert270; the first column
    // sits rIns inside the right edge (vert/eaVert) or lIns inside the left.
    ['vert', { xPt: -23, yPt: -39, widthPt: 44, heightPt: 80 }, [87, 11]],
    ['eaVert', { xPt: -23, yPt: -39, widthPt: 44, heightPt: 80 }, [87, 11]],
    ['vert270', { xPt: -21, yPt: -41, widthPt: 44, heightPt: 80 }, [7, 55]],
  ] as const)('keeps %s text-box insets on their physical edges', (textVert, bounds, [startX, startY]) => {
    const story: StoryLayout = {
      story: 'textbox', flowBounds: bounds, inkBounds: bounds, clipBounds: bounds,
      blocks: [], advancePt: 0, diagnostics: [],
    };
    const acquireCompleteStory = vi.fn(() => story);
    const shape = {
      type: 'shape', textInsetL: 2, textInsetT: 3, textInsetR: 4, textInsetB: 5,
      textAutofit: 'none', textVert,
    } as unknown as ShapeRun;
    const layout = acquireShapeTextBoxLayout(
      shape,
      { xPt: 5, yPt: 8, widthPt: 86, heightPt: 52 },
      {
        id: 'shape:v',
        source: { story: 'textbox', storyInstance: 'shape:v', path: [] },
        flowDomainId: 'shape-owner',
        context: paragraphContext,
        measurer: { context: canvas(), fontFamilyClasses: {} },
        environment: {
          pageIndex: 0, totalPages: 1, pageWritingMode: 'horizontal-tb', documentHasEastAsianText: false,
        },
        input: { kind: 'complete', source: { story: 'textbox', storyInstance: 'shape:v', path: [] }, blockCount: 1 },
        acquireCompleteStory,
      },
    );
    expect(acquireCompleteStory).toHaveBeenCalledWith(expect.objectContaining({
      container: expect.objectContaining({ bounds }),
    }));
    if (!layout) throw new Error('expected a text-box layout');
    const t = layout.transform;
    // The logical line start of the first column, mapped to the page.
    expect([t.a * bounds.xPt + t.c * bounds.yPt + t.e, t.b * bounds.xPt + t.d * bounds.yPt + t.f])
      .toEqual([startX, startY]);
  });

  it('keeps vertical spAutoFit floating-table geometry in the shifted local story frame', () => {
    const rect = (xPt: number, yPt: number, widthPt = 10, heightPt = 10) => ({
      xPt, yPt, widthPt, heightPt,
    });
    const floatingChild: TableLayout = {
      ...table(),
      id: 'textbox-floating-child',
      flowBounds: rect(2, 5),
      inkBounds: rect(2, 5),
    };
    const sourcePlacement = {
      kind: 'floating-table-placement',
      occurrenceId: 'textbox-float',
      ownership: 'source',
      physicalPageIndex: 0,
      displayPageNumber: 1,
      hostCellId: 'textbox-host-cell',
      sourceBlockIndex: 0,
      anchorBlockIndex: 0,
      tableId: floatingChild.id,
      overlap: 'overlap',
      positioning: {},
      anchorBounds: rect(2, 5),
      child: floatingChild,
    } as unknown as FloatingTablePlacementLayout;
    const resolvedPlacement: ResolvedFloatingTablePlacementLayout = {
      kind: 'resolved-floating-table-placement',
      occurrenceId: sourcePlacement.occurrenceId,
      xPt: 2,
      yPt: 5,
      bounds: rect(2, 5),
      exclusionBounds: rect(1, 4, 12, 12),
      overlap: 'overlap',
      child: floatingChild,
      source: sourcePlacement,
    };
    const hostTable: TableLayout = {
      ...table(),
      id: 'textbox-floating-host',
      flowBounds: rect(0, -20, 80, 60),
      inkBounds: rect(0, -20, 80, 60),
      advancePt: 60,
      floatingTables: [sourcePlacement],
      resolvedFloatingTables: [resolvedPlacement],
    };
    const retainedStory: StoryLayout = {
      story: 'textbox',
      flowBounds: hostTable.flowBounds,
      inkBounds: hostTable.inkBounds,
      blocks: [hostTable],
      advancePt: 60,
      diagnostics: [],
    };
    const shape = {
      type: 'shape',
      textVert: 'eaVert',
      textAutofit: 'sp',
      textInsetL: 0,
      textInsetT: 0,
      textInsetR: 0,
      textInsetB: 0,
    } as unknown as ShapeRun;
    const layout = acquireShapeTextBoxLayout(
      shape,
      { xPt: 0, yPt: 0, widthPt: 40, heightPt: 80 },
      {
        id: 'shape:vertical-float',
        source: { story: 'textbox', storyInstance: 'shape:vertical-float', path: [] },
        flowDomainId: 'shape-owner',
        context: paragraphContext,
        measurer: { context: canvas(), fontFamilyClasses: {} },
        environment: {
          pageIndex: 0,
          totalPages: 1,
          pageWritingMode: 'horizontal-tb',
          documentHasEastAsianText: false,
        },
        input: {
          kind: 'complete',
          source: { story: 'textbox', storyInstance: 'shape:vertical-float', path: [] },
          blockCount: 1,
        },
        acquireCompleteStory: () => retainedStory,
      },
    );

    expect(layout?.flowBounds.widthPt).toBe(60);
    const shiftedHost = layout?.story.blocks[0];
    expect(shiftedHost?.kind).toBe('table');
    if (shiftedHost?.kind !== 'table') throw new Error('expected shifted table');
    const shiftedSource = shiftedHost.floatingTables?.[0];
    const shiftedResolved = shiftedHost.resolvedFloatingTables?.[0];
    expect(shiftedSource?.anchorBounds.yPt).toBe(-5);
    expect(shiftedResolved).toMatchObject({
      xPt: 2,
      yPt: -5,
      bounds: { yPt: -5 },
      exclusionBounds: { yPt: -6 },
      child: { flowBounds: { yPt: -5 } },
    });
    expect(shiftedResolved?.source).toBe(shiftedSource);
    expect(shiftedResolved?.child).toBe(shiftedSource?.child);
  });

  it('uses the production shared story algorithms for paragraph/table order', () => {
    const textParagraph = (text: string): DocParagraph => ({
      type: 'paragraph',
      alignment: 'left',
      indentLeft: 0,
      indentRight: 0,
      indentFirst: 0,
      spaceBefore: 0,
      spaceAfter: 0,
      lineSpacing: null,
      numbering: null,
      tabStops: [],
      runs: [{
        type: 'text',
        text,
        bold: false,
        italic: false,
        underline: false,
        strikethrough: false,
        fontSize: 10,
        color: null,
        fontFamily: 'serif',
        fontFamilyEastAsia: '',
        isLink: false,
        background: null,
        vertAlign: null,
        hyperlink: null,
      }],
      defaultFontSize: 10,
      defaultFontFamily: 'serif',
      widowControl: false,
    } as unknown as DocParagraph);
    const tableWithContent = (
      content: readonly BodyElement[],
      widthPt: number,
    ): Extract<BodyElement, { type: 'table' }> => ({
      type: 'table',
      colWidths: [widthPt],
      rows: [{
        cells: [{
          content,
          colSpan: 1,
          vMerge: null,
          borders: {
            top: null, bottom: null, left: null, right: null,
            insideH: null, insideV: null,
          },
          background: null,
          vAlign: 'top',
          widthPt,
        }],
        rowHeight: null,
        rowHeightRule: 'auto',
        isHeader: false,
      }],
      borders: {
        top: null, bottom: null, left: null, right: null,
        insideH: null, insideV: null,
      },
      cellMarginTop: 0,
      cellMarginBottom: 0,
      cellMarginLeft: 0,
      cellMarginRight: 0,
      jc: 'left',
      layout: 'fixed',
      overlap: 'overlap',
      tblpPr: null,
    } as unknown as Extract<BodyElement, { type: 'table' }>);
    const nestedTable = tableWithContent([
      textParagraph('nested cell') as unknown as BodyElement,
    ], 40);
    const tableBlock = tableWithContent([
      textParagraph('cell before nested') as unknown as BodyElement,
      nestedTable,
      textParagraph('cell after nested') as unknown as BodyElement,
    ], 60);
    const floatingParagraph = textParagraph('before');
    floatingParagraph.runs.push({
      type: 'shape',
      widthPt: 18,
      heightPt: 12,
      anchorXPt: 35,
      anchorYPt: 0,
      anchorXFromMargin: false,
      anchorYFromPara: true,
      wrapMode: 'square',
      zOrder: 2,
      subpaths: [],
      presetGeometry: 'rect',
      fill: 'D9EAF7',
      stroke: null,
    } as unknown as DocRun);
    const mathParagraph = textParagraph('');
    mathParagraph.runs = [{
      type: 'math',
      nodes: [{ type: 'text', text: 'x + y' }],
      display: false,
      fontSize: 10,
    } as unknown as DocRun];
    const shape = {
      type: 'shape',
      widthPt: 100,
      heightPt: 20,
      anchorXPt: 0,
      anchorYPt: 0,
      anchorXFromMargin: false,
      anchorYFromPara: true,
      zOrder: 0,
      subpaths: [],
      presetGeometry: 'rect',
      fill: null,
      stroke: null,
      wrapMode: 'none',
      textInsetL: 2,
      textInsetT: 2,
      textInsetR: 2,
      textInsetB: 2,
      textAutofit: 'sp',
      textBlocks: [{ text: 'lossy fallback', fontSizePt: 10 }],
      textBoxContent: [
        floatingParagraph,
        tableBlock,
        {
          type: 'unsupportedTextBoxBlock',
          qName: 'w:altChunk',
          sourcePath: [2],
        },
        textParagraph('after'),
        mathParagraph,
      ],
    } as unknown as ShapeRun;
    const verticalTableBlock = tableWithContent([
      textParagraph('表の前') as unknown as BodyElement,
      tableWithContent([
        textParagraph('内側') as unknown as BodyElement,
      ], 35),
      textParagraph('表の後') as unknown as BodyElement,
    ], 55);
    const verticalShape = {
      ...shape,
      anchorXPt: 110,
      textVert: 'eaVert',
      textBlocks: [{ text: 'lossy vertical fallback', fontSizePt: 10 }],
      textBoxContent: [
        textParagraph('縦書き'),
        verticalTableBlock,
      ],
    } as unknown as ShapeRun;
    const host = textParagraph('');
    host.runs = [
      shape as unknown as DocRun,
      verticalShape as unknown as DocRun,
    ];
    const model = {
      section: {
        pageWidth: 240,
        pageHeight: 160,
        marginTop: 10,
        marginRight: 10,
        marginBottom: 10,
        marginLeft: 10,
        headerDistance: 5,
        footerDistance: 5,
        titlePage: false,
        evenAndOddHeaders: false,
        sectionStart: 'nextPage',
        columns: null,
      },
      body: [host],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [],
      endnotes: [],
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;

    const normalized = normalizeInternalDocumentModel(model);
    const layout = layoutDocument(
      normalized.document,
      createLayoutServices(normalized.document, { measureContext: canvas() }),
      { currentDateMs: 0 },
    );
    const bodyParagraph = layout.pages[0]?.layers.body.find(
      (node) => node.kind === 'paragraph',
    );
    if (!bodyParagraph || bodyParagraph.kind !== 'paragraph') {
      throw new Error('expected retained body paragraph');
    }
    const textBox = bodyParagraph.textBoxes[0];

    expect(textBox?.flowBounds.heightPt).toBeGreaterThan(20);
    expect(textBox?.clipBounds).toBeUndefined();
    expect(textBox?.story.blocks.map((block) => block.kind)).toEqual([
      'paragraph', 'table', 'paragraph', 'paragraph',
    ]);
    expect(textBox?.story.blocks.map((block) => block.source.path)).toEqual([
      [0], [1], [3], [4],
    ]);
    expect(normalized.mathOccurrences).toEqual([
      expect.objectContaining({
        source: {
          story: 'textbox',
          storyInstance: 'body:body:0.0',
          path: [4, 0],
        },
      }),
    ]);
    expect(textBox?.story.diagnostics).toEqual([{
      code: 'UNSUPPORTED_FEATURE',
      severity: 'warning',
      source: expect.objectContaining({ story: 'textbox', path: [2] }),
      message: 'Unsupported text-box block w:altChunk',
    }]);
    expect(layout.diagnostics).toEqual(textBox?.story.diagnostics);
    expect(textBox?.story.blocks[0]).toMatchObject({
      kind: 'paragraph',
      lines: [expect.objectContaining({
        placements: expect.arrayContaining([
          expect.objectContaining({ kind: 'text', text: 'before' }),
        ]),
      })],
    });
    const retainedFloatingParagraph = textBox?.story.blocks[0];
    expect(retainedFloatingParagraph?.kind).toBe('paragraph');
    if (retainedFloatingParagraph?.kind !== 'paragraph') {
      throw new Error('expected retained floating-content paragraph');
    }
    expect(retainedFloatingParagraph.drawings).toHaveLength(1);
    expect(retainedFloatingParagraph.exclusions).toEqual([
      expect.objectContaining({ wrap: 'square' }),
    ]);
    const retainedTable = textBox?.story.blocks[1];
    expect(retainedTable?.kind).toBe('table');
    if (retainedTable?.kind !== 'table') throw new Error('expected retained text-box table');
    expect(retainedTable.rows[0]?.cells[0]?.blocks.map((block) => block.layout.kind))
      .toEqual(['paragraph', 'table', 'paragraph']);

    const verticalTextBox = bodyParagraph.textBoxes[1];
    expect(verticalTextBox).toMatchObject({
      verticalMode: 'eaVert',
      transform: { a: 0, b: 1, c: -1, d: 0 },
    });
    expect(verticalTextBox?.story.blocks.map((block) => block.kind))
      .toEqual(['paragraph', 'table']);
    const verticalParagraph = verticalTextBox?.story.blocks[0];
    expect(verticalParagraph?.kind).toBe('paragraph');
    const topLevelOrientations = verticalParagraph?.kind === 'paragraph'
      ? verticalParagraph.lines.flatMap((line) => line.placements)
          .flatMap((placement) => placement.kind === 'text'
            ? placement.paintOps.map((operation) => operation.glyphOrientation)
            : [])
      : [];
    expect(topLevelOrientations).toContain('upright');
    const verticalTable = verticalTextBox?.story.blocks[1];
    expect(verticalTable?.kind).toBe('table');
    const cellParagraph = verticalTable?.kind === 'table'
      ? verticalTable.rows[0]?.cells[0]?.blocks[0]?.layout
      : undefined;
    expect(cellParagraph?.kind).toBe('paragraph');
    const cellOrientations = cellParagraph?.kind === 'paragraph'
      ? cellParagraph.lines.flatMap((line) => line.placements)
          .flatMap((placement) => placement.kind === 'text'
            ? placement.paintOps.map((operation) => operation.glyphOrientation)
            : [])
      : [];
    expect(cellOrientations).toContain('upright');
  });
});

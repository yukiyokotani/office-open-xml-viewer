import { describe, it, expect } from 'vitest';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import { layoutBodyModel } from './test-support/document-layout.test-support.js';
import { createLayoutServices } from './layout-runtime.js';
import { acquireShapeTextBoxLayout } from './layout/paragraph.js';
import { normalizeInternalDocumentModel, textBoxAcquisitionInput } from './parser-model.js';
import { layoutDocument } from './document-layout.js';
import type { ParagraphLayoutContext } from './layout-context.js';
import { wordCjkClosingMarkHangRefused } from './layout/line-compatibility.js';
import type { LineLayout, ParagraphLayout, TextPlacement } from './layout/types.js';
import type {
  BodyElement, CellElement, DocParagraph, DocSettings, DocTable, DocTableCell, DocTableRow,
  DocxDocumentModel, DocxTextRun, SectionProps, ShapeRun,
} from './types';

// word-justification-punctuation-hang (#2038): synthetic Windows Word
// controls, every jc in compatibility modes unset, 11, 12, 14 and 15.

/** Stub advances in em: the space 0.25, everything else 0.5. */
function advanceEm(character: string): number {
  return character === ' ' ? 0.25 : 0.5;
}
function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].reduce((sum, character) => sum + advanceEm(character) * p, 0),
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

// A 360 pt band from x = 20 to x = 380: 36 CJK characters of 10 pt.
const section: SectionProps = {
  pageWidth: 400, pageHeight: 400,
  marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
  headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
};

function textRun(text: string, color: string | null = null): DocParagraph['runs'][number] {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 20, color, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  };
  return { type: 'text', ...run } as DocParagraph['runs'][number];
}

function para(runs: DocParagraph['runs'], alignment: DocParagraph['alignment']): BodyElement {
  const p: DocParagraph = {
    alignment, indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs, defaultFontSize: 20, defaultFontFamily: 'NotInMetrics',
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

function lines(body: BodyElement, settings?: DocSettings): LineLayout[] {
  const node = layoutBodyModel([body], section, makeCtx(), {}, settings).pages[0]!.layers.body[0]!;
  if (node.kind !== 'paragraph') throw new Error('expected a paragraph');
  return [...node.lines];
}

function characters(line: LineLayout): Array<{ character: string; xPt: number }> {
  return line.placements.filter((p): p is TextPlacement => p.kind === 'text')
    .flatMap((placement) => {
      const characters = [...placement.text];
      expect(placement.clusters).toHaveLength(characters.length);
      return placement.clusters.map((cluster, index) => ({
        character: characters[index]!,
        xPt: placement.origin.xPt + cluster.offset.xPt,
      }));
    })
    .filter(({ character }) => character.trim() !== '')
    .sort((a, b) => a.xPt - b.xPt);
}
const text = (line: LineLayout): string => characters(line).map(({ character }) => character).join('');

const KANA = 'あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよ';

describe('word-justification-punctuation-hang', () => {
  const overflowing = (mark: string) => para([textRun(KANA.slice(0, 36) + mark + 'らりるれろ')], 'left');
  const withAlignment = (body: BodyElement, alignment: string): BodyElement =>
    ({ ...body, alignment }) as BodyElement;

  it('hangs a line-end 。 under both but pushes it under the kashida types and thaiDistribute', () => {
    for (const settings of [{ compatibilityMode: 14 }, { compatibilityMode: 15 }] as DocSettings[]) {
      const [hung] = lines(withAlignment(overflowing('。'), 'both'), settings);
      expect(text(hung!)).toBe(KANA.slice(0, 36) + '。');
      for (const alignment of ['lowKashida', 'thaiDistribute']) {
        const [pushed, next] = lines(withAlignment(overflowing('。'), alignment), settings);
        expect(text(pushed!), alignment).toBe(KANA.slice(0, 35));
        expect(text(next!).startsWith(KANA[35]! + '。'), alignment).toBe(true);
      }
    }
  });

  it('pushes nonjustified closing marks only in the measured mode-15 kinsoku-on class', () => {
    // Ordinary left, previously unmeasured right, and Strict start exercise
    // distinct admission names without repeating their mode cross-product.
    for (const alignment of ['left', 'right', 'start']) {
      const [first, next] = lines(withAlignment(overflowing('。'), alignment), { compatibilityMode: 15 });
      expect(text(first!), alignment).toBe(KANA.slice(0, 35));
      expect(text(next!).startsWith(KANA[35]! + '。'), alignment).toBe(true);
    }
  });

  it('preserves paragraph-off, document-off and unmeasured-mode hanging behavior', () => {
    // False excludes both the pre-15 one-mark bound and the mode-15 push.
    for (const [alignment, mode, mark] of [['center', 14, '。」'], ['left', 15, '。']] as const) {
      const body = { ...withAlignment(overflowing(mark), alignment), kinsoku: false } as BodyElement;
      expect(text(lines(body, { compatibilityMode: mode })[0]!)).toBe(KANA.slice(0, 36) + mark);
    }
    expect(text(lines(withAlignment(overflowing('。」'), 'right'), { compatibilityMode: 15, kinsoku: false })[0]!))
      .toBe(KANA.slice(0, 36) + '。」');
    for (const mode of [13, 16]) {
      expect(text(lines(withAlignment(overflowing('。」'), 'center'), { compatibilityMode: mode })[0]!))
        .toBe(KANA.slice(0, 36) + '。」');
    }
  });

  it('pushes at the kashida fit widths: 324 pt (mediumKashida) and 288 pt (highKashida)', () => {
    for (const settings of [{ compatibilityMode: 14 }, { compatibilityMode: 15 }] as DocSettings[]) {
      for (const [alignment, fit] of [['mediumKashida', 32], ['highKashida', 28]] as const) {
        const body = para([textRun(KANA.slice(0, fit) + '。' + 'らりるれろ')], alignment);
        const [pushed, next] = lines(body, settings);
        expect(text(pushed!), alignment).toBe(KANA.slice(0, fit - 1));
        expect(text(next!).startsWith(KANA[fit - 1]! + '。'), alignment).toBe(true);
      }
    }
  });

  it('hangs at most one mark: 。」 past the margin pushes the character before them', () => {
    const [pushed, next] = lines(withAlignment(overflowing('。」'), 'both'), { compatibilityMode: 15 });
    expect(text(pushed!)).toBe(KANA.slice(0, 35));
    expect(text(next!).startsWith(KANA[35]! + '。」')).toBe(true);
  });

  it('judges the character after the mark across a run seam', () => {
    // 。 ends the first run; 」 starts a differently coloured second run.
    for (const alignment of ['both', 'left'] as const) {
      const body = para([textRun(KANA.slice(0, 36) + '。'), textRun('」らり', 'FF0000')], alignment);
      const [pushed, next] = lines(body, { compatibilityMode: 14 });
      expect(text(pushed!), alignment).toBe(KANA.slice(0, 35));
      expect(text(next!).startsWith(KANA[35]! + '。」'), alignment).toBe(true);
    }
  });

  it('keeps a Latin period hanging under the kashida types (outside the measured CJK scope)', () => {
    const body = para([textRun('a'.repeat(36) + '. more')], 'lowKashida');
    const [first] = lines(body, { compatibilityMode: 15 });
    expect(text(first!)).toBe('a'.repeat(36) + '.');
  });

  // The policy is admitted only for the measured class: ordinary left-to-right
  // paragraphs of the main body and of its horizontal table cells. Every
  // other paragraph keeps the previous (v0.89) overflowPunct behaviour, under
  // which 。 hangs even when 」 follows it.
  const PAIR = KANA.slice(0, 10) + '。」らりる';
  const kanaCount = (line: LineLayout): number => line.placements
    .filter((p): p is TextPlacement => p.kind === 'text')
    .reduce((sum, placement) => sum + [...placement.text].filter((c) => KANA.includes(c)).length, 0);

  /** Every retained paragraph whose text starts with あ, anywhere in a value. */
  function paragraphsWithKana(value: unknown, found: ParagraphLayout[] = [], seen = new Set<object>()): ParagraphLayout[] {
    if (value === null || typeof value !== 'object' || seen.has(value)) return found;
    seen.add(value);
    const candidate = value as Partial<ParagraphLayout>;
    if (candidate.kind === 'paragraph' && Array.isArray(candidate.lines)
      && candidate.lines[0] && text(candidate.lines[0]).startsWith('あ')) {
      found.push(candidate as ParagraphLayout);
    }
    for (const child of Object.values(value)) paragraphsWithKana(child, found, seen);
    return found;
  }

  /** Kana on the first line of the one あ paragraph laid out for `body`. */
  function firstLineKana(body: readonly BodyElement[], sectionProps: SectionProps = section, compatibilityMode = 15): number {
    const layout = layoutBodyModel(body, sectionProps, makeCtx(), {}, { compatibilityMode });
    const found = paragraphsWithKana(layout.pages[0]);
    expect(found).toHaveLength(1);
    return kanaCount(found[0]!.lines[0]!);
  }

  const cellOf = (content: CellElement[], textDirection?: string): DocTableCell => ({
    content, colSpan: 1, vMerge: null,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    background: null, vAlign: 'top', widthPt: 100,
    ...(textDirection ? { textDirection } : {}),
  } as DocTableCell);
  const tableOf = (cells: DocTableCell[], exactHeightPt: number | null = 100): BodyElement => ({
    type: 'table',
    colWidths: cells.map(() => 100),
    rows: [{ cells, rowHeight: exactHeightPt, rowHeightRule: exactHeightPt === null ? 'auto' : 'exact', isHeader: false } as DocTableRow],
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left', layout: 'fixed',
  } as unknown as BodyElement);
  const cellPara = (runs: DocParagraph['runs'], alignment: DocParagraph['alignment'] = 'both'): CellElement =>
    para(runs, alignment) as unknown as CellElement;

  it('admits horizontal table cells of the main body (control)', () => {
    for (const textDirection of [undefined, 'lrTb']) {
      expect(firstLineKana([tableOf([cellOf([cellPara([textRun(PAIR)])], textDirection)], null)]), String(textDirection)).toBe(9);
    }
  });

  it('reads the character after the mark only when the one-mark limit needs it', () => {
    const unread = () => { throw new Error('the following character was read'); };
    const kinsoku = { enabled: true, lineStartForbidden: new Set([0x300D]) };
    expect(wordCjkClosingMarkHangRefused(undefined, '。', unread, kinsoku, false)).toBe(false);
    expect(wordCjkClosingMarkHangRefused('none', '。', unread, kinsoku, false)).toBe(true);
    expect(wordCjkClosingMarkHangRefused('one', '.', unread, kinsoku, false)).toBe(false);
    expect(wordCjkClosingMarkHangRefused('one', '。', unread, { ...kinsoku, enabled: false }, false)).toBe(false);
    expect(wordCjkClosingMarkHangRefused('one', '。', () => '」', kinsoku, false)).toBe(true);
  });

  it('judges a supplementary-plane character after the mark as one scalar', () => {
    // 𠀋 (U+2000B) may start a line: the 。 before it still hangs.
    const body = para([textRun(KANA.slice(0, 36) + '。'), textRun('𠀋らり', 'FF0000')], 'both');
    expect(text(lines(body, { compatibilityMode: 15 })[0]!)).toBe(KANA.slice(0, 36) + '。');
  });

  it('never hangs a mark in a table cell of the main body (blk-0898-cjk-kincell)', () => {
    // A single 。 one slot past the cell edge: the body hangs it under both,
    // distribute and left; Word's cells push the character before it.
    const single = KANA.slice(0, 10) + '。らりる';
    for (const alignment of ['both', 'distribute', 'left', 'lowKashida'] as const) {
      expect(firstLineKana([tableOf([cellOf([cellPara([textRun(single)], alignment)])], null)]), alignment).toBe(9);
    }
    // New pre-15 centered-cell admission and Strict end in mode15.
    for (const [alignment, mode] of [['center', 14], ['end', 15]] as const) {
      expect(firstLineKana([tableOf([cellOf([cellPara([textRun(single)], alignment)])], null)], section, mode), `${alignment}/${mode}`).toBe(9);
    }
  });

  function storyFirstLineKana(where: 'header' | 'footer' | 'footnote' | 'endnote'): number {
    const content = [para([textRun(KANA.slice(0, 36) + '。らりる')], 'both')];
    const host = para([textRun('x')], 'left') as unknown as DocParagraph;
    if (where === 'footnote' || where === 'endnote') {
      host.runs.push({ ...host.runs[0]!, text: '1', noteRef: { kind: where, id: 'n1' } } as never);
    }
    const story = (kind: string) => where === kind ? { body: content } : null;
    const model = {
      section, body: [host as unknown as BodyElement],
      headers: { default: story('header'), first: null, even: null },
      footers: { default: story('footer'), first: null, even: null },
      footnotes: where === 'footnote' ? [{ id: 'n1', content }] : [],
      endnotes: where === 'endnote' ? [{ id: 'n1', content }] : [],
      fontFamilyClasses: {}, settings: { compatibilityMode: 15 },
    } as unknown as DocxDocumentModel;
    const layout = layoutDocument(model, createLayoutServices(model, { measureContext: makeCtx() }), { currentDateMs: 0 });
    const found = layout.pages.flatMap((page) => paragraphsWithKana(page));
    expect(found.length, where).toBeGreaterThan(0);
    return kanaCount(found[0]!.lines[0]!);
  }

  it('admits header and footnote paragraphs, which Word lays out as the body', () => {
    expect(storyFirstLineKana('header')).toBe(36);
    expect(storyFirstLineKana('footnote')).toBe(36);
    // Control: the same 。」 pair is pushed in both.
    const pairStory = (where: 'header' | 'footnote'): number => {
      const content = [para([textRun(KANA.slice(0, 36) + '。」らりる')], 'both')];
      const host = para([textRun('x')], 'left') as unknown as DocParagraph;
      if (where === 'footnote') host.runs.push({ ...host.runs[0]!, text: '1', noteRef: { kind: 'footnote', id: 'n1' } } as never);
      const model = {
        section, body: [host as unknown as BodyElement],
        headers: { default: where === 'header' ? { body: content } : null, first: null, even: null },
        footers: { default: null, first: null, even: null },
        footnotes: where === 'footnote' ? [{ id: 'n1', content }] : [], endnotes: [],
        fontFamilyClasses: {}, settings: { compatibilityMode: 15 },
      } as unknown as DocxDocumentModel;
      const layout = layoutDocument(model, createLayoutServices(model, { measureContext: makeCtx() }), { currentDateMs: 0 });
      return kanaCount(layout.pages.flatMap((page) => paragraphsWithKana(page))[0]!.lines[0]!);
    };
    expect(pairStory('header')).toBe(35);
    expect(pairStory('footnote')).toBe(35);
  });

  it('keeps the previous hang in footers and endnotes (unmeasured)', () => {
    for (const where of ['footer', 'endnote'] as const) {
      const content = [para([textRun(KANA.slice(0, 36) + '。」らりる')], 'both')];
      const host = para([textRun('x')], 'left') as unknown as DocParagraph;
      if (where === 'endnote') host.runs.push({ ...host.runs[0]!, text: '1', noteRef: { kind: 'endnote', id: 'n1' } } as never);
      const model = {
        section, body: [host as unknown as BodyElement],
        headers: { default: null, first: null, even: null },
        footers: { default: where === 'footer' ? { body: content } : null, first: null, even: null },
        footnotes: [], endnotes: where === 'endnote' ? [{ id: 'n1', content }] : [],
        fontFamilyClasses: {}, settings: { compatibilityMode: 15 },
      } as unknown as DocxDocumentModel;
      const layout = layoutDocument(model, createLayoutServices(model, { measureContext: makeCtx() }), { currentDateMs: 0 });
      const found = layout.pages.flatMap((page) => paragraphsWithKana(page));
      expect(found.length, where).toBeGreaterThan(0);
      expect(kanaCount(found[0]!.lines[0]!), where).toBe(36);
    }
  });

  it('limits pre-15 centered and right-aligned paragraphs to one hanging mark', () => {
    // Each positively admitted legacy mode, plus Strict end's alias, must
    // enter the one-mark policy rather than the old unlimited hanging path.
    for (const [alignment, mode] of [['center', undefined], ['center', 11], ['center', 12], ['center', 14], ['end', 14]] as const) {
      const [first, next] = lines(withAlignment(overflowing('。」'), alignment), { compatibilityMode: mode });
      expect(text(first!), `${alignment}/${mode}`).toBe(KANA.slice(0, 35));
      expect(text(next!).startsWith(KANA[35]! + '。」'), `${alignment}/${mode}`).toBe(true);
    }
  });

  it('keeps the previous hang in frame paragraphs', () => {
    const framed = {
      ...(para([textRun(PAIR)], 'both') as object),
      framePr: {
        dropCap: 'none', lines: 1, wrap: 'around', hAnchor: 'margin', vAnchor: 'margin',
        hRule: 'auto', hSpace: 0, vSpace: 0, w: 100, x: 0, y: 0,
      },
    } as unknown as BodyElement;
    expect(firstLineKana([framed, para([textRun('x')], 'left')])).toBe(10);
  });

  it('keeps the previous hang in rotated and other authored-vertical table cells', () => {
    for (const direction of ['tbRl', 'btLr', 'tbRlV', 'tbLrV', 'lrTbV']) {
      expect(firstLineKana([tableOf([cellOf([cellPara([textRun(PAIR)])], direction)])]), direction).toBe(10);
    }
  });

  it('keeps the previous hang in a table nested in an authored-vertical cell', () => {
    // A nested table keeps the outer cell unrotated, but its direction is
    // still authored vertical.
    const nested = tableOf([cellOf([cellPara([textRun(PAIR)])])], null) as unknown as CellElement;
    expect(firstLineKana([tableOf([cellOf([nested], 'tbRl')], null)])).toBe(10);
    // Control: the same nesting in a horizontal outer cell is admitted.
    expect(firstLineKana([tableOf([cellOf([nested])], null)])).toBe(9);
  });

  it('restates the admission when a page-dependent cell paragraph is re-acquired', () => {
    const page = {
      type: 'field', fieldType: 'page', instruction: 'PAGE', fallbackText: '1',
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'NotInMetrics', background: null, vertAlign: null,
    } as unknown as DocParagraph['runs'][number];
    const content = [cellPara([textRun(PAIR), page])];
    // The admission is read from the source cells, not from a flag the
    // pagination caller passes: an unrotated vertical cell (tbLrV) carries none.
    for (const direction of ['tbRl', 'tbLrV']) {
      expect(firstLineKana([tableOf([cellOf(content, direction)])]), direction).toBe(10);
    }
    expect(firstLineKana([tableOf([cellOf(content)], null)])).toBe(9);
  });

  it('keeps the previous hang in an upright table of a vertical section', () => {
    const vertical = { ...section, textDirection: 'tbRl' } as SectionProps;
    expect(firstLineKana([tableOf([cellOf([cellPara([textRun(PAIR)])])], null)], vertical)).toBe(10);
  });

  function textBoxLines(
    hostPolicy: ParagraphLayoutContext['cjkClosingMarkHang'],
    childAlignment: string,
    childText: string,
    options: { bidi?: boolean; textVert?: string; hostOverflowPunct?: boolean } = {},
  ): LineLayout[] {
    const ctx = makeCtx();
    const model = {
      section, body: [], headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null }, fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: ctx });
    const context: ParagraphLayoutContext = {
      lineGrid: { active: false, pitchPt: null },
      characterGrid: { active: false, kind: null, pitchPt: null, deltaPt: 0 },
      rightIndentGrid: { pitchPt: null, paragraphAllowsAdjustment: true },
      physicalIndentLeftPt: 0, physicalIndentRightPt: 0, firstIndentPt: 0,
      lineSpacing: null, spaceBeforePt: 0, spaceAfterPt: 0,
      baseRtl: false, isJustified: false, stretchLastLine: false,
      tabStops: [], hasRuby: false, hasEastAsianText: true,
      kinsoku: DEFAULT_KINSOKU_RULES, defaultTabPt: 36,
      overflowPunct: options.hostOverflowPunct ?? true,
      ...(hostPolicy ? { cjkClosingMarkHang: hostPolicy } : {}),
    };
    const shape = {
      type: 'shape', zOrder: 0, subpaths: [], presetGeometry: 'rect', fill: null, stroke: null,
      textBlocks: [{
        text: childText, fontSizePt: 20, alignment: childAlignment,
        ...(options.bidi ? { bidi: true } : {}),
        runs: [{ text: childText, fontSizePt: 20, fontFamily: 'NotInMetrics' }],
      }],
      textAnchor: 't', textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
      textVert: options.textVert ?? null,
    } as unknown as ShapeRun;
    const source = { story: 'textbox' as const, storyInstance: 'hang', path: [] as number[] };
    // 100 pt along the line (both axes, so vertical boxes match): 10 characters.
    const layout = acquireShapeTextBoxLayout(shape, { xPt: 0, yPt: 0, widthPt: 100, heightPt: 100 }, {
      id: 'hang-box', source, flowDomainId: 'hang', context,
      measurer: { context: ctx, fontFamilyClasses: {} },
      environment: {
        pageIndex: 0, totalPages: 1, pageWritingMode: 'horizontal-tb',
        documentHasEastAsianText: true, layoutServices: services,
      },
      input: textBoxAcquisitionInput(shape, source),
    });
    const block = layout?.story.blocks[0];
    if (block?.kind !== 'paragraph') throw new Error('expected a text-box paragraph');
    return [...block.lines];
  }

  it('never gives a compatibility text-box paragraph the host\'s policy', () => {
    for (const hostPolicy of ['none', 'one'] as const) {
      for (const textVert of [undefined, 'vert', 'eaVert']) {
        expect(kanaCount(textBoxLines(hostPolicy, 'both', PAIR, { textVert })[0]!), `${hostPolicy} ${textVert}`).toBe(10);
        expect(kanaCount(textBoxLines(hostPolicy, 'lowKashida', PAIR, { textVert })[0]!), `${hostPolicy} ${textVert}`).toBe(10);
      }
      expect(kanaCount(textBoxLines(hostPolicy, 'both', PAIR, { bidi: true })[0]!)).toBe(10);
    }
  });

  it('keeps a compatibility text box on the host\'s overflowPunct', () => {
    // overflowPunct off in the host: the Latin period does not hang.
    const [first] = textBoxLines(undefined, 'left', 'a'.repeat(10) + '. more', { hostOverflowPunct: false });
    expect(text(first!)).not.toContain('.');
    // Control: overflowPunct on hangs it.
    const [hung] = textBoxLines(undefined, 'left', 'a'.repeat(10) + '. more');
    expect(text(hung!)).toBe('a'.repeat(10) + '.');
  });

  it('keeps the previous hang in complete text-box stories, horizontal or vertical', () => {
    const boxedKana = (textVert: string | null): number => {
      const shape = {
        type: 'shape', zOrder: 0, subpaths: [], presetGeometry: 'rect', fill: null, stroke: null,
        textBlocks: [], textAnchor: 't', textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
        textVert, widthPt: 100, heightPt: 100,
        textBoxContent: [para([textRun(PAIR)], 'both')],
      };
      const normalized = normalizeInternalDocumentModel({
        section, body: [para([shape as unknown as DocParagraph['runs'][number]], 'both')],
        headers: { default: null, first: null, even: null },
        footers: { default: null, first: null, even: null }, fontFamilyClasses: {},
      } as unknown as DocxDocumentModel);
      const layout = layoutDocument(normalized.document,
        createLayoutServices(normalized.document, { measureContext: makeCtx() }), { currentDateMs: 0 });
      const host = layout.pages[0]!.layers.body.find((node) => node.kind === 'paragraph');
      if (host?.kind !== 'paragraph') throw new Error('expected the host paragraph');
      const block = host.textBoxes[0]?.story.blocks[0];
      if (block?.kind !== 'paragraph') throw new Error('expected a text-box paragraph');
      return kanaCount(block.lines[0]!);
    };
    for (const textVert of [null, 'vert', 'vert270', 'eaVert']) expect(boxedKana(textVert), String(textVert)).toBe(10);
  });
});

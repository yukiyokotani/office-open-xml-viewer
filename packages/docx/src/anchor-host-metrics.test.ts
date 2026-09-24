import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import { describe, expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import {
  buildSegments,
  layoutLines,
  lineBoxHeight,
  type LayoutSeg,
  type LayoutTextSeg,
} from './line-layout.js';
import type { DocRun, DocxDocumentModel } from './types.js';

function metricServices(unicodeRanges: readonly (readonly [number, number])[]) {
  const empty = { default: null, first: null, even: null };
  return createLayoutServices({
    section: {
      pageWidth: 612, pageHeight: 792, marginTop: 72, marginRight: 72,
      marginBottom: 72, marginLeft: 72, headerDistance: 36, footerDistance: 36,
      titlePage: false, evenAndOddHeaders: false,
    },
    body: [], headers: empty, footers: empty,
  } as DocxDocumentModel, {
    measureContext: linearCtx(),
    localMetrics: {
      'yu mincho': {
        family: 'Yu Mincho', requestedFamily: 'Yu Mincho',
        lineHeightRatio: 1.5, designAscentRatio: 1.1,
        designDescentRatio: 0.4, unicodeRanges,
        sourceIdentity: 'test-resource:yu-mincho',
      },
    },
  });
}

function linearCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = (): number => Number.parseFloat(/([\d.]+)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText: (text: string) => ({
      width: [...text].length * px() * 0.5,
      fontBoundingBoxAscent: px() * 0.8,
      fontBoundingBoxDescent: px() * 0.2,
      actualBoundingBoxAscent: px() * 0.8,
      actualBoundingBoxDescent: px() * 0.2,
    } as TextMetrics),
  } as unknown as CanvasRenderingContext2D;
}

function anchorHost(): DocRun {
  return {
    type: 'anchorHost',
    fontSize: 20,
    fontFamily: 'Arial',
    fontFamilyEastAsia: 'Yu Mincho',
    bold: false,
    italic: false,
  } as DocRun;
}

function textRun(text: string, fontSize = 11): DocRun {
  return {
    type: 'text',
    text,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    fontSize,
    color: null,
    fontFamily: 'Yu Mincho',
    fontFamilyEastAsia: 'Yu Mincho',
    isLink: false,
    background: null,
    vertAlign: null,
  } as DocRun;
}

function layOut(runs: DocRun[]) {
  return layoutLines(
    linearCtx(),
    buildSegments(runs, { pageIndex: 0, totalPages: 1 })
      .map((segment) => ({ ...segment })) as LayoutSeg[],
    400,
    0,
    1,
    undefined,
    undefined,
    {},
    undefined,
    DEFAULT_KINSOKU_RULES,
    undefined,
    36,
    400,
    false,
    false,
    false,
  );
}

describe('floating drawing anchor-host metrics', () => {
  it('admits selected host geometry for auto and atLeast, but not exact spacing', () => {
    const covered = metricServices([[0x3042, 0x3042]]);
    const partial = metricServices([[0x41, 0x5a]]);
    const segment = (layoutServices: ReturnType<typeof metricServices>, rule?: 'exact' | 'atLeast') =>
      buildSegments([anchorHost()], {
        pageIndex: 0, totalPages: 1, layoutServices,
        ...(rule ? { lineSpacing: { rule, value: 18, explicit: true } } : {}),
      })[0] as LayoutTextSeg;
    expect(segment(covered).resolvedLineHeightRatio).toBe(1.5);
    expect(segment(partial).resolvedLineHeightRatio).toBeUndefined();
    expect(segment(covered, 'exact').resolvedLineHeightRatio).toBeUndefined();
    expect(segment(covered, 'atLeast').resolvedLineHeightRatio).toBe(1.5);
  });

  it('emits a zero-width metric segment using the anchor character formatting', () => {
    const segments = buildSegments([anchorHost()], { pageIndex: 0, totalPages: 1 });

    expect(segments).toHaveLength(1);
    expect(segments[0]).toMatchObject({
      text: '',
      measuredWidth: 0,
      fontSize: 20,
      fontFamily: 'Yu Mincho',
      metricOnly: true,
      metricEastAsian: true,
    });
  });

  it('uses the zero-width metric segment to allocate East Asian document-grid cells', () => {
    const segments = buildSegments([anchorHost()], { pageIndex: 0, totalPages: 1 });
    const [line] = layoutLines(
      linearCtx(),
      segments.map((segment) => ({ ...segment })) as LayoutSeg[],
      400,
      0,
      1,
      undefined,
      undefined,
      {},
      undefined,
      DEFAULT_KINSOKU_RULES,
      undefined,
      36,
      400,
      false,
      false,
      false,
    );

    expect((line.segments[0] as LayoutTextSeg).measuredWidth).toBe(0);
    expect(line.eastAsian).toBe(true);
    expect(lineBoxHeight(
      null,
      line.ascent,
      line.descent,
      1,
      { type: 'lines', linePitchPt: 18 },
      false,
      line.intendedSingle,
      line.eastAsian,
      line.gridCountSingle,
    )).toBe(36);
  });

  it('reserves host line height without using its zero-ink box for a visible run baseline', () => {
    const [hostOnly] = layOut([anchorHost()]);
    const [visibleOnly] = layOut([textRun('Caption')]);
    const [mixed] = layOut([anchorHost(), textRun('Caption')]);

    // The host formatting still reserves the same two-cell line advance used by
    // anchor-only rows; only the visible glyph baseline comes from visible ink.
    expect(mixed.height).toBe(hostOnly.height);
    expect(mixed.ascent).toBe(hostOnly.ascent);
    expect(mixed.descent).toBe(hostOnly.descent);
    expect(mixed.intendedSingle).toBe(hostOnly.intendedSingle);
    expect(mixed.gridCountSingle).toBe(hostOnly.gridCountSingle);
    expect(mixed.visibleAscent).toBe(visibleOnly.ascent);
    expect(mixed.visibleDescent).toBe(visibleOnly.descent);
    expect(hostOnly.visibleAscent).toBe(hostOnly.ascent);
    expect(hostOnly.visibleDescent).toBe(hostOnly.descent);
  });
});

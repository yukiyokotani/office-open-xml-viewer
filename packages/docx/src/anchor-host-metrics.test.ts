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
import { isAnchorOnlyParagraph } from './layout/paragraph-visibility.js';

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

/** The label font's zero-advance metric run before a picture list label: the
 * one anchor-host run that keeps line metrics (it anchors nothing). */
function labelMetricRun(): DocRun {
  return { ...anchorHost(), listLabel: 'start' } as unknown as DocRun;
}

function layOut(runs: DocRun[], layoutServices?: ReturnType<typeof metricServices>) {
  return layoutLines(
    linearCtx(),
    buildSegments(runs, { pageIndex: 0, totalPages: 1,
      ...(layoutServices ? { layoutServices } : {}) })
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
  it('admits selected label-metric geometry for auto and atLeast, but not exact spacing', () => {
    const covered = metricServices([[0x3042, 0x3042]]);
    const partial = metricServices([[0x41, 0x5a]]);
    const segment = (layoutServices: ReturnType<typeof metricServices>, rule?: 'exact' | 'atLeast') =>
      buildSegments([labelMetricRun()], {
        pageIndex: 0, totalPages: 1, layoutServices,
        ...(rule ? { lineSpacing: { rule, value: 18, explicit: true } } : {}),
      })[0] as LayoutTextSeg;
    expect(segment(covered).resolvedLineHeightRatio).toBe(1.5);
    expect(segment(partial).resolvedLineHeightRatio).toBeUndefined();
    expect(segment(covered, 'exact').resolvedLineHeightRatio).toBeUndefined();
    expect(segment(covered, 'atLeast').resolvedLineHeightRatio).toBe(1.5);
  });

  it('emits a zero-width metric segment using the label metric run formatting', () => {
    const segments = buildSegments([labelMetricRun()], { pageIndex: 0, totalPages: 1 });

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
    const segments = buildSegments([labelMetricRun()], { pageIndex: 0, totalPages: 1 });
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

  it('reserves the label metric height without using its zero-ink box for a visible run baseline', () => {
    const [hostOnly] = layOut([labelMetricRun()]);
    const [visibleOnly] = layOut([textRun('Caption')]);
    const [mixed] = layOut([labelMetricRun(), textRun('Caption')]);

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

  describe('WORD_ANCHOR_HOST_NO_LINE_METRIC', () => {
    const picture = { type: 'image', imagePath: 'p.png', mimeType: 'image/png', widthPt: 20, heightPt: 16,
      anchor: false, anchorXPt: 0, anchorYPt: 0, anchorXFromMargin: false, anchorYFromPara: false } as DocRun;

    it('sizes no line from the anchoring run beside other content', () => {
      const [withHost] = layOut([anchorHost(), picture]);
      const [withoutHost] = layOut([picture]);
      expect(withHost.ascent).toBe(withoutHost.ascent);
      expect(withHost.descent).toBe(withoutHost.descent);
      const [text] = layOut([textRun('Caption')]);
      const [textWithHost] = layOut([anchorHost(), textRun('Caption')]);
      expect(textWithHost.ascent).toBe(text.ascent);
      expect(textWithHost.descent).toBe(text.descent);
    });

    it('keeps a uniformly positioned line unchanged beside an anchor', () => {
      // word-uniform-run-position-leading needs face data: a resource face
      // covering the Latin text.
      const services = metricServices([[0x41, 0x7a]]);
      const lowered = { ...textRun('Lowered'), position: -3 } as DocRun;
      const floating = { ...picture, anchor: true } as DocRun;
      const [alone] = layOut([lowered], services);
      const [anchored] = layOut([anchorHost(), floating, lowered], services);
      expect(anchored.ascent).toBe(alone.ascent);
      expect(anchored.descent).toBe(alone.descent);
      expect(anchored.intendedSingle).toBe(alone.intendedSingle);
      expect(alone.uniformPositionAuto).toBeDefined();
      expect(anchored.uniformPositionAuto).toEqual(alone.uniformPositionAuto);
    });

    it('emits every host without metrics in one pass', () => {
      const runs = [...Array.from({ length: 400 }, () => anchorHost()), textRun('Tail')];
      let reads = 0;
      const counted = new Proxy(runs, {
        get(target, key, receiver) {
          if (typeof key === 'string' && /^\d+$/.test(key)) reads += 1;
          return Reflect.get(target, key, receiver);
        },
      });
      const segments = buildSegments(counted, { pageIndex: 0, totalPages: 1 });
      expect(segments.filter((segment) => 'text' in segment && segment.anchorHostWithoutMetric)).toHaveLength(400);
      // A per-host rescan reads the runs about 400 * 401 / 2 times.
      expect(reads).toBeLessThan(runs.length * 50);
    });

    // The classification and the segment builder share runIsOmitted: a run
    // the view omits emits nothing and is no content, whatever its type.
    describe('revision-omitted runs of every type', () => {
      const deletion = { revision: { kind: 'deletion', author: 'A' } };
      const insertion = { revision: { kind: 'insertion', author: 'A' } };
      const nonText: readonly [string, DocRun][] = [
        ['inline picture', picture],
        ['field', { type: 'field', fieldType: 'page', instruction: 'PAGE', fallbackText: '9', fontSize: 10,
          fontFamily: 'Arial', bold: false, italic: false, underline: false, strikethrough: false, color: null,
          vertAlign: null } as unknown as DocRun],
        ['math', { type: 'math', nodes: [{ kind: 'run', text: 'x' }], display: false, fontSize: 10 } as unknown as DocRun],
        ['line break', { type: 'break', breakType: 'line' } as DocRun],
        ['text', textRun('Gone')],
      ];
      it.each(nonText)('omits a deleted %s in the final view and keeps it in the markup view', (_, run) => {
        const deleted = { ...run, ...deletion } as DocRun;
        const final = buildSegments([deleted], { pageIndex: 0, totalPages: 1, showTrackedChanges: false });
        const markup = buildSegments([deleted], { pageIndex: 0, totalPages: 1, showTrackedChanges: true });
        expect(final).toHaveLength(0);
        expect(markup.length).toBeGreaterThan(0);
        expect(isAnchorOnlyParagraph([anchorHost(), deleted] as never, false)).toBe(true);
        expect(isAnchorOnlyParagraph([anchorHost(), deleted] as never, true)).toBe(false);
      });
      it.each(nonText)('keeps an inserted %s in both views', (_, run) => {
        const inserted = { ...run, ...insertion } as DocRun;
        for (const view of [false, true]) {
          expect(buildSegments([inserted], { pageIndex: 0, totalPages: 1, showTrackedChanges: view }).length)
            .toBeGreaterThan(0);
          expect(isAnchorOnlyParagraph([anchorHost(), inserted] as never, view)).toBe(false);
        }
      });
    });

    it('classifies a paragraph displaying only floating drawings (isAnchorOnlyParagraph)', () => {
      const deleted = { ...textRun('Gone'), revision: { kind: 'deletion', author: 'A' } } as DocRun;
      const softHyphen = textRun('\u00AD');
      const floating = { ...picture, anchor: true } as DocRun;
      for (const runs of [[anchorHost()], [anchorHost(), deleted], [anchorHost(), floating], [floating]]) {
        expect(isAnchorOnlyParagraph(runs as never, false)).toBe(true);
      }
      // What remains without the anchors decides, exactly as for the
      // paragraph without them: a soft hyphen is that paragraph's content.
      expect(isAnchorOnlyParagraph([anchorHost(), softHyphen] as never, false)).toBe(false);
      // Markup view shows the deletion; a list label's metric run is content.
      expect(isAnchorOnlyParagraph([anchorHost(), deleted] as never, true)).toBe(false);
      expect(isAnchorOnlyParagraph([{ ...anchorHost(), listLabel: 'start' }] as never, false)).toBe(false);
      for (const runs of [[anchorHost(), textRun('x')], [anchorHost(), picture], [anchorHost(), textRun(' ')], []]) {
        expect(isAnchorOnlyParagraph(runs as never, false)).toBe(false);
      }
    });
  });
});

import { describe, expect, it } from 'vitest';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import {
  collectBodyFrameGroups,
  effectiveFrameIdentity,
  measureParagraphIntrinsicWidth,
} from './frame.js';
import { acquireRetainedFrameGroup } from './paragraph.js';
import { paragraphAcquisitionInput } from '../parser-model.js';
import { resolveNumberingMarkerGeometry } from './numbering-marker.js';
import type { AnchorAcquisitionInput, AnchorEdgesInput } from './anchor-input.js';
import type { TextLayoutService } from './text.js';
import type { BodyElement, DocParagraph, FramePr } from '../types.js';

function frame(extra: Partial<FramePr> & { __anchorLock?: boolean } = {}): FramePr {
  return {
    dropCap: 'none', lines: 1, wrap: 'around', hAnchor: 'text', vAnchor: 'text',
    hRule: 'auto', hSpace: 0, vSpace: 0, ...extra,
  };
}

function paragraph(framePr?: FramePr): DocParagraph {
  return {
    type: 'paragraph', framePr,
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{
      type: 'text', text: 'frame', bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'serif',
      isLink: false, background: null, vertAlign: null, hyperlink: null,
    }],
    defaultFontSize: 10, defaultFontFamily: 'serif',
  } as unknown as DocParagraph;
}

const missingAnchorEdges = (): AnchorEdgesInput => ({
  topPt: null, topStatus: 'missing',
  rightPt: null, rightStatus: 'missing',
  bottomPt: null, bottomStatus: 'missing',
  leftPt: null, leftStatus: 'missing',
});

function framedParagraphWithTopAndBottomAnchor(): DocParagraph {
  const occurrenceId = 'frame-top-and-bottom';
  const acquisition: AnchorAcquisitionInput = {
    occurrenceId,
    simplePosition: {
      enabled: false, status: 'valid',
      xPt: 0, xStatus: 'valid', yPt: 0, yStatus: 'valid',
    },
    horizontal: {
      relativeFrom: 'column', relativeFromStatus: 'valid',
      choice: { kind: 'offset', valuePt: -10 },
    },
    vertical: {
      relativeFrom: 'paragraph', relativeFromStatus: 'valid',
      choice: { kind: 'offset', valuePt: 0 },
    },
    extent: {
      widthPt: 40, heightPt: 20,
      widthStatus: 'valid', heightStatus: 'valid',
    },
    parentEffectExtent: missingAnchorEdges(),
    anchorDistances: missingAnchorEdges(),
    relativeSize: { horizontal: null, vertical: null },
    wrap: {
      kind: 'topAndBottom', authoredKinds: ['wrapTopAndBottom'], side: null,
      distances: missingAnchorEdges(), effectExtent: null, polygon: null,
    },
    behavior: {
      behindDoc: false, behindDocStatus: 'valid',
      relativeHeight: 1, relativeHeightStatus: 'valid',
      locked: false, lockedStatus: 'valid',
      allowOverlap: true, allowOverlapStatus: 'valid',
      layoutInCell: true, layoutInCellStatus: 'valid',
    },
    group: null,
  };
  const text = (value: string) => ({
    type: 'text', text: value, bold: false, italic: false, underline: false,
    strikethrough: false, fontSize: 10, color: null, fontFamily: 'serif',
    isLink: false, background: null, vertAlign: null, hyperlink: null,
  });
  return {
    ...paragraph(frame({ w: 40 })),
    runs: [
      text('A'.repeat(9)),
      {
        type: 'anchorHost', fontSize: 10,
        __anchorOccurrenceId: occurrenceId,
      },
      {
        type: 'image',
        imagePath: 'word/media/frame-anchor.png',
        mimeType: 'image/png',
        widthPt: 40, heightPt: 20,
        anchor: true,
        __anchorAcquisition: acquisition,
      },
      text('B'.repeat(8)),
    ],
  } as unknown as DocParagraph;
}

function framedParagraphWithFollowingSquareAnchor(): DocParagraph {
  const occurrenceId = 'frame-following-square';
  const acquisition: AnchorAcquisitionInput = {
    occurrenceId,
    simplePosition: {
      enabled: false, status: 'valid',
      xPt: 0, xStatus: 'valid', yPt: 0, yStatus: 'valid',
    },
    horizontal: {
      relativeFrom: 'column', relativeFromStatus: 'valid',
      choice: { kind: 'offset', valuePt: 70 },
    },
    vertical: {
      relativeFrom: 'paragraph', relativeFromStatus: 'valid',
      choice: { kind: 'offset', valuePt: 0 },
    },
    extent: {
      widthPt: 80, heightPt: 60,
      widthStatus: 'valid', heightStatus: 'valid',
    },
    parentEffectExtent: missingAnchorEdges(),
    anchorDistances: missingAnchorEdges(),
    relativeSize: { horizontal: null, vertical: null },
    wrap: {
      kind: 'square', authoredKinds: ['wrapSquare'], side: 'bothSides',
      distances: missingAnchorEdges(), effectExtent: null, polygon: null,
    },
    behavior: {
      behindDoc: false, behindDocStatus: 'valid',
      relativeHeight: 1, relativeHeightStatus: 'valid',
      locked: false, lockedStatus: 'valid',
      allowOverlap: true, allowOverlapStatus: 'valid',
      layoutInCell: true, layoutInCellStatus: 'valid',
    },
    group: null,
  };
  return {
    ...paragraph(frame({ w: 160 })),
    runs: [
      {
        type: 'anchorHost', fontSize: 10,
        __anchorOccurrenceId: occurrenceId,
      },
      {
        type: 'image',
        imagePath: 'word/media/frame-following.png',
        mimeType: 'image/png',
        widthPt: 80, heightPt: 60,
        anchor: true,
        __anchorAcquisition: acquisition,
      },
      {
        type: 'text', text: 'A', bold: false, italic: false, underline: false,
        strikethrough: false, fontSize: 10, color: null, fontFamily: 'serif',
        isLink: false, background: null, vertAlign: null, hyperlink: null,
      },
    ],
  } as unknown as DocParagraph;
}

function context() {
  return {
    lineGrid: { active: false, pitchPt: null },
    characterGrid: { active: false, kind: null, pitchPt: null, deltaPt: 0 },
    rightIndentGrid: { pitchPt: null, paragraphAllowsAdjustment: true },
    physicalIndentLeftPt: 0, physicalIndentRightPt: 0, firstIndentPt: 0,
    lineSpacing: null, spaceBeforePt: 0, spaceAfterPt: 0,
    baseRtl: false, isJustified: false, stretchLastLine: false,
    tabStops: [], hasRuby: false, hasEastAsianText: false,
    kinsoku: DEFAULT_KINSOKU_RULES, defaultTabPt: 36,
  } as const;
}

function measureContext(): CanvasRenderingContext2D {
  return {
    font: '10px serif', letterSpacing: '0px',
    measureText(text: string) {
      return {
        width: text.length * 5,
        actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
        fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

function textService(fingerprint = 'frame-test-text'): TextLayoutService {
  return {
    fingerprint, localMetrics: {},
    resolve(request) {
      return {
        requestedFamily: request.fonts.ascii ?? 'serif',
        resolvedFamily: request.fonts.ascii ?? 'serif',
        route: { familyList: 'serif', scope: 'generic', fingerprint: 'serif' },
        source: 'generic', weight: request.weight ?? 400, style: request.style ?? 'normal',
        diagnostics: [], genericFamily: 'serif',
      };
    },
    shape(request) {
      const clusters = [...request.text].map((_character, index) => ({
        range: { start: index, end: index + 1 }, offsetPt: index * 5, advancePt: 5,
      }));
      const route = { familyList: 'serif', scope: 'generic', fingerprint: 'serif' } as const;
      return {
        text: request.text,
        spans: [{
          text: request.text, start: 0, end: request.text.length,
          script: 'ascii', breakBefore: true,
          font: {
            requestedFamily: 'serif', resolvedFamily: 'serif', route,
            source: 'generic', weight: request.weight ?? 400,
            style: request.style ?? 'normal', diagnostics: [], genericFamily: 'serif',
          },
          fontRoute: route, advancePt: request.text.length * 5,
          ascentPt: 8, descentPt: 2,
        }],
        advancePt: request.text.length * 5, ascentPt: 8, descentPt: 2,
        graphemeBoundaries: [0, request.text.length], clusters, diagnostics: [],
      };
    },
  };
}

function frameOptions(
  group: NonNullable<ReturnType<typeof collectBodyFrameGroups> extends WeakMap<DocParagraph, infer G> ? G : never>,
  paragraphs: readonly DocParagraph[],
  overrides: Partial<Parameters<typeof acquireRetainedFrameGroup>[1]> = {},
): Parameters<typeof acquireRetainedFrameGroup>[1] {
  const service = textService();
  return {
    contexts: paragraphs.map(() => context()),
    inputs: paragraphs.map((item, index) =>
      paragraphAcquisitionInput(item, { story: 'body', storyInstance: 'body', path: [index] })),
    borderEdges: paragraphs.map(() => undefined),
    borderExtentsPt: paragraphs.map(() => 0),
    measurer: { context: measureContext(), fontFamilyClasses: {} },
    maximumWidthPt: 180,
    placementSignature: 'column:10,20,180,280',
    place: (contentWidthPt: number, contentHeightPt: number) => ({
      bounds: { xPt: 10, yPt: 20, widthPt: contentWidthPt, heightPt: contentHeightPt },
      exclusionBounds: { xPt: 10, yPt: 20, widthPt: contentWidthPt, heightPt: contentHeightPt },
    }),
    environment: {
      pageIndex: 0, totalPages: 1, documentHasEastAsianText: false,
      pageWritingMode: 'horizontal-tb',
      layoutServices: {
        text: service,
        images: {
          fingerprint: 'images',
          resolve: () => ({ widthPt: 0, heightPt: 0, mimeType: 'application/octet-stream' }),
        },
        math: {
          fingerprint: 'math',
          resolve: (resourceKey: string) => ({
            resourceKey, widthEm: 0, ascentEm: 0, descentEm: 0, diagnostics: [],
          }),
        },
      },
    },
    anchorFrames: {
      page: { xPt: 0, yPt: 0, widthPt: 200, heightPt: 300 },
      margin: { xPt: 10, yPt: 10, widthPt: 180, heightPt: 280 },
      column: { xPt: 10, yPt: 10, widthPt: 180, heightPt: 280 },
      pageParity: 'odd' as const,
    },
    acquisitionSession: group,
    ...overrides,
  };
}

describe('retained text-frame grouping', () => {
  it('uses the complete effective CT_FramePr value including private anchorLock', () => {
    const base = frame();
    const variants: Array<readonly [string, Partial<FramePr> & { __anchorLock?: boolean }]> = [
      ['dropCap', { dropCap: 'drop' }],
      ['lines', { lines: 2 }],
      ['w', { w: 40 }],
      ['h', { h: 30 }],
      ['vSpace', { vSpace: 2 }],
      ['hSpace', { hSpace: 3 }],
      ['wrap', { wrap: 'none' }],
      ['hAnchor', { hAnchor: 'page' }],
      ['vAnchor', { vAnchor: 'margin' }],
      ['x', { x: 4 }],
      ['xAlign', { xAlign: 'right' }],
      ['y', { y: 5 }],
      ['yAlign', { yAlign: 'bottom' }],
      ['hRule', { hRule: 'exact' }],
      ['anchorLock', { __anchorLock: true }],
    ];
    for (const [attribute, difference] of variants) {
      const variant = frame(difference);
      expect(effectiveFrameIdentity(base), attribute).not.toBe(effectiveFrameIdentity(variant));
      const baseParagraph = paragraph(base);
      const variantParagraph = paragraph(variant);
      const groups = collectBodyFrameGroups([
        baseParagraph as unknown as BodyElement,
        variantParagraph as unknown as BodyElement,
      ]);
      expect(groups.get(baseParagraph)?.members, attribute).toEqual([baseParagraph]);
      expect(groups.get(variantParagraph)?.members, attribute).toEqual([variantParagraph]);
      expect(groups.get(baseParagraph), attribute).not.toBe(groups.get(variantParagraph));
    }
    expect(effectiveFrameIdentity(frame({ w: 40 }))).toBe(effectiveFrameIdentity(frame({ w: 40 })));
  });

  it('groups only adjacent paragraphs with identical effective frame properties', () => {
    const first = paragraph(frame({ w: 40 }));
    const second = paragraph(frame({ w: 40 }));
    const ordinary = paragraph();
    const third = paragraph(frame({ w: 40 }));
    const groups = collectBodyFrameGroups([
      first as unknown as BodyElement,
      second as unknown as BodyElement,
      ordinary as unknown as BodyElement,
      third as unknown as BodyElement,
    ]);

    expect(groups.get(first)?.members).toEqual([first, second]);
    expect(groups.get(second)).toBe(groups.get(first));
    expect(groups.get(third)?.members).toEqual([third]);
    expect(groups.get(ordinary)).toBeUndefined();
  });

  it('returns the same retained acquisition for an identical final placement signature', () => {
    const shared = frame({ w: 50 });
    const firstParagraph = paragraph({ ...shared });
    const secondParagraph = paragraph({ ...shared });
    const groups = collectBodyFrameGroups([
      firstParagraph as unknown as BodyElement,
      secondParagraph as unknown as BodyElement,
    ]);
    const group = groups.get(firstParagraph)!;
    const options = frameOptions(group, [firstParagraph, secondParagraph]);

    const first = acquireRetainedFrameGroup(group, options);
    const second = acquireRetainedFrameGroup(group, options);

    expect(second).toBe(first);
    expect(Object.isFrozen(second)).toBe(true);
    expect(Object.isFrozen(second.members)).toBe(true);
    expect(Object.isFrozen(second.members[0]?.fragment)).toBe(true);
    expect(second.members[0]?.fragment).toBe(first.members[0]?.fragment);
    expect(second.members[1]?.fragment).toBe(first.members[1]?.fragment);
  });

  it.each([
    ['LTR right-aligned marker', false, 'right' as const, 'nothing' as const, undefined, false, 65],
    ['LTR centered marker', false, 'center' as const, 'nothing' as const, undefined, false, 65],
    ['RTL right-aligned marker', true, 'right' as const, 'nothing' as const, undefined, false, 65],
    ['LTR tab suffix', false, 'right' as const, 'tab' as const, undefined, false, 81],
    ['RTL tab suffix', true, 'right' as const, 'tab' as const, undefined, false, 81],
    ['picture bullet', false, 'right' as const, 'nothing' as const, 60, false, 65],
    ['numbering-only paragraph', false, 'right' as const, 'nothing' as const, undefined, true, 60],
    ['picture-bullet-only paragraph', false, 'right' as const, 'nothing' as const, 60, true, 60],
  ])('includes final numbering geometry in auto-frame intrinsic width: %s', (
    _name, baseRtl, jc, suff, pictureWidthPt, emptyBody, expectedWidthPt,
  ) => {
    const markerInput = {
      fontSizePt: 10,
      fonts: { ascii: 'serif', eastAsia: 'serif', highAnsi: 'serif', complexScript: 'serif' },
      weight: 400,
      style: 'normal' as const,
      complexScript: false,
    };
    const numbered = {
      ...paragraph(),
      indentFirst: 20,
      runs: emptyBody ? [] : [{ ...paragraph().runs[0]!, text: 'x' }],
      numbering: {
        text: pictureWidthPt === undefined ? '123456789012' : '',
        jc, suff,
        ...(pictureWidthPt === undefined ? {} : {
          picBulletImagePath: 'bullet.png', picBulletWidthPt: pictureWidthPt,
        }),
      },
      numberingMarkerShapeInput: markerInput,
    } as unknown as DocParagraph;
    const service = textService();
    const geometry = resolveNumberingMarkerGeometry(
      numbered.numbering!,
      markerInput,
      {
        authoredFirstIndentPt: numbered.indentFirst,
        physicalIndentLeftPt: 0,
        tabStops: [],
        defaultTabPt: 36,
      },
      service,
    );
    const resolvedContext = {
      ...context(),
      baseRtl,
      firstIndentPt: geometry.bodyOffsetPt,
    };

    const widthPt = measureParagraphIntrinsicWidth(
      numbered,
      resolvedContext,
      180,
      { context: measureContext(), fontFamilyClasses: {} },
      {
        pageIndex: 0,
        totalPages: 1,
        pageWritingMode: 'horizontal-tb',
        documentHasEastAsianText: false,
        layoutServices: {
          text: service,
          images: { fingerprint: 'images', resolve: () => ({ widthPt: 0, heightPt: 0, mimeType: 'image/png' }) },
          math: { fingerprint: 'math', resolve: (resourceKey: string) => ({ resourceKey, widthEm: 0, ascentEm: 0, descentEm: 0, diagnostics: [] }) },
        },
      },
      geometry,
    );

    expect(widthPt).toBe(expectedWidthPt);
  });

  it('uses one logical-leading coordinate system for RTL auto width and retained marker placement', () => {
    const framed = {
      ...paragraph(frame()),
      indentFirst: 20,
      runs: [{ ...paragraph().runs[0]!, text: 'x' }],
      numbering: {
        text: '123456789012',
        jc: 'right',
        suff: 'nothing',
      },
    } as unknown as DocParagraph;
    const group = collectBodyFrameGroups([framed as unknown as BodyElement]).get(framed)!;
    const options = frameOptions(group, [framed]);
    const service = options.environment.layoutServices!.text;
    const markerInput = paragraphAcquisitionInput(
      framed,
      { story: 'body', storyInstance: 'body', path: [0] },
    ).numberingMarkerShapeInput!;
    const marker = resolveNumberingMarkerGeometry(framed.numbering!, markerInput, {
      authoredFirstIndentPt: framed.indentFirst,
      physicalIndentLeftPt: 7,
      tabStops: [],
      defaultTabPt: 36,
    }, service);
    const acquired = acquireRetainedFrameGroup(group, {
      ...options,
      contexts: [{
        ...context(),
        baseRtl: true,
        physicalIndentLeftPt: 7,
        physicalIndentRightPt: 19,
        firstIndentPt: marker.bodyOffsetPt,
      }],
    });

    expect(acquired.box.bounds.widthPt).toBe(72);
    const placement = acquired.members[0]?.fragment.lines
      .flatMap((line) => line.placements)
      .find((item) => item.kind === 'text' && item.role === 'numbering-marker');
    expect(placement?.bounds).toMatchObject({ xPt: 43, widthPt: 60 });
  });

  it('retains a text numbering marker when an auto-width frame has no body runs', () => {
    const framed = {
      ...paragraph(frame()),
      indentFirst: 20,
      runs: [],
      numbering: {
        text: '123456789012',
        jc: 'right',
        suff: 'nothing',
      },
    } as unknown as DocParagraph;
    const group = collectBodyFrameGroups([framed as unknown as BodyElement]).get(framed)!;
    const acquired = acquireRetainedFrameGroup(group, frameOptions(group, [framed]));

    expect(acquired.box.bounds.widthPt).toBe(60);
    const placement = acquired.members[0]?.fragment.lines
      .flatMap((line) => line.placements)
      .find((item) => item.kind === 'text' && item.role === 'numbering-marker');
    expect(placement?.bounds).toMatchObject({ xPt: -30, widthPt: 60 });
  });

  it('reflows a parser-owned frame anchor before retaining its exclusion', () => {
    const framed = framedParagraphWithTopAndBottomAnchor();
    const group = collectBodyFrameGroups([framed as unknown as BodyElement]).get(framed)!;
    const acquired = acquireRetainedFrameGroup(group, frameOptions(group, [framed]));
    const retained = acquired.members[0]!.fragment;
    const lineTexts = retained.lines.map((line) => line.placements
      .filter((placement) => placement.kind === 'text')
      .map((placement) => placement.text)
      .join(''));
    const exclusion = retained.exclusions.find((candidate) =>
      candidate.anchorOccurrenceId?.endsWith('frame-top-and-bottom'));

    // The anchor between the two runs opens no break opportunity: the lines
    // break as the unbroken word does (WORD_ANCHOR_HOST_NO_LINE_METRIC).
    expect(lineTexts).toEqual(['AAAAAAAA', 'ABBBBBBB', 'B']);
    expect(retained.lines[0]!.bounds.yPt).toBe(retained.flowBounds.yPt + 20);
    expect(exclusion?.bounds.yPt).toBe(retained.flowBounds.yPt);
  });

  it('carries a retained host exclusion into the next member of the frame group', () => {
    const anchored = framedParagraphWithFollowingSquareAnchor();
    const following = {
      ...paragraph(frame({ w: 160 })),
      runs: [{
        type: 'text', text: 'B'.repeat(40), bold: false, italic: false,
        underline: false, strikethrough: false, fontSize: 10, color: null,
        fontFamily: 'serif', isLink: false, background: null,
        vertAlign: null, hyperlink: null,
      }],
    } as unknown as DocParagraph;
    const group = collectBodyFrameGroups([
      anchored as unknown as BodyElement,
      following as unknown as BodyElement,
    ]).get(anchored)!;
    const acquired = acquireRetainedFrameGroup(
      group,
      frameOptions(group, [anchored, following]),
    );
    const followingLines = acquired.members[1]!.fragment.lines.map((line) =>
      line.placements.filter((placement) => placement.kind === 'text')
        .map((placement) => placement.text).join(''));

    expect(acquired.members[0]!.fragment.exclusions).toHaveLength(1);
    expect(acquired.members[1]!.fragment.exclusions.map((candidate) =>
      candidate.anchorOccurrenceId)).toEqual([
      acquired.members[0]!.fragment.exclusions[0]!.anchorOccurrenceId,
    ]);
    expect(followingLines).toEqual([
      'B'.repeat(32),
      'B'.repeat(8),
    ]);
  });

  it('retains a picture bullet when an auto-width frame has no body runs', () => {
    const framed = {
      ...paragraph(frame()),
      indentFirst: 20,
      runs: [],
      numbering: {
        text: '',
        jc: 'right',
        suff: 'nothing',
        picBulletImagePath: 'bullet.png',
        picBulletWidthPt: 60,
        picBulletHeightPt: 10,
      },
    } as unknown as DocParagraph;
    const group = collectBodyFrameGroups([framed as unknown as BodyElement]).get(framed)!;
    const acquired = acquireRetainedFrameGroup(group, frameOptions(group, [framed]));

    expect(acquired.box.bounds.widthPt).toBe(60);
    const placement = acquired.members[0]?.fragment.lines
      .flatMap((line) => line.placements)
      .find((item) => item.kind === 'resource' && item.resourceKind === 'picture-bullet');
    expect(placement?.bounds).toMatchObject({ xPt: -30, widthPt: 60 });
  });

  it('skips the intrinsic probe for explicit frame width and performs one final acquisition', () => {
    const framed = paragraph(frame({ w: 50 }));
    const group = collectBodyFrameGroups([framed as unknown as BodyElement]).get(framed)!;
    let calls = 0;
    const baseService = textService('counting-text');
    const countingService: TextLayoutService = {
      ...baseService,
      shape(request) {
        calls += 1;
        return baseService.shape(request);
      },
    };
    const options = frameOptions(group, [framed], {
      acquisitionSession: {},
    });
    const acquired = acquireRetainedFrameGroup(group, {
      ...options,
      environment: {
        ...options.environment,
        layoutServices: { ...options.environment.layoutServices!, text: countingService },
      },
    });

    // One paragraph acquisition shapes the segment, retained placement, and
    // paint projection. Additional groups of three expose a second layout pass.
    expect(calls).toBe(3);
    expect(acquired.members[0]?.fragment.flowBounds.yPt).toBe(20);
  });

  it('scopes cached acquisitions to the owning session and complete service fingerprint', () => {
    const framed = paragraph(frame({ w: 50 }));
    const group = collectBodyFrameGroups([framed as unknown as BodyElement]).get(framed)!;
    const session = {};
    const initial = frameOptions(group, [framed], { acquisitionSession: session });
    const first = acquireRetainedFrameGroup(group, initial);
    expect(acquireRetainedFrameGroup(group, initial)).toBe(first);

    const changedService = textService('frame-test-text:changed');
    const changed = acquireRetainedFrameGroup(group, {
      ...initial,
      environment: {
        ...initial.environment,
        layoutServices: { ...initial.environment.layoutServices!, text: changedService },
      },
    });
    expect(changed).not.toBe(first);

    const changedRules = acquireRetainedFrameGroup(group, {
      ...initial,
      contexts: [{
        ...initial.contexts[0]!,
        kinsoku: {
          ...initial.contexts[0]!.kinsoku,
          lineStartForbidden: new Set([0x3001]),
        },
      }],
    });
    expect(changedRules).not.toBe(first);

    const anotherSession = acquireRetainedFrameGroup(group, {
      ...initial,
      acquisitionSession: {},
    });
    expect(anotherSession).not.toBe(first);

    const relocated = acquireRetainedFrameGroup(group, {
      ...initial,
      placementSignature: 'column:30,40,180,280',
      place: (contentWidthPt, contentHeightPt) => ({
        bounds: { xPt: 30, yPt: 40, widthPt: contentWidthPt, heightPt: contentHeightPt },
        exclusionBounds: { xPt: 30, yPt: 40, widthPt: contentWidthPt, heightPt: contentHeightPt },
      }),
    });
    expect(relocated).not.toBe(first);
    expect(relocated.members[0]?.fragment.flowBounds.xPt)
      .toBe(first.members[0]!.fragment.flowBounds.xPt + 20);
    expect(relocated.members[0]?.fragment.flowBounds.yPt)
      .toBe(first.members[0]!.fragment.flowBounds.yPt + 20);
  });
});

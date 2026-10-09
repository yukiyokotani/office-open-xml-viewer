// Shared stub for WORD_COMPRESSED_SPACE_LINE_FIT tests: the Word-control
// fixture's embedded-face hmtx advances, glyf ink bounds and OS/2
// xAvgCharWidth, routed through the production text service.
import { readFileSync } from 'node:fs';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import { createLayoutServices } from '../layout-runtime.js';
import type { ParagraphLayoutContext } from '../layout-context.js';
import { createFontResolver } from '../layout/font-service.js';
import { acquireParagraphLayout } from '../layout/paragraph.js';
import { createTextLayoutService, type ResolvedFontMetric } from '../layout/text.js';
import { buildSegments, layoutLines, type LayoutTextSeg, type LineLayoutEnvironment } from '../line-layout.js';
import { paragraphAcquisitionInput } from '../parser-model.js';
import type { DocParagraph, DocRun, DocxDocumentModel } from '../types.js';

export interface Variant {
  readonly stage: string;
  readonly document: string;
  readonly variant: string;
  readonly compatibilityMode: number;
  readonly characterSpacingControl: string | null;
  readonly enableOpenTypeFeatures: boolean;
  readonly ascii: string;
  readonly eastAsia: string;
  readonly sizePt: number;
  readonly bold: boolean;
  readonly text: string;
  readonly justification: 'left' | 'both' | 'distribute';
  readonly sourceRuns: string;
  readonly kern: string;
  readonly autoSpaceDE: boolean;
  readonly autoSpaceDN: boolean;
  readonly borderEighths: number;
  readonly outOfScope: string | null;
  readonly widthsTwips: readonly number[];
  readonly wordWraps: string;
}
export interface Face {
  readonly unitsPerEm: number;
  readonly xAvgCharWidth: number;
  readonly advances: Readonly<Record<string, number>>;
  readonly ink: Readonly<Record<string, readonly number[]>>;
}
const controls = JSON.parse(readFileSync(new URL('../word-space-fit-controls.json', import.meta.url), 'utf8')) as {
  readonly fonts: Readonly<Record<string, Face>>;
  readonly variants: readonly Variant[];
};
export const FONTS = controls.fonts;
export const VARIANTS = controls.variants;
export const BAND_DEFICIT_PT: Readonly<Record<number, number>> = { 0: 0, 4: 1 };

function faceOf(familyList: string, weight: number): Face {
  const family = /"([^"]+)"/.exec(familyList)?.[1] ?? familyList.split(',')[0]!.trim();
  const face = FONTS[`${family}|${weight}`];
  if (!face) throw new Error(`no face ${family} ${weight}`);
  return face;
}

export function advancePt(face: Face, text: string, sizePt: number): number {
  let units = 0;
  for (const character of text) {
    const value = face.advances[character];
    if (value === undefined) throw new Error(`missing advance for ${JSON.stringify(character)}`);
    units += value;
  }
  return (units * sizePt) / face.unitsPerEm;
}

let measurements = 0;

/** Glyph measurements (Canvas and text-service) since the last reset. */
export function stubMeasurements(reset = false): number {
  const value = measurements;
  if (reset) measurements = 0;
  return value;
}

let canvasFont = '';
export const canvas = {
  get font() { return canvasFont; },
  set font(value: string) { canvasFont = value; },
  letterSpacing: '0px',
  fontKerning: 'auto',
  measureText(text: string) {
    measurements += 1;
    const px = Number(/([\d.]+)px/.exec(canvasFont)?.[1] ?? 10);
    const weight = /\bbold\b|\b700\b/.test(canvasFont) ? 700 : 400;
    const family = canvasFont.slice(canvasFont.indexOf('px') + 2).trim();
    const width = advancePt(faceOf(family, weight), text, px);
    return {
      width, actualBoundingBoxAscent: px * 0.8, actualBoundingBoxDescent: px * 0.2,
      fontBoundingBoxAscent: px * 0.88, fontBoundingBoxDescent: px * 0.12,
    } as TextMetrics;
  },
} as unknown as CanvasRenderingContext2D;

export function services(families: readonly string[]) {
  const empty = { default: null, first: null, even: null };
  const base = createLayoutServices({
    section: {
      pageWidth: 612, pageHeight: 792, marginTop: 72, marginRight: 72,
      marginBottom: 72, marginLeft: 72, headerDistance: 36, footerDistance: 36,
      titlePage: false, evenAndOddHeaders: false,
    },
    body: [], headers: empty, footers: empty,
  } as unknown as DocxDocumentModel, { measureContext: canvas });
  const routes = families.flatMap((family) => [400, 700].map((weight) => ({
    requestedFamily: family, resolvedFamily: family, source: 'local' as const,
    resourceIdentity: `office-local:local("${family}")`, weight,
  })));
  const fontMetrics: Record<string, ResolvedFontMetric> = {};
  for (const [key, face] of Object.entries(FONTS)) {
    const [family, weight] = key.split('|') as [string, string];
    if (!families.includes(family)) continue;
    fontMetrics[key] = {
      family, requestedFamily: family, weight: Number(weight),
      sourceIdentity: `office-local:local("${family}")`,
      averageCharWidthRatio: face.xAvgCharWidth / face.unitsPerEm,
      unicodeRanges: [...new Set(Object.keys(face.advances))]
        .map((character) => character.codePointAt(0)!)
        .map((code) => [code, code] as const),
    };
  }
  const text = createTextLayoutService({
    fonts: createFontResolver(routes),
    measurer: {
      fingerprint: 'word-space-fit-controls',
      measure: (request) => {
        measurements += 1;
        const face = faceOf(request.fontRoute.familyList, request.weight);
        const characters = [...request.text];
        const visible = characters.filter((character) => character !== ' ');
        const advance = advancePt(face, request.text, request.fontSizePt);
        const scalePt = request.fontSizePt / face.unitsPerEm;
        // Glyph outline bounds of the embedded face (glyf), for ink-based
        // punctuation compression; spaces carry no ink.
        let last = characters.length - 1;
        while (last >= 0 && characters[last] === ' ') last -= 1;
        const inkBounds = visible.length === 0 ? undefined : {
          xMinPt: face.ink[characters.find((character) => character !== ' ')!]![0] * scalePt,
          xMaxPt: advancePt(face, characters.slice(0, last).join(''), request.fontSizePt)
            + face.ink[characters[last]!]![1] * scalePt,
          ascentPt: request.fontSizePt * 0.88, descentPt: request.fontSizePt * 0.12,
        };
        return {
          advancePt: advance, ascentPt: request.fontSizePt * 0.88,
          descentPt: request.fontSizePt * 0.12,
          ...(inkBounds ? { inkBounds, horizontalInkBoundsAreTight: true } : {}),
        };
      },
    },
    fontMetrics,
  });
  return { ...base, text };
}


const SERVICES = new Map<string, ReturnType<typeof services>>();
function cachedServices(families: readonly string[]) {
  const key = families.join('\0');
  let value = SERVICES.get(key);
  if (!value) SERVICES.set(key, value = services(families));
  return value;
}

export interface StubRun {
  readonly text: string;
  readonly ascii: string;
  readonly eastAsia: string;
  readonly sizePt: number;
  readonly bold: boolean;
  readonly kerning?: number;
}

export interface StubParagraph {
  readonly runs: readonly StubRun[];
  readonly environment: Partial<LineLayoutEnvironment>;
  readonly bandPt: number;
  readonly justification: 'left' | 'both' | 'distribute' | 'center';
  /** Build uncached text services (for measurement counting). */
  readonly freshServices?: boolean;
}

function stubRuns(paragraph: StubParagraph): DocRun[] {
  return paragraph.runs.map((run) => ({
    type: 'text', text: run.text, fontFamily: run.ascii, fontFamilyHighAnsi: run.ascii,
    fontFamilyEastAsia: run.eastAsia, fontSize: run.sizePt, bold: run.bold,
    italic: false, underline: false, strikethrough: false, kerning: run.kerning,
    lang: 'en-US', langEastAsia: 'ja-JP',
  })) as unknown as DocRun[];
}

/** Production measurement and retained left-aligned paragraph layout over the
 * fixture faces, for checks on placed cluster and source-owner geometry. */
export function retainStubParagraph(paragraph: StubParagraph) {
  const families = [...new Set(paragraph.runs.flatMap((run) => [run.ascii, run.eastAsia]))];
  const docParagraph = {
    alignment: paragraph.justification, indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: stubRuns(paragraph),
  } as unknown as DocParagraph;
  const context: ParagraphLayoutContext = {
    lineGrid: { active: false, pitchPt: null },
    characterGrid: { active: false, kind: null, pitchPt: null, deltaPt: 0 },
    rightIndentGrid: { pitchPt: null, paragraphAllowsAdjustment: true },
    physicalIndentLeftPt: 0, physicalIndentRightPt: 0, firstIndentPt: 0,
    lineSpacing: null, spaceBeforePt: 0, spaceAfterPt: 0,
    baseRtl: false, isJustified: false, stretchLastLine: false,
    tabStops: [], hasRuby: false, hasEastAsianText: true,
    kinsoku: DEFAULT_KINSOKU_RULES, defaultTabPt: 36,
  };
  const placement = {
    startYPt: 0, paragraphXPt: 0, availableWidthPt: paragraph.bandPt, maximumYPt: 1e7,
    suppressSpaceBefore: false,
  };
  const measurer = { context: canvas, fontFamilyClasses: {} };
  const environment = {
    pageIndex: 0, totalPages: 1, documentHasEastAsianText: true,
    pageWritingMode: 'horizontal-tb' as const, layoutServices: cachedServices(families),
    autoSpaceDE: false, autoSpaceDN: false, ...paragraph.environment,
  };
  const source = { story: 'body' as const, storyInstance: 'body', path: [0] };
  const input = paragraphAcquisitionInput(docParagraph, source);
  return acquireParagraphLayout(input, {
    id: 'word-space-fit', source,
    flowDomainId: 'body', ordinaryFlow: true, context, placement, measurer, environment,
    exclusions: [],
  });
}

/** Production buildSegments + layoutLines over the fixture faces; returns each
 * line's segment texts and measured widths. */
export function layoutStubParagraph(paragraph: StubParagraph) {
  const families = [...new Set(paragraph.runs.flatMap((run) => [run.ascii, run.eastAsia]))];
  const runs = stubRuns(paragraph);
  const segments = buildSegments(runs, {
    pageIndex: 0, totalPages: 1, layoutServices: paragraph.freshServices ? services(families) : cachedServices(families),
    // The Word controls author both automatic-spacing flags off; stub
    // paragraphs follow them unless a test sets the flags.
    autoSpaceDE: false, autoSpaceDN: false, ...paragraph.environment,
  });
  const justified = paragraph.justification === 'both' || paragraph.justification === 'distribute';
  return layoutLines(canvas, segments, paragraph.bandPt, 0, 1, [], undefined, {}, 0, undefined,
    undefined, 36, 0, false, justified, paragraph.justification === 'distribute', undefined,
    'bounded', undefined, false).map((line) => line.segments.map((segment) => ({
    text: (segment as LayoutTextSeg).text ?? '',
    width: Math.round(segment.measuredWidth * 1e6) / 1e6,
    compression: (segment as LayoutTextSeg).latinSpaceCompressionPx ?? 0,
  })));
}

import type { SectionLayoutContext } from '../layout-context.js';
import type { TextLayoutService } from './text.js';
import {
  ERROR_PLACEHOLDER_DETAIL_LINES,
  errorPlaceholderDetail,
  errorPlaceholderTitle,
  wrapErrorPlaceholderText,
  type ResolvedErrorPlaceholder,
} from '@silurus/ooxml-core/internal/part-failures';

export type { WireErrorPlaceholder } from '@silurus/ooxml-core/internal/part-failures';
import { createPageLayers } from './page-graph.js';
import type {
  DocumentLayout,
  DrawingLayout,
  DrawingPaintCommand,
  LayoutRect,
  SourceRef,
} from './types.js';

export interface ErrorPageSize {
  readonly widthPt: number;
  readonly heightPt: number;
}

const ERROR_SOURCE: SourceRef = Object.freeze({
  story: 'body',
  storyInstance: 'parse-error',
  path: Object.freeze([]),
});

function placeholderLines(
  text: string,
  maxWidthPt: number,
  fontSizePt: number,
  service: TextLayoutService,
  maxLines: number,
  weight?: number,
): string[] {
  return wrapErrorPlaceholderText(text, maxWidthPt, maxLines, (value) => service.shape({
    text: value,
    fontSizePt,
    ...(weight === undefined ? {} : { weight }),
    fonts: {},
    genericFamily: 'sans-serif',
  }).advancePt);
}

export const PARSE_ERROR_PAGE_TITLE = 'This document could not be displayed';

/**
 * Lay out the synthetic page shown for a document whose main part failed to
 * parse. `placeholder` replaces the library's title/detail text; `null` (the
 * default) is the built-in text. The text is shaped here, at acquisition, so
 * paint stays measurement-free.
 */
export function layoutParseErrorPage(
  message: string,
  size: ErrorPageSize,
  text: TextLayoutService,
  placeholder: ResolvedErrorPlaceholder | null = null,
): DocumentLayout {
  if (!(size.widthPt > 0 && size.heightPt > 0)) throw new RangeError('Error page size must be positive');
  const padPt = Math.max(18, Math.min(size.widthPt, size.heightPt) * 0.06);
  const frame: LayoutRect = {
    xPt: padPt,
    yPt: padPt,
    widthPt: size.widthPt - padPt * 2,
    heightPt: size.heightPt - padPt * 2,
  };
  const detailSizePt = Math.max(8, Math.min(size.widthPt, size.heightPt) * 0.02);
  const fontRoute = text.resolve({ fonts: {}, slot: 'ascii', genericFamily: 'sans-serif' }).route;
  const detail = errorPlaceholderDetail(placeholder, message);
  const detailLines = detail === null ? [] : placeholderLines(detail, size.widthPt - padPt * 4, detailSizePt, text, ERROR_PLACEHOLDER_DETAIL_LINES);
  const detailLineHeightPt = detailSizePt * 1.4;
  const title = placeholderLines(
    errorPlaceholderTitle(placeholder, PARSE_ERROR_PAGE_TITLE),
    size.widthPt - padPt * 4,
    13,
    text,
    1,
    600,
  )[0] ?? '';
  const commands: DrawingPaintCommand[] = [
    { kind: 'fill-rect', rect: { xPt: 0, yPt: 0, widthPt: size.widthPt, heightPt: size.heightPt }, fill: '#ffffff' },
    { kind: 'stroke-rect', rect: frame, stroke: '#c8ccd2', lineWidthPt: 1, dashPt: [6, 5] },
    {
      kind: 'text', rect: { xPt: 0, yPt: size.heightPt * 0.27, widthPt: size.widthPt, heightPt: 36 },
      text: '⚠', fill: '#b23b3b', fontRoute, fontSizePt: 28,
      fontWeight: 400, fontStyle: 'normal', align: 'center', baseline: 'middle',
    },
    {
      kind: 'text', rect: { xPt: padPt * 2, yPt: size.heightPt * 0.40, widthPt: size.widthPt - padPt * 4, heightPt: 24 },
      text: title, fill: '#333333', fontRoute, fontSizePt: 13,
      fontWeight: 600, fontStyle: 'normal', align: 'center', baseline: 'middle',
    },
    ...detailLines.map((line, index): DrawingPaintCommand => ({
      kind: 'text',
      rect: {
        xPt: padPt * 2,
        yPt: size.heightPt * 0.50 + detailLineHeightPt * index,
        widthPt: size.widthPt - padPt * 4,
        heightPt: detailLineHeightPt,
      },
      text: line,
      fill: '#666666',
      fontRoute,
      fontSizePt: detailSizePt,
      fontWeight: 400,
      fontStyle: 'normal',
      align: 'center',
      baseline: 'middle',
    })),
  ];
  const node: DrawingLayout = {
    kind: 'drawing',
    id: 'parse-error-page',
    source: ERROR_SOURCE,
    flowDomainId: 'parse-error',
    flowBounds: frame,
    inkBounds: frame,
    advancePt: frame.heightPt,
    ordinaryFlow: false,
    commands,
  };
  const section: SectionLayoutContext = {
    geometry: {
      pageWidth: size.widthPt,
      pageHeight: size.heightPt,
      marginTop: padPt,
      marginRight: padPt,
      marginBottom: padPt,
      marginLeft: padPt,
      headerDistance: 0,
      footerDistance: 0,
    },
    columns: [{ xPt: padPt, wPt: frame.widthPt }],
    columnSeparator: false,
    grid: { kind: 'none', linePitchPt: null, charSpacePt: null },
    textDirection: 'lrTb',
    verticalAlignment: 'top',
  };
  return {
    pages: [{
      pageIndex: 0,
      geometry: {
        xPt: 0,
        yPt: 0,
        widthPt: size.widthPt,
        heightPt: size.heightPt,
        contentTopPt: padPt,
        contentBottomPt: size.heightPt - padPt,
      },
      flowDomains: [{
        id: 'parse-error', kind: 'body', logicalBounds: frame, physicalBounds: frame,
      }],
      section,
      sectionOccurrenceId: 'parse-error-section',
      pageBorder: null,
      parityBlank: false,
      bookmarkStarts: [],
      pageNumber: {
        displayNumber: 1,
        format: 'decimal',
        sectionOccurrenceId: 'parse-error-section',
      },
      columnSeparators: [],
      sectionRegions: [{
        id: 'parse-error-region',
        sectionOccurrenceId: 'parse-error-section',
        coordinateSpace: {
          writingMode: 'horizontal-tb',
          logicalToPhysical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
          physicalToLogical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
        },
        blockStartPt: padPt,
        blockEndPt: size.heightPt - padPt,
        columnFlowDirection: 'ltr',
        columnIndexes: [0],
        flowDomainIds: ['parse-error'],
        section,
      }],
      layers: createPageLayers([{ layer: 'body', node }]),
      readingOrder: [node.id],
    }],
    diagnostics: [{
      code: 'UNSUPPORTED_FEATURE',
      severity: 'error',
      source: ERROR_SOURCE,
      message,
    }],
  };
}

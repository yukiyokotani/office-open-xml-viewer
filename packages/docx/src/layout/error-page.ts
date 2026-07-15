import type { SectionLayoutContext } from '../layout-context.js';
import type { TextLayoutService } from './text.js';
import { graphemeClusterOffsets } from '@silurus/ooxml-core';
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

function wrapWords(
  text: string,
  maxWidthPt: number,
  fontSizePt: number,
  service: TextLayoutService,
  maxLines: number,
): string[] {
  const words = text.trim().split(/\s+/).filter(Boolean);
  const lines: string[] = [];
  let line = '';
  const widthOf = (value: string): number => service.shape({
    text: value,
    fontSizePt,
    fonts: {},
    genericFamily: 'sans-serif',
  }).advancePt;
  const ellipsize = (value: string): string => {
    const ellipsis = '…';
    const offsets = [0, ...graphemeClusterOffsets(value), value.length]
      .filter((offset, index, values) => index === 0 || offset !== values[index - 1]);
    let end = offsets.length - 1;
    while (end > 0 && widthOf(`${value.slice(0, offsets[end])}${ellipsis}`) > maxWidthPt) end -= 1;
    return `${value.slice(0, offsets[end] ?? 0)}${ellipsis}`;
  };
  const truncateLast = (): void => {
    if (lines.length === 0) lines.push(ellipsize(''));
    else lines[lines.length - 1] = ellipsize(lines[lines.length - 1]);
  };
  const pushOverlong = (word: string): boolean => {
    const offsets = [0, ...graphemeClusterOffsets(word), word.length]
      .filter((offset, index, values) => index === 0 || offset !== values[index - 1]);
    let start = 0;
    while (start < offsets.length - 1) {
      if (lines.length >= maxLines) {
        truncateLast();
        return true;
      }
      let end = start + 1;
      while (end < offsets.length && widthOf(word.slice(offsets[start], offsets[end])) <= maxWidthPt) end += 1;
      const fittingEnd = Math.max(start + 1, end - 1);
      const chunk = word.slice(offsets[start], offsets[fittingEnd]);
      if (lines.length >= maxLines - 1 && fittingEnd < offsets.length - 1) {
        lines.push(ellipsize(chunk));
        return true;
      }
      lines.push(chunk);
      start = fittingEnd;
    }
    return false;
  };
  for (let wordIndex = 0; wordIndex < words.length; wordIndex += 1) {
    const word = words[wordIndex];
    const candidate = line ? `${line} ${word}` : word;
    const candidateWidth = widthOf(candidate);
    if (line && candidateWidth > maxWidthPt) {
      lines.push(line);
      line = '';
      if (lines.length >= maxLines) {
        truncateLast();
        break;
      }
      if (widthOf(word) > maxWidthPt) {
        if (pushOverlong(word)) break;
      }
      else line = word;
    } else if (!line && candidateWidth > maxWidthPt) {
      if (pushOverlong(word)) break;
    } else {
      line = candidate;
    }
    if (wordIndex < words.length - 1 && lines.length >= maxLines && !line) {
      truncateLast();
      break;
    }
  }
  if (line && lines.length < maxLines) lines.push(line);
  return lines;
}

export function layoutParseErrorPage(
  message: string,
  size: ErrorPageSize,
  text: TextLayoutService,
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
  const detailLines = wrapWords(message, size.widthPt - padPt * 4, detailSizePt, text, 4);
  const detailLineHeightPt = detailSizePt * 1.4;
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
      text: 'This document could not be displayed', fill: '#333333', fontRoute, fontSizePt: 13,
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
      flowDomains: [{ id: 'parse-error', kind: 'body', bounds: frame }],
      section,
      sectionOccurrenceId: 'parse-error:section',
      parityBlank: false,
      bookmarkStarts: [],
      pageNumber: {
        displayNumber: 1,
        format: 'decimal',
        sectionOccurrenceId: 'parse-error:section',
      },
      sectionRegions: [{
        id: 'parse-error:region',
        sectionOccurrenceId: 'parse-error:section',
        coordinateSpace: {
          writingMode: 'horizontal-tb',
          logicalToPhysical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
        },
        blockStartPt: frame.yPt,
        blockEndPt: frame.yPt + frame.heightPt,
        flowDomainIds: ['parse-error'],
        section,
      }],
      layers: {
        paintOrder: [{
          layer: 'body', nodeId: node.id, coordinateSpace: 'logical-body-points',
          logicalBlock: { blockStartPt: frame.yPt, blockExtentPt: frame.heightPt },
        }],
        background: [], behindText: [], header: [], body: [node], notes: [], front: [], footer: [],
      },
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

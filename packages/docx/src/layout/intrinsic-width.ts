import { graphemeClusterOffsets } from '@silurus/ooxml-core';
import type { DocTableCell } from '../types.js';
import type { ParagraphLayoutContext } from '../layout-context.js';
import {
  buildFont,
  buildSegments,
  hasCJKBreakOpportunity,
  layoutLines,
  segAdvanceWidth,
  snapToCharsAllocatedWidthPx,
  snapToCharsClass,
  slicedPunctuationCompressions,
  splitTextForLayout,
  type LayoutLine,
  type LayoutSeg,
  type LayoutTextSeg,
  type DocGridCtx,
} from '../line-layout.js';
import type {
  ParagraphMeasurementEnvironment,
  TextMeasurer,
} from '../paragraph-measure.js';
import { paragraphCharacterGrid } from '../paragraph-measure.js';
import { calcEffectiveFontPx } from './text.js';
import { wordSnapToCharsEastAsianCellCount } from './line-compatibility.js';
import type { ParagraphLayoutSource, TextFontSlots } from './text.js';
import type { TableLayoutSource } from './table-source-acquisition.js';
import type { DeepReadonly } from './types.js';
import { stableFingerprint } from './fingerprint.js';
import {
  numberingMarkerLogicalInterval,
  type NumberingMarkerGeometry,
} from './numbering-marker.js';
import { wordAutofitEmptyParagraphHasNoIntrinsicContent } from './table-compatibility.js';

export interface ParagraphIntrinsicWidths {
  readonly minWidthPt: number;
  readonly maxWidthPt: number;
}

export interface TableCellIntrinsicWidths {
  readonly minWidthPt: number;
  readonly maxWidthPt: number;
}

export interface TableCellIntrinsicWidthDependencies {
  paragraph(paragraph: ParagraphLayoutSource): TableCellIntrinsicWidths;
  nestedTable(table: TableLayoutSource): TableCellIntrinsicWidths;
}

export interface ParagraphIntrinsicWidthOptions {
  /** Word table AutoFit keeps a retained whitespace-only paragraph run as
   * content even though ordinary line layout trims paragraph-final spaces. */
  readonly preserveWhitespaceOnlyContent?: boolean;
}

/** Fold public cell content into one intrinsic interval. OOXML width/style
 * precedence is deliberately absent: parser/model projection and the column
 * solver own those separate responsibilities. */
export function measureTableCellIntrinsicWidths(
  cell: DeepReadonly<DocTableCell>,
  margins: Readonly<{ left: number; right: number }>,
  dependencies: TableCellIntrinsicWidthDependencies,
): TableCellIntrinsicWidths {
  let minContentWidthPt = 0;
  let maxContentWidthPt = 0;
  for (const element of cell.content) {
    // ECMA-376 §17.18.87 defines AutoFit minima from cell contents. The
    // registered Word observation refines the otherwise-unspecified empty-mark
    // boundary: indentation positions no content when the cell paragraph has
    // neither runs nor numbering. Cell margins remain applied below.
    const emptyParagraph = element.type === 'paragraph'
      && wordAutofitEmptyParagraphHasNoIntrinsicContent(element);
    const intrinsic = element.type === 'paragraph'
      ? emptyParagraph ? { minWidthPt: 0, maxWidthPt: 0 } : dependencies.paragraph(element)
      : dependencies.nestedTable(element);
    minContentWidthPt = Math.max(minContentWidthPt, intrinsic.minWidthPt);
    maxContentWidthPt = Math.max(maxContentWidthPt, intrinsic.maxWidthPt);
  }
  const horizontalMarginsPt = Math.max(0, margins.left) + Math.max(0, margins.right);
  return {
    minWidthPt: minContentWidthPt + horizontalMarginsPt,
    maxWidthPt: Math.max(minContentWidthPt, maxContentWidthPt) + horizontalMarginsPt,
  };
}

function compatibleTextKey(segment: LayoutTextSeg): string {
  const request = segment.textShapeRequest;
  const slots = (value: TextFontSlots | undefined) => value
    ? [
        value.ascii ?? null,
        value.highAnsi ?? null,
        value.eastAsia ?? null,
        value.complexScript ?? null,
      ]
    : null;
  return stableFingerprint('paragraph-intrinsic-text', [
    segment.textLayoutService?.fingerprint ?? null,
    request ? [
      slots(request.fonts),
      slots(request.themeFonts),
      request.themeFontPresence ? [
        request.themeFontPresence.ascii ?? false,
        request.themeFontPresence.highAnsi ?? false,
        request.themeFontPresence.eastAsia ?? false,
        request.themeFontPresence.complexScript ?? false,
      ] : null,
      request.fontHint ?? null,
      request.fontSizePt,
      request.weight ?? null,
      request.style ?? null,
      request.complexScript ?? false,
      request.eastAsiaLanguage ?? null,
      request.eastAsiaFontCharset ?? null,
      request.genericFamily ?? null,
      request.letterSpacingPt ?? null,
      request.kerning ?? null,
    ] : null,
    segment.bold,
    segment.italic,
    calcEffectiveFontPx(segment, 1),
    segment.fontFamily,
    segment.fontRoute ?? null,
    segment.charScale ?? 1,
    segment.charSpacing ?? 0,
    segment.fitTextPerGapPx ?? null,
    segment.fitTextTrailingPadPx ?? null,
    segment.fitTextRegionIndex ?? null,
    segment.snapToCharacterGrid !== false,
    segment.widthBalanceGridDeltaFactor ?? null,
    segment.widthBalanceSpaceSequence ?? false,
    segment.widthBalanceSpaceAdjustmentPt ?? null,
    // The snap-to-character-grid allocator consumes contiguous script blocks.
    // A shaping-compatible Latin→East-Asian seam is therefore still a semantic
    // grid boundary and must survive this intrinsic-only merge.
    segment.script,
    segment.tateChuYoko ?? false,
    // A tate-chu-yoko run is one authored one-em cell (§17.3.2.10). Two
    // adjacent runs with identical fonts remain two cells, so their source-run
    // boundary is semantic rather than a shaping-only seam.
    segment.tateChuYoko ? (segment.sourceRunIndex ?? null) : null,
    // Ruby belongs to its authored base run. Extending that base across a run
    // seam would change the annotation's ownership during the intrinsic probe.
    segment.ruby ? [
      segment.sourceRunIndex ?? null,
      segment.ruby.text,
      segment.ruby.fontSizePt,
      segment.ruby.hpsRaisePt ?? null,
    ] : null,
    segment.verticalRun ?? false,
  ]);
}

/** Run boundaries with identical effective metrics are not shaping boundaries.
 * Merge only for the intrinsic probe; retained source/run ownership stays intact. */
function mergeCompatibleTextSegments(segments: readonly LayoutSeg[]): LayoutSeg[] {
  const merged: LayoutSeg[] = [];
  for (const segment of segments) {
    const previous = merged.at(-1);
    if (
      previous
      && 'text' in previous
      && 'text' in segment
      && compatibleTextKey(previous) === compatibleTextKey(segment)
    ) {
      const previousTextLength = previous.text.length;
      const text = previous.text + segment.text;
      const punctuationCompressions = [
        ...(previous.punctuationCompressions ?? []),
        ...(segment.punctuationCompressions ?? []).map((compression) => ({
          end: previousTextLength + compression.end,
          adjustmentPt: compression.adjustmentPt,
        })),
      ];
      merged[merged.length - 1] = {
        ...previous,
        text,
        punctuationCompressions: punctuationCompressions.length > 0
          ? punctuationCompressions
          : undefined,
        textShapeRequest: previous.textShapeRequest
          ? { ...previous.textShapeRequest, text }
          : undefined,
      };
      continue;
    }
    merged.push({ ...segment });
  }
  return merged;
}

function measureTextRange(
  pieces: readonly Readonly<{ segment: LayoutTextSeg; start: number; end: number }>[],
  joinedText: string,
  start: number,
  end: number,
  measurer: TextMeasurer,
  characterGrid: DocGridCtx | undefined,
): number {
  let widthPt = 0;
  let pendingSnapBlock: Readonly<{
    kind: 'latin' | 'complexScript';
    naturalWidthPt: number;
  }> | null = null;
  const pitchPt = characterGrid?.type === 'snapToChars'
    && characterGrid.characterPitchPt != null
    && characterGrid.characterPitchPt > 0
      ? characterGrid.characterPitchPt
      : null;
  const flushSnapBlock = (): void => {
    if (!pendingSnapBlock || pitchPt == null) return;
    widthPt += snapToCharsAllocatedWidthPx(
      pendingSnapBlock.naturalWidthPt,
      pendingSnapBlock.kind,
      pitchPt,
    );
    pendingSnapBlock = null;
  };
  for (const piece of pieces) {
    const overlapStart = Math.max(start, piece.start);
    const overlapEnd = Math.min(end, piece.end);
    if (overlapStart >= overlapEnd) continue;
    const text = joinedText.slice(overlapStart, overlapEnd);
    const localStart = overlapStart - piece.start;
    const localEnd = overlapEnd - piece.start;
    const candidate = {
      ...piece.segment,
      text,
      punctuationCompressions: slicedPunctuationCompressions(
        piece.segment,
        localStart,
        localEnd,
      ),
    };
    const measureCandidate = (measured: LayoutTextSeg): number => {
      if (measured.textLayoutService && measured.textShapeRequest) {
        const shaped = measured.textLayoutService.shape({
          ...measured.textShapeRequest,
          text: measured.text,
          fontSizePt: calcEffectiveFontPx(measured, 1),
          measure: true,
          clusterGeometry: false,
        });
        return segAdvanceWidth(measured, shaped.advancePt, characterGrid, 1);
      }
      measurer.context.font = buildFont(
        measured.bold,
        measured.italic,
        calcEffectiveFontPx(measured, 1),
        measured.fontFamily,
        measurer.fontFamilyClasses as Record<string, string>,
        measured.fontRoute,
      );
      return segAdvanceWidth(
        measured,
        measurer.context.measureText(measured.text).width,
        characterGrid,
        1,
      );
    };
    const kind = snapToCharsClass(candidate, characterGrid);
    if (kind === 'eastAsia' && pitchPt != null) {
      flushSnapBlock();
      const shapedClusters = candidate.textLayoutService && candidate.textShapeRequest
        ? candidate.textLayoutService.shape({
            ...candidate.textShapeRequest,
            text,
            fontSizePt: calcEffectiveFontPx(candidate, 1),
            measure: true,
            clusterGeometry: true,
          }).clusters
        : undefined;
      const boundaries = shapedClusters?.length
        ? null
        : [...new Set([
            0,
            ...graphemeClusterOffsets(text),
            text.length,
          ])].sort((a, b) => a - b);
      const ranges = shapedClusters?.map((cluster) => ({
        start: cluster.range.start,
        end: cluster.range.end,
        naturalWidthPt: segAdvanceWidth(
          {
            ...candidate,
            text: text.slice(cluster.range.start, cluster.range.end),
            punctuationCompressions: slicedPunctuationCompressions(
              candidate,
              cluster.range.start,
              cluster.range.end,
            ),
          },
          cluster.advancePt,
          characterGrid,
          1,
        ),
      })) ?? boundaries!.slice(0, -1).map((clusterStart, index) => {
        const clusterEnd = boundaries![index + 1]!;
        const cluster = {
          ...candidate,
          text: text.slice(clusterStart, clusterEnd),
          punctuationCompressions: slicedPunctuationCompressions(
            candidate,
            clusterStart,
            clusterEnd,
          ),
        };
        return {
          start: clusterStart,
          end: clusterEnd,
          naturalWidthPt: measureCandidate(cluster),
        };
      });
      let cells = 0;
      for (const range of ranges) {
        if (range.end <= range.start) continue;
        cells += wordSnapToCharsEastAsianCellCount(
          range.naturalWidthPt,
          pitchPt,
        );
      }
      widthPt += snapToCharsAllocatedWidthPx(
        ranges.reduce((sum, range) => sum + range.naturalWidthPt, 0),
        kind,
        pitchPt,
        Math.max(1, cells),
      );
    } else {
      const naturalWidthPt = measureCandidate(candidate);
      if ((kind === 'latin' || kind === 'complexScript') && pitchPt != null) {
        const previousBlock = pendingSnapBlock as Readonly<{
          kind: 'latin' | 'complexScript';
          naturalWidthPt: number;
        }> | null;
        if (previousBlock?.kind === kind) {
          pendingSnapBlock = {
            kind,
            naturalWidthPt: previousBlock.naturalWidthPt + naturalWidthPt,
          };
        } else {
          flushSnapBlock();
          pendingSnapBlock = { kind, naturalWidthPt };
        }
      } else {
        flushSnapBlock();
        widthPt += naturalWidthPt;
      }
    }
  }
  flushSnapBlock();
  return widthPt;
}

function minimumTextAtomWidthPt(
  segments: readonly LayoutSeg[],
  context: ParagraphLayoutContext,
  measurer: TextMeasurer,
): number {
  const characterGrid = paragraphCharacterGrid(context);
  let maximumPt = 0;
  for (let segmentIndex = 0; segmentIndex < segments.length; segmentIndex += 1) {
    const segment = segments[segmentIndex];
    if (!('text' in segment) || segment.text.length === 0) continue;
    const pieces: Array<{ segment: LayoutTextSeg; start: number; end: number }> = [];
    let joinedText = '';
    const append = (piece: LayoutTextSeg): void => {
      const start = joinedText.length;
      joinedText += piece.text;
      pieces.push({ segment: piece, start, end: joinedText.length });
    };
    append(segment);
    while (segmentIndex + 1 < segments.length) {
      const following = segments[segmentIndex + 1];
      if (!('text' in following) || following.joinPrev !== true) break;
      append(following);
      segmentIndex += 1;
    }

    let tokenStart = 0;
    for (const token of splitTextForLayout(joinedText)) {
      const trimmed = token.replace(/\s+$/u, '');
      const trimmedStart = tokenStart;
      const trimmedEnd = tokenStart + trimmed.length;
      tokenStart += token.length;
      if (!trimmed) continue;
      if (!hasCJKBreakOpportunity(trimmed)) {
        maximumPt = Math.max(
          maximumPt,
          measureTextRange(
            pieces,
            joinedText,
            trimmedStart,
            trimmedEnd,
            measurer,
            characterGrid,
          ),
        );
        continue;
      }

      const boundaries = [0, ...graphemeClusterOffsets(trimmed), trimmed.length];
      const clusters: Array<{ text: string; start: number; end: number }> = [];
      for (let index = 1; index < boundaries.length; index += 1) {
        clusters.push({
          text: trimmed.slice(boundaries[index - 1], boundaries[index]),
          start: trimmedStart + boundaries[index - 1],
          end: trimmedStart + boundaries[index],
        });
      }
      const atoms: Array<{ text: string; start: number; end: number }> = [];
      let atom = clusters[0];
      for (let index = 1; index < clusters.length; index += 1) {
        const previous = [...atom.text].at(-1)?.codePointAt(0);
        const next = clusters[index].text.codePointAt(0);
        const breakAllowed = previous !== undefined
          && next !== undefined
          && !context.kinsoku.lineEndForbidden.has(previous)
          && !context.kinsoku.lineStartForbidden.has(next);
        if (breakAllowed) {
          atoms.push(atom);
          atom = clusters[index];
        } else {
          atom = {
            text: atom.text + clusters[index].text,
            start: atom.start,
            end: clusters[index].end,
          };
        }
      }
      if (atom) atoms.push(atom);
      for (const unbreakable of atoms) {
        maximumPt = Math.max(
          maximumPt,
          measureTextRange(
            pieces,
            joinedText,
            unbreakable.start,
            unbreakable.end,
            measurer,
            characterGrid,
          ),
        );
      }
    }
  }
  return maximumPt;
}

function logicalLineInterval(
  line: LayoutLine,
  lineIndex: number,
  context: ParagraphLayoutContext,
): Readonly<{ startPt: number; endPt: number }> {
  const leadingIndentPt = context.baseRtl
    ? context.physicalIndentRightPt
    : context.physicalIndentLeftPt;
  const startPt = leadingIndentPt
    + (lineIndex === 0 ? context.firstIndentPt : 0)
    + line.xOffset;
  const widthPt = line.segments.reduce((sum, segment) => sum + segment.measuredWidth, 0);
  return { startPt, endPt: startPt + widthPt };
}

export function measureParagraphIntrinsicWidths(
  paragraph: ParagraphLayoutSource,
  context: ParagraphLayoutContext,
  maximumWidthPt: number,
  measurer: TextMeasurer,
  environment: ParagraphMeasurementEnvironment,
  numbering?: NumberingMarkerGeometry,
  options: ParagraphIntrinsicWidthOptions = {},
): ParagraphIntrinsicWidths {
  if (!Number.isFinite(maximumWidthPt) || maximumWidthPt < 0) {
    throw new RangeError('maximumWidthPt must be finite and non-negative');
  }
  if (maximumWidthPt === 0) return { minWidthPt: 0, maxWidthPt: 0 };

  const segments = mergeCompatibleTextSegments(buildSegments(paragraph.runs, environment));
  const paragraphWidthPt = Math.max(
    1,
    maximumWidthPt - context.physicalIndentLeftPt - context.physicalIndentRightPt,
  );
  const lines = segments.length === 0 ? [] : layoutLines(
    measurer.context,
    segments,
    paragraphWidthPt,
    context.firstIndentPt,
    1,
    [...context.tabStops],
    undefined,
    measurer.fontFamilyClasses as Record<string, string>,
    context.physicalIndentLeftPt,
    context.kinsoku,
    paragraphCharacterGrid(context),
    context.defaultTabPt,
    paragraphWidthPt + context.physicalIndentRightPt,
    context.baseRtl,
    context.isJustified,
    context.stretchLastLine,
    undefined,
    'intrinsic',
    environment.verticalGlyphMeasurement,
    context.overflowPunct !== false,
  );
  const oppositeIndentPt = context.baseRtl
    ? context.physicalIndentLeftPt
    : context.physicalIndentRightPt;
  let minimumLeftPt = 0;
  let maximumRightPt = 0;
  lines.forEach((line, index) => {
    const interval = logicalLineInterval(line, index, context);
    minimumLeftPt = Math.min(minimumLeftPt, interval.startPt);
    maximumRightPt = Math.max(maximumRightPt, interval.endPt);
  });
  const markerInterval = numbering ? numberingMarkerLogicalInterval({
    leadingIndentPt: context.baseRtl
      ? context.physicalIndentRightPt
      : context.physicalIndentLeftPt,
    authoredFirstIndentPt: paragraph.indentFirst,
    markerShiftPt: numbering.markerShiftPt,
    markerWidthPt: numbering.markerWidthPt,
  }) : undefined;
  if (markerInterval) {
    minimumLeftPt = Math.min(minimumLeftPt, markerInterval.startPt);
    maximumRightPt = Math.max(maximumRightPt, markerInterval.endPt);
  }
  let minimumAtomPt = minimumTextAtomWidthPt(segments, context, measurer);
  for (const line of lines) {
    let penPt = 0;
    const lineWidthPt = line.segments.reduce((sum, segment) => sum + segment.measuredWidth, 0);
    for (const segment of line.segments) {
      penPt += segment.measuredWidth;
      if ('imagePath' in segment && !segment.anchor) {
        minimumAtomPt = Math.max(minimumAtomPt, segment.measuredWidth);
      } else if ('math' in segment) {
        minimumAtomPt = Math.max(minimumAtomPt, segment.measuredWidth);
      } else if ('isTab' in segment) {
        minimumAtomPt = Math.max(
          minimumAtomPt,
          segment.resolvedAlignment === 'left' ? penPt : lineWidthPt,
        );
      }
    }
  }
  const leadingIndentPt = context.baseRtl
    ? context.physicalIndentRightPt
    : context.physicalIndentLeftPt;
  const whitespaceSegments = options.preserveWhitespaceOnlyContent
    && segments.length > 0
    && segments.every((segment) => 'text' in segment && /^[\s\u00a0]+$/u.test(segment.text))
      ? segments as readonly LayoutTextSeg[]
      : null;
  const whitespaceOnlyWidthPt = whitespaceSegments
    ? (() => {
        let joinedText = '';
        const pieces = whitespaceSegments.map((segment) => {
          const start = joinedText.length;
          joinedText += segment.text;
          return { segment, start, end: joinedText.length };
        });
        return measureTextRange(
          pieces,
          joinedText,
          0,
          joinedText.length,
          measurer,
          paragraphCharacterGrid(context),
        );
      })()
    : 0;
  if (whitespaceOnlyWidthPt > 0) {
    const whitespaceStartPt = leadingIndentPt + context.firstIndentPt;
    minimumLeftPt = Math.min(minimumLeftPt, whitespaceStartPt);
    maximumRightPt = Math.max(maximumRightPt, whitespaceStartPt + whitespaceOnlyWidthPt);
  }
  const maxWidthPt = Math.min(
    maximumWidthPt,
    Math.max(0, maximumRightPt - minimumLeftPt + oppositeIndentPt),
  );
  const continuationStartPt = leadingIndentPt;
  let minLeftPt = Math.min(0, continuationStartPt);
  minimumAtomPt = Math.max(minimumAtomPt, whitespaceOnlyWidthPt);
  let minRightPt = Math.max(0, continuationStartPt + minimumAtomPt);
  const firstStartPt = leadingIndentPt + context.firstIndentPt;
  minLeftPt = Math.min(minLeftPt, firstStartPt);
  minRightPt = Math.max(minRightPt, firstStartPt + minimumAtomPt);
  if (markerInterval) {
    minLeftPt = Math.min(minLeftPt, markerInterval.startPt);
    minRightPt = Math.max(minRightPt, markerInterval.endPt);
  }
  const minWidthPt = Math.min(
    maximumWidthPt,
    Math.max(0, minRightPt - minLeftPt + oppositeIndentPt),
  );
  return { minWidthPt, maxWidthPt };
}

import type { BodyElement, NumberingInfo, ShapeRun, TabStop } from '../types.js';
import type {
  DeepReadonly,
  NumberingMarkerShapeInput,
  ParsedUnsupportedTextBoxBlock,
  SourceRef,
} from './types.js';
import { snapshotPlainData } from './plain-data.js';

function compatibilityNumberingMarkerInput(
  numbering: Readonly<NumberingInfo>,
  fallbackFontSizePt: number,
): NumberingMarkerShapeInput {
  const ascii = numbering.fontFamily ?? null;
  return {
    fontSizePt: fallbackFontSizePt,
    fonts: {
      ascii,
      highAnsi: ascii,
      eastAsia: numbering.fontFamilyEastAsia ?? ascii,
      complexScript: ascii,
    },
    weight: 400,
    style: 'normal',
    complexScript: false,
  };
}

type NumberingMarkerProjector = (
  numbering: NumberingInfo,
  fallbackFontSizePt: number,
) => NumberingMarkerShapeInput;

export interface NormalizedTextBoxRunInput {
  readonly text: string;
  readonly fontSizePt: number;
  readonly color?: string;
  readonly fontFamily?: string;
  readonly fontFamilyEastAsia?: string;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly ruby?: import('../types.js').RubyAnnotation | null;
  readonly rtl?: boolean;
}

export interface NormalizedTextBoxParagraphInput {
  readonly source: SourceRef;
  readonly spacing: Readonly<{ beforePt: number; afterPt: number }>;
  readonly alignment: string;
  readonly indentLeftPt: number;
  readonly indentRightPt: number;
  readonly indentFirstPt: number;
  readonly lineSpacing: Readonly<{ value: number; rule: 'auto' | 'exact' | 'atLeast'; explicit?: boolean }> | null;
  readonly tabStops: readonly Readonly<TabStop>[];
  readonly bidi?: boolean;
  readonly contextualSpacing: boolean;
  readonly styleId?: string | null;
  readonly numbering?: import('../types.js').NumberingInfo | null;
  readonly numberingMarkerShapeInput?: import('./types.js').NumberingMarkerShapeInput;
  readonly image?: Readonly<{
    imagePath: string;
    mimeType: string;
    svgImagePath?: string;
    widthPt: number;
    heightPt: number;
  }>;
  readonly runs: readonly NormalizedTextBoxRunInput[];
}

export type CompleteTextBoxBlockInput =
  | DeepReadonly<Extract<BodyElement, { type: 'paragraph' | 'table' }>>
  | ParsedUnsupportedTextBoxBlock;

/** Immutable parser/model seam. Parser-produced shapes retain the complete rich
 * block sequence; hand-built public shapes retain the deliberately lossy A3
 * paragraph projection as an explicit compatibility variant. */
export type TextBoxAcquisitionInput =
  | Readonly<{
      kind: 'complete';
      source: SourceRef;
      blockCount: number;
    }>
  | Readonly<{
      kind: 'compatibility';
      source: SourceRef;
      paragraphs: readonly NormalizedTextBoxParagraphInput[];
    }>;

/**
 * Converts the deliberately lossy public `ShapeRun.textBlocks` fallback to the
 * single parser-independent paragraph input boundary. ECMA-376 §20.4.2.38 rich
 * `txbxContent` remains B2's owner; this adapter never guesses glyph geometry.
 */
export function normalizeTextBoxInput(
  shape: DeepReadonly<ShapeRun>,
  source: SourceRef = { story: 'textbox', storyInstance: 'shape', path: [] },
  numberingMarkerProjector: NumberingMarkerProjector = compatibilityNumberingMarkerInput,
): readonly NormalizedTextBoxParagraphInput[] {
  const projected = (shape.textBlocks ?? []).map((block, index): NormalizedTextBoxParagraphInput => {
    const blockSource: SourceRef = {
      story: 'textbox',
      storyInstance: source.storyInstance,
      path: [...source.path, index],
    };
    const runs = block.runs?.length ? block.runs : [{
      text: block.text,
      fontSizePt: block.fontSizePt,
      color: block.color,
      fontFamily: block.fontFamily,
      bold: block.bold,
      italic: block.italic,
    }];
    const paragraphMarkColor = block.paragraphMarkColor === undefined
      ? block.color ?? runs[0]?.color
      : block.paragraphMarkColor ?? undefined;
    const numbering = block.numbering
      ? {
          ...block.numbering,
          ...(block.numbering.color == null && !block.numbering.colorAuto && paragraphMarkColor
            ? { color: paragraphMarkColor }
            : {}),
        }
      : null;
    return {
      source: blockSource,
      spacing: { beforePt: block.spaceBefore ?? 0, afterPt: block.spaceAfter ?? 0 },
      runs: runs.map((run) => ({
        text: run.text,
        fontSizePt: run.fontSizePt,
        ...(run.color ?? block.color ? { color: `#${run.color ?? block.color}` } : {}),
        ...(run.fontFamily || block.fontFamily ? { fontFamily: run.fontFamily ?? block.fontFamily ?? undefined } : {}),
        ...(run.fontFamilyEastAsia ? { fontFamilyEastAsia: run.fontFamilyEastAsia } : {}),
        bold: run.bold ?? block.bold ?? false,
        italic: run.italic ?? block.italic ?? false,
        ...(run.ruby ? { ruby: run.ruby } : {}),
        ...(run.rtl === undefined ? {} : { rtl: run.rtl }),
      })),
      alignment: block.alignment ?? 'left',
      indentLeftPt: block.indentLeft ?? 0,
      indentRightPt: block.indentRight ?? 0,
      indentFirstPt: block.indentFirst ?? 0,
      lineSpacing: block.lineSpacingVal == null ? null : {
          value: block.lineSpacingVal,
          rule: block.lineSpacingRule === 'exact' || block.lineSpacingRule === 'atLeast'
            ? block.lineSpacingRule : 'auto',
          explicit: true,
        },
      tabStops: (block.tabStops ?? []).map((stop) => ({ ...stop })),
      ...(block.bidi === undefined ? {} : { bidi: block.bidi }),
      contextualSpacing: block.contextualSpacing ?? false,
      ...(block.styleId === undefined ? {} : { styleId: block.styleId }),
      ...(numbering ? {
        numbering,
        // Hand-built public ShapeRun values have no parser-only effective lvl
        // rPr snapshot. Keep their historical public font fallback; parser
        // input replaces this at its projection boundary before acquisition.
        numberingMarkerShapeInput: numberingMarkerProjector(numbering, block.fontSizePt),
      } : {}),
      ...(block.imagePath ? { image: {
        imagePath: block.imagePath,
        mimeType: block.mimeType ?? 'application/octet-stream',
        ...(block.svgImagePath ? { svgImagePath: block.svgImagePath } : {}),
        widthPt: block.imageWidthPt ?? 0,
        heightPt: block.imageHeightPt ?? 0,
      } } : {}),
    };
  });
  return snapshotPlainData(projected, 'DOCX text box acquisition input');
}

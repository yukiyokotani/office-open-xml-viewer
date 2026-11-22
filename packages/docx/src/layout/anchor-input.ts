/** Parser-private, structured-clone-safe anchor facts. These declarations stay
 * inside retained layout and do not widen the stable public run model. */
export type AnchorValueStatus = 'missing' | 'invalid' | 'valid';

export interface AnchorEdgesInput {
  readonly topPt: number | null;
  readonly topStatus: AnchorValueStatus;
  readonly rightPt: number | null;
  readonly rightStatus: AnchorValueStatus;
  readonly bottomPt: number | null;
  readonly bottomStatus: AnchorValueStatus;
  readonly leftPt: number | null;
  readonly leftStatus: AnchorValueStatus;
}

export type AnchorAxisChoiceInput =
  | { readonly kind: 'missing' | 'invalid' }
  | { readonly kind: 'align'; readonly value: string }
  | { readonly kind: 'offset'; readonly valuePt: number }
  | { readonly kind: 'percent'; readonly fraction: number };

export interface AnchorRawTransformInput {
  readonly offsetXEmu: number | null;
  readonly offsetYEmu: number | null;
  readonly extentWidthEmu: number | null;
  readonly extentHeightEmu: number | null;
  readonly childOffsetXEmu: number | null;
  readonly childOffsetYEmu: number | null;
  readonly childExtentWidthEmu: number | null;
  readonly childExtentHeightEmu: number | null;
  readonly rotationUnits: number | null;
  readonly flipH: boolean | null;
  readonly flipV: boolean | null;
}

export interface AnchorAcquisitionInput {
  readonly occurrenceId: string;
  readonly simplePosition: {
    readonly enabled: boolean | null; readonly status: AnchorValueStatus;
    readonly xPt: number | null; readonly xStatus: AnchorValueStatus;
    readonly yPt: number | null; readonly yStatus: AnchorValueStatus;
  };
  readonly horizontal: {
    readonly relativeFrom: string | null; readonly relativeFromStatus: AnchorValueStatus;
    readonly choice: AnchorAxisChoiceInput;
  };
  readonly vertical: {
    readonly relativeFrom: string | null; readonly relativeFromStatus: AnchorValueStatus;
    readonly choice: AnchorAxisChoiceInput;
  };
  readonly extent: {
    readonly widthPt: number | null; readonly heightPt: number | null;
    readonly widthStatus: AnchorValueStatus; readonly heightStatus: AnchorValueStatus;
  };
  readonly parentEffectExtent: AnchorEdgesInput;
  readonly anchorDistances: AnchorEdgesInput;
  readonly relativeSize: {
    readonly horizontal: {
      readonly relativeFrom: string | null; readonly relativeFromStatus: AnchorValueStatus;
      readonly fraction: number | null; readonly fractionStatus: AnchorValueStatus;
    } | null;
    readonly vertical: {
      readonly relativeFrom: string | null; readonly relativeFromStatus: AnchorValueStatus;
      readonly fraction: number | null; readonly fractionStatus: AnchorValueStatus;
    } | null;
  };
  readonly wrap: {
    readonly kind: 'missing' | 'invalid' | 'none' | 'square' | 'tight' | 'through' | 'topAndBottom';
    readonly authoredKinds: readonly string[];
    readonly side: string | null;
    readonly distances: AnchorEdgesInput;
    readonly effectExtent: AnchorEdgesInput | null;
    readonly polygon: {
      readonly edited: boolean;
      readonly coordinateSpace: { readonly width: 21600; readonly height: 21600 };
      readonly points: readonly {
        readonly x: number | null; readonly y: number | null;
        readonly rawX: string | null; readonly rawY: string | null;
      }[];
      readonly invalidPointCount: number;
    } | null;
  };
  readonly behavior: {
    readonly behindDoc: boolean | null; readonly behindDocStatus: AnchorValueStatus;
    readonly relativeHeight: number | null; readonly relativeHeightStatus: AnchorValueStatus;
    readonly locked: boolean | null; readonly lockedStatus: AnchorValueStatus;
    readonly allowOverlap: boolean | null; readonly allowOverlapStatus: AnchorValueStatus;
    readonly layoutInCell: boolean | null; readonly layoutInCellStatus: AnchorValueStatus;
  };
  readonly group: {
    readonly childSourceId: string;
    readonly sourceIndex: number;
    readonly sourceCount: number;
    readonly transformChain: readonly AnchorRawTransformInput[];
    readonly childTransform: AnchorRawTransformInput | null;
    readonly resolvedChildFrame: {
      readonly offsetXPt: number;
      readonly offsetYPt: number;
      readonly widthPt: number;
      readonly heightPt: number;
      readonly rotationDeg: number;
      readonly flipH: boolean;
      readonly flipV: boolean;
    };
  } | null;
  /** WORD_STORY_VML_PAGE_POSITION: a text-box story's text- or line-relative
   * wrapping VML shape stays at or below the page's top margin. */
  readonly storyTopMarginFloor?: boolean;
  /** WORD_STORY_VML_PAGE_POSITION: a text-box or note story's VML shape
   * placed from the page origin, painted and indexed in its text box's host
   * frame. */
  readonly storyPagePosition?: boolean;
}

export interface InternalAnchorRunWire {
  readonly __anchorAcquisition?: AnchorAcquisitionInput;
}

import type {
  AnchorAcquisitionInput,
  AnchorAxisChoiceInput,
  AnchorEdgesInput,
  AnchorRawTransformInput,
} from './anchor-input.js';
import {
  WORD_ZERO_RELATIVE_SIZE_EXTENT_FALLBACK,
  wordZeroRelativeSizeUsesExtent,
} from './anchor-compatibility.js';
import { snapshotPlainData } from './plain-data.js';

export interface AnchorFrameRect {
  xPt: number;
  yPt: number;
  widthPt: number;
  heightPt: number;
}

export interface AnchorReferenceFramesInput {
  page: AnchorFrameRect | null;
  margin: AnchorFrameRect | null;
  column: AnchorFrameRect | null;
  paragraph: AnchorFrameRect | null;
  line: AnchorFrameRect | null;
  character: AnchorFrameRect | null;
  pageParity: 'odd' | 'even' | null;
}

export interface AnchorFrameInput {
  readonly acquisition: Readonly<AnchorAcquisitionInput>;
  readonly frames: AnchorReferenceFramesInput;
}

export type AnchorFrameIssueCode =
  | 'invalid-simple-position'
  | 'missing-simple-coordinate'
  | 'missing-relative-from'
  | 'invalid-relative-from'
  | 'unsupported-relative-from'
  | 'missing-reference-frame'
  | 'invalid-reference-frame'
  | 'missing-page-parity'
  | 'missing-axis-choice'
  | 'invalid-axis-choice'
  | 'invalid-axis-value'
  | 'missing-size'
  | 'invalid-size'
  | 'missing-relative-size-reference'
  | 'invalid-relative-size-reference'
  | 'missing-relative-size-fraction'
  | 'invalid-relative-size-fraction'
  | 'invalid-effect-extent'
  | 'invalid-distance'
  | 'missing-wrap-kind'
  | 'invalid-wrap-kind'
  | 'invalid-wrap-side'
  | 'invalid-wrap-polygon'
  | 'missing-required-behavior'
  | 'invalid-required-behavior';

export interface AnchorFrameIssue {
  readonly code: AnchorFrameIssueCode;
  readonly path: string;
  readonly message: string;
}

export type AnchorAxisDiagnostic = Readonly<{
  axis: 'horizontal' | 'vertical';
  status: 'resolved';
  relativeFrom: string;
  referenceFrame:
    | 'page'
    | 'margin'
    | 'column'
    | 'paragraph'
    | 'line'
    | 'character'
    | 'leftMargin'
    | 'rightMargin'
    | 'topMargin'
    | 'bottomMargin';
  choiceKind: 'simple-position' | 'align' | 'offset' | 'percent';
  choiceValue: string | number;
  baseStartPt: number;
  baseEndPt: number;
  resolvedOriginPt: number;
  pageParity: 'odd' | 'even' | null;
}> | Readonly<{
  axis: 'horizontal' | 'vertical';
  status: 'unsupported';
  relativeFrom: string | null;
  choiceKind: AnchorAxisChoiceInput['kind'] | 'simple-position';
  choiceValue: string | number | null;
  issueCode: AnchorFrameIssueCode;
}>;

type ResolvedAxisDiagnostic = Extract<AnchorAxisDiagnostic, { status: 'resolved' }>;

export interface AnchorSizeDiagnostic {
  readonly source: 'extent' | 'relative';
  readonly valuePt: number;
  readonly relativeFrom: string | null;
  readonly referenceFrame: ResolvedAxisDiagnostic['referenceFrame'] | null;
  readonly fraction: number | null;
  readonly compatibilityFallback?: typeof WORD_ZERO_RELATIVE_SIZE_EXTENT_FALLBACK.id;
}

export interface AnchorEffectiveEdges {
  readonly topPt: number;
  readonly rightPt: number;
  readonly bottomPt: number;
  readonly leftPt: number;
}

export interface AnchorWrapGeometry {
  readonly kind: 'none' | 'square' | 'tight' | 'through' | 'topAndBottom';
  readonly side: 'bothSides' | 'left' | 'right' | 'largest' | null;
  readonly distances: AnchorEffectiveEdges;
  readonly distanceSources: Readonly<{
    top: 'anchor' | 'wrap' | 'implicit-zero';
    right: 'anchor' | 'wrap' | 'implicit-zero';
    bottom: 'anchor' | 'wrap' | 'implicit-zero';
    left: 'anchor' | 'wrap' | 'implicit-zero';
  }>;
  readonly effectExtent: AnchorEffectiveEdges;
  readonly effectExtentSource: 'parent' | 'wrap-child' | 'none';
  readonly coordinateSpace: { readonly width: 21600; readonly height: 21600 } | null;
  readonly polygon: Readonly<{
    edited: boolean;
    points: readonly Readonly<{ xPt: number; yPt: number }>[];
  }> | null;
}

export interface AnchorTransformMetadata {
  readonly coordinateSpace: 'anchor-frame';
  readonly groupApplication: 'parser-resolved-child-frame';
  readonly group: Readonly<{
    childSourceId: string;
    sourceIndex: number;
    sourceCount: number;
    transformChain: readonly Readonly<AnchorRawTransformInput>[];
    childTransform: Readonly<AnchorRawTransformInput> | null;
    resolvedChildFrame: Readonly<{
      offsetXPt: number;
      offsetYPt: number;
      widthPt: number;
      heightPt: number;
      rotationDeg: number;
      flipH: boolean;
      flipV: boolean;
    }>;
  }> | null;
}

export interface AnchorFrameGeometry {
  readonly objectFrame: Readonly<AnchorFrameRect>;
  readonly inkBounds: Readonly<AnchorFrameRect>;
  readonly wrapBounds: Readonly<AnchorFrameRect> | null;
  readonly size: Readonly<{
    horizontal: Readonly<AnchorSizeDiagnostic>;
    vertical: Readonly<AnchorSizeDiagnostic>;
  }>;
  readonly parentEffectExtent: AnchorEffectiveEdges;
  readonly wrap: Readonly<AnchorWrapGeometry>;
  readonly transform: Readonly<AnchorTransformMetadata>;
}

export type AnchorFrameResult = Readonly<{
  status: 'resolved';
  occurrenceId: string;
  axes: Readonly<{
    horizontal: AnchorAxisDiagnostic;
    vertical: AnchorAxisDiagnostic;
  }>;
  issues: readonly AnchorFrameIssue[];
  geometry: Readonly<AnchorFrameGeometry>;
}> | Readonly<{
  status: 'unsupported';
  occurrenceId: string;
  axes: Readonly<{
    horizontal: AnchorAxisDiagnostic;
    vertical: AnchorAxisDiagnostic;
  }>;
  issues: readonly AnchorFrameIssue[];
}>;

type Axis = 'horizontal' | 'vertical';
type ReferenceFrameName = ResolvedAxisDiagnostic['referenceFrame'];

interface AxisBase {
  readonly startPt: number;
  readonly endPt: number;
  readonly referenceFrame: ReferenceFrameName;
}

interface ResolvedSize {
  readonly valuePt: number;
  readonly diagnostic: AnchorSizeDiagnostic;
}

interface EffectiveEdgeResult {
  readonly values: AnchorEffectiveEdges;
  readonly sources: AnchorWrapGeometry['distanceSources'];
}

const FIXED_WRAP_COORDINATE_SIZE = 21600;

function issue(
  code: AnchorFrameIssueCode,
  path: string,
  message: string,
): AnchorFrameIssue {
  return { code, path, message };
}

function finite(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value);
}

function rectIsValid(rect: AnchorFrameRect): boolean {
  return finite(rect.xPt)
    && finite(rect.yPt)
    && finite(rect.widthPt)
    && finite(rect.heightPt)
    && rect.widthPt >= 0
    && rect.heightPt >= 0;
}

function choiceValue(choice: AnchorAxisChoiceInput): string | number | null {
  if (choice.kind === 'align') return choice.value;
  if (choice.kind === 'offset') return choice.valuePt;
  if (choice.kind === 'percent') return choice.fraction;
  return null;
}

function unsupportedAxis(
  axis: Axis,
  acquisition: Readonly<AnchorAcquisitionInput>,
  problem: AnchorFrameIssue,
  simple = false,
): AnchorAxisDiagnostic {
  const value = acquisition[axis];
  return {
    axis,
    status: 'unsupported',
    relativeFrom: simple ? 'page' : value.relativeFrom,
    choiceKind: simple ? 'simple-position' : value.choice.kind,
    choiceValue: simple
      ? (axis === 'horizontal'
        ? acquisition.simplePosition.xPt
        : acquisition.simplePosition.yPt)
      : choiceValue(value.choice),
    issueCode: problem.code,
  };
}

function frameForDirectReference(
  name: 'page' | 'margin' | 'column' | 'paragraph' | 'line' | 'character',
  axis: Axis,
  frames: AnchorReferenceFramesInput,
  path: string,
): { base?: AxisBase; problem?: AnchorFrameIssue } {
  const frame = frames[name];
  if (frame === null) {
    return { problem: issue('missing-reference-frame', path, `${name} frame is required`) };
  }
  if (!rectIsValid(frame)) {
    return { problem: issue('invalid-reference-frame', path, `${name} frame must be finite and non-negative`) };
  }
  return {
    base: {
      startPt: axis === 'horizontal' ? frame.xPt : frame.yPt,
      endPt: axis === 'horizontal'
        ? frame.xPt + frame.widthPt
        : frame.yPt + frame.heightPt,
      referenceFrame: name,
    },
  };
}

function marginEdgeBase(
  edge: 'leftMargin' | 'rightMargin' | 'topMargin' | 'bottomMargin',
  axis: Axis,
  frames: AnchorReferenceFramesInput,
  path: string,
): { base?: AxisBase; problem?: AnchorFrameIssue } {
  const pageResult = frameForDirectReference('page', axis, frames, path);
  if (!pageResult.base) return pageResult;
  const marginResult = frameForDirectReference('margin', axis, frames, path);
  if (!marginResult.base) return marginResult;
  const page = frames.page as AnchorFrameRect;
  const margin = frames.margin as AnchorFrameRect;
  const horizontal = edge === 'leftMargin' || edge === 'rightMargin';
  if (horizontal !== (axis === 'horizontal')) {
    return {
      problem: issue(
        'unsupported-relative-from',
        path,
        `${edge} is not valid for the ${axis} axis`,
      ),
    };
  }
  const pageStart = horizontal ? page.xPt : page.yPt;
  const pageEnd = horizontal
    ? page.xPt + page.widthPt
    : page.yPt + page.heightPt;
  const marginStart = horizontal ? margin.xPt : margin.yPt;
  const marginEnd = horizontal
    ? margin.xPt + margin.widthPt
    : margin.yPt + margin.heightPt;
  if (marginStart < pageStart || marginEnd > pageEnd) {
    return {
      problem: issue(
        'invalid-reference-frame',
        path,
        'margin frame must be contained by the page frame',
      ),
    };
  }
  const leading = edge === 'leftMargin' || edge === 'topMargin';
  return {
    base: {
      startPt: leading ? pageStart : marginEnd,
      endPt: leading ? marginStart : pageEnd,
      referenceFrame: edge,
    },
  };
}

function axisBase(
  axis: Axis,
  relativeFrom: string,
  frames: AnchorReferenceFramesInput,
  path: string,
): { base?: AxisBase; problem?: AnchorFrameIssue; parityRequired?: boolean } {
  if (relativeFrom === 'page' || relativeFrom === 'margin') {
    return frameForDirectReference(relativeFrom, axis, frames, path);
  }
  if (axis === 'horizontal' && (relativeFrom === 'column' || relativeFrom === 'character')) {
    return frameForDirectReference(relativeFrom, axis, frames, path);
  }
  if (axis === 'vertical' && (relativeFrom === 'paragraph' || relativeFrom === 'line')) {
    return frameForDirectReference(relativeFrom, axis, frames, path);
  }
  if (axis === 'horizontal' && (relativeFrom === 'leftMargin' || relativeFrom === 'rightMargin')) {
    return marginEdgeBase(relativeFrom, axis, frames, path);
  }
  if (axis === 'vertical' && (relativeFrom === 'topMargin' || relativeFrom === 'bottomMargin')) {
    return marginEdgeBase(relativeFrom, axis, frames, path);
  }
  if (relativeFrom === 'insideMargin' || relativeFrom === 'outsideMargin') {
    if (frames.pageParity === null) {
      return {
        problem: issue(
          'missing-page-parity',
          path,
          `${relativeFrom} requires explicit page parity`,
        ),
      };
    }
    const inside = relativeFrom === 'insideMargin';
    const leading = inside === (frames.pageParity === 'odd');
    const edge = axis === 'horizontal'
      ? (leading ? 'leftMargin' : 'rightMargin')
      : (leading ? 'topMargin' : 'bottomMargin');
    return { ...marginEdgeBase(edge, axis, frames, path), parityRequired: true };
  }
  return {
    problem: issue(
      'unsupported-relative-from',
      path,
      `${relativeFrom} is not a valid ${axis} reference`,
    ),
  };
}

function resolveSize(
  axis: Axis,
  acquisition: Readonly<AnchorAcquisitionInput>,
  frames: AnchorReferenceFramesInput,
): { resolved?: ResolvedSize; problem?: AnchorFrameIssue } {
  const relative = acquisition.relativeSize[axis];
  const axisName = axis === 'horizontal' ? 'width' : 'height';
  const resolveExtent = (
    compatibilityRelative: Readonly<{ relativeFrom: string | null; fraction: number }> | null = null,
  ): { resolved?: ResolvedSize; problem?: AnchorFrameIssue } => {
    const status = axis === 'horizontal'
      ? acquisition.extent.widthStatus
      : acquisition.extent.heightStatus;
    const value = axis === 'horizontal'
      ? acquisition.extent.widthPt
      : acquisition.extent.heightPt;
    if (status === 'missing') {
      return { problem: issue('missing-size', `extent.${axisName}`, `${axisName} is required`) };
    }
    if (status !== 'valid' || !finite(value) || value <= 0) {
      return {
        problem: issue(
          'invalid-size',
          `extent.${axisName}`,
          `${axisName} extent must be finite and positive`,
        ),
      };
    }
    return {
      resolved: {
        valuePt: value,
        diagnostic: {
          source: 'extent',
          valuePt: value,
          relativeFrom: compatibilityRelative?.relativeFrom ?? null,
          referenceFrame: null,
          fraction: compatibilityRelative?.fraction ?? null,
          ...(compatibilityRelative === null
            ? {}
            : { compatibilityFallback: WORD_ZERO_RELATIVE_SIZE_EXTENT_FALLBACK.id }),
        },
      },
    };
  };
  if (relative === null) {
    return resolveExtent();
  }
  const path = `relativeSize.${axis}`;
  if (relative.fractionStatus === 'missing' || relative.fraction === null) {
    return {
      problem: issue(
        'missing-relative-size-fraction',
        `${path}.fraction`,
        'relative size fraction is required',
      ),
    };
  }
  if (relative.fractionStatus !== 'valid' || !finite(relative.fraction)) {
    return {
      problem: issue(
        'invalid-relative-size-fraction',
        `${path}.fraction`,
        'relative size fraction must be finite',
      ),
    };
  }
  if (relative.fraction < 0) {
    return {
      problem: issue(
        'invalid-relative-size-fraction',
        `${path}.fraction`,
        'relative size fraction must be non-negative',
      ),
    };
  }
  if (wordZeroRelativeSizeUsesExtent(relative.fraction)) {
    // Compatibility-owned zero-relative-size fallback; acquisition still
    // retains the authored value while geometry uses wp:extent.
    return resolveExtent({
      relativeFrom: relative.relativeFrom,
      fraction: relative.fraction,
    });
  }
  if (relative.relativeFromStatus === 'missing' || relative.relativeFrom === null) {
    return {
      problem: issue(
        'missing-relative-size-reference',
        `${path}.relativeFrom`,
        'relative size reference is required',
      ),
    };
  }
  if (relative.relativeFromStatus !== 'valid') {
    return {
      problem: issue(
        'invalid-relative-size-reference',
        `${path}.relativeFrom`,
        'relative size reference is invalid',
      ),
    };
  }
  const baseResult = axisBase(axis, relative.relativeFrom, frames, `${path}.relativeFrom`);
  if (!baseResult.base) {
    return {
      problem: issue(
        baseResult.problem?.code === 'missing-reference-frame'
          ? 'missing-relative-size-reference'
          : 'invalid-relative-size-reference',
        `${path}.relativeFrom`,
        baseResult.problem?.message ?? 'relative size reference cannot be resolved',
      ),
    };
  }
  const valuePt = (baseResult.base.endPt - baseResult.base.startPt) * relative.fraction;
  if (!finite(valuePt) || valuePt < 0) {
    return {
      problem: issue(
        'invalid-relative-size-fraction',
        `${path}.fraction`,
        'relative size result must be finite and non-negative',
      ),
    };
  }
  return {
    resolved: {
      valuePt,
      diagnostic: {
        source: 'relative',
        valuePt,
        relativeFrom: relative.relativeFrom,
        referenceFrame: baseResult.base.referenceFrame,
        fraction: relative.fraction,
      },
    },
  };
}

function resolveAxis(
  axis: Axis,
  sizePt: number,
  acquisition: Readonly<AnchorAcquisitionInput>,
  frames: AnchorReferenceFramesInput,
): { valuePt?: number; diagnostic?: AnchorAxisDiagnostic; problem?: AnchorFrameIssue } {
  const authored = acquisition[axis];
  const path = axis;
  if (authored.relativeFromStatus === 'missing' || authored.relativeFrom === null) {
    const problem = issue(
      'missing-relative-from',
      `${path}.relativeFrom`,
      `${axis} relativeFrom is required`,
    );
    return { diagnostic: unsupportedAxis(axis, acquisition, problem), problem };
  }
  if (authored.relativeFromStatus !== 'valid') {
    const problem = issue(
      'invalid-relative-from',
      `${path}.relativeFrom`,
      `${axis} relativeFrom is invalid`,
    );
    return { diagnostic: unsupportedAxis(axis, acquisition, problem), problem };
  }
  const baseResult = axisBase(axis, authored.relativeFrom, frames, `${path}.relativeFrom`);
  if (!baseResult.base) {
    const problem = baseResult.problem as AnchorFrameIssue;
    return { diagnostic: unsupportedAxis(axis, acquisition, problem), problem };
  }
  const choice = authored.choice;
  if (choice.kind === 'missing') {
    const problem = issue('missing-axis-choice', `${path}.choice`, `${axis} choice is required`);
    return { diagnostic: unsupportedAxis(axis, acquisition, problem), problem };
  }
  if (choice.kind === 'invalid') {
    const problem = issue('invalid-axis-choice', `${path}.choice`, `${axis} choice is invalid`);
    return { diagnostic: unsupportedAxis(axis, acquisition, problem), problem };
  }
  const lengthPt = baseResult.base.endPt - baseResult.base.startPt;
  let resolvedOriginPt: number;
  let value: string | number;
  if (choice.kind === 'offset') {
    if (!finite(choice.valuePt)) {
      const problem = issue('invalid-axis-value', `${path}.choice`, 'offset must be finite');
      return { diagnostic: unsupportedAxis(axis, acquisition, problem), problem };
    }
    resolvedOriginPt = baseResult.base.startPt + choice.valuePt;
    value = choice.valuePt;
  } else if (choice.kind === 'percent') {
    if (!finite(choice.fraction)) {
      const problem = issue('invalid-axis-value', `${path}.choice`, 'percentage must be finite');
      return { diagnostic: unsupportedAxis(axis, acquisition, problem), problem };
    }
    resolvedOriginPt = baseResult.base.startPt + lengthPt * choice.fraction;
    value = choice.fraction;
  } else if (choice.kind === 'align') {
    const valid = axis === 'horizontal'
      ? ['left', 'center', 'right', 'inside', 'outside'].includes(choice.value)
      : ['top', 'center', 'bottom', 'inside', 'outside'].includes(choice.value);
    if (!valid) {
      const problem = issue('invalid-axis-value', `${path}.choice`, `${choice.value} is invalid`);
      return { diagnostic: unsupportedAxis(axis, acquisition, problem), problem };
    }
    if ((choice.value === 'inside' || choice.value === 'outside') && frames.pageParity === null) {
      const problem = issue(
        'missing-page-parity',
        'frames.pageParity',
        `${choice.value} alignment requires explicit page parity`,
      );
      return { diagnostic: unsupportedAxis(axis, acquisition, problem), problem };
    }
    const placement = alignedAnchorPlacement(axis, choice.value, frames.pageParity);
    resolvedOriginPt = placement === 'leading'
      ? baseResult.base.startPt
      : placement === 'trailing'
        ? baseResult.base.endPt - sizePt
        : baseResult.base.startPt + (lengthPt - sizePt) / 2;
    value = choice.value;
  } else {
    const problem = issue('invalid-axis-choice', `${path}.choice`, `${axis} choice is invalid`);
    return { diagnostic: unsupportedAxis(axis, acquisition, problem), problem };
  }
  if (!finite(resolvedOriginPt)) {
    const problem = issue('invalid-axis-value', `${path}.choice`, 'resolved origin is not finite');
    return { diagnostic: unsupportedAxis(axis, acquisition, problem), problem };
  }
  return {
    valuePt: resolvedOriginPt,
    diagnostic: {
      axis,
      status: 'resolved',
      relativeFrom: authored.relativeFrom,
      referenceFrame: baseResult.base.referenceFrame,
      choiceKind: choice.kind,
      choiceValue: value,
      baseStartPt: baseResult.base.startPt,
      baseEndPt: baseResult.base.endPt,
      resolvedOriginPt,
      pageParity: choice.kind === 'align'
        && (choice.value === 'inside' || choice.value === 'outside')
        ? frames.pageParity
        : null,
    },
  };
}

/**
 * Which edge of the alignment base an ECMA-376 §20.4.3.1 `wp:align` value
 * holds the object against. `inside`/`outside` follow the page parity
 * (odd pages lead with inside). Callers validate the value beforehand.
 */
export function alignedAnchorPlacement(
  axis: Axis,
  value: string,
  pageParity: 'odd' | 'even' | null,
): 'leading' | 'center' | 'trailing' {
  const leadingName = axis === 'horizontal' ? 'left' : 'top';
  const trailingName = axis === 'horizontal' ? 'right' : 'bottom';
  const parityLeading = pageParity === 'odd';
  if (
    value === leadingName
    || (value === 'inside' && parityLeading)
    || (value === 'outside' && !parityLeading)
  ) {
    return 'leading';
  }
  if (
    value === trailingName
    || (value === 'inside' && !parityLeading)
    || (value === 'outside' && parityLeading)
  ) {
    return 'trailing';
  }
  return 'center';
}

function simpleAxis(
  axis: Axis,
  coordinatePt: number,
  page: AnchorFrameRect,
): { valuePt: number; diagnostic: AnchorAxisDiagnostic } {
  const startPt = axis === 'horizontal' ? page.xPt : page.yPt;
  const endPt = axis === 'horizontal'
    ? page.xPt + page.widthPt
    : page.yPt + page.heightPt;
  const valuePt = startPt + coordinatePt;
  return {
    valuePt,
    diagnostic: {
      axis,
      status: 'resolved',
      relativeFrom: 'page',
      referenceFrame: 'page',
      choiceKind: 'simple-position',
      choiceValue: coordinatePt,
      baseStartPt: startPt,
      baseEndPt: endPt,
      resolvedOriginPt: valuePt,
      pageParity: null,
    },
  };
}

const EDGE_NAMES = ['top', 'right', 'bottom', 'left'] as const;
type EdgeName = typeof EDGE_NAMES[number];

function edgeStatus(edges: AnchorEdgesInput, edge: EdgeName) {
  return edges[`${edge}Status` as const];
}

function edgeValue(edges: AnchorEdgesInput, edge: EdgeName) {
  return edges[`${edge}Pt` as const];
}

function effectEdges(
  edges: Readonly<AnchorEdgesInput>,
  path: string,
  present: boolean,
): { values?: AnchorEffectiveEdges; problem?: AnchorFrameIssue } {
  const authored = EDGE_NAMES.some((edge) => edgeStatus(edges, edge) !== 'missing');
  if (!present && !authored) {
    return { values: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 } };
  }
  const values: Record<`${EdgeName}Pt`, number> = {
    topPt: 0,
    rightPt: 0,
    bottomPt: 0,
    leftPt: 0,
  };
  for (const edge of EDGE_NAMES) {
    const status = edgeStatus(edges, edge);
    const value = edgeValue(edges, edge);
    if (status !== 'valid' || !finite(value)) {
      return {
        problem: issue(
          'invalid-effect-extent',
          `${path}.${edge}`,
          'present effectExtent requires four finite edge values',
        ),
      };
    }
    values[`${edge}Pt`] = value;
  }
  return { values };
}

function effectiveDistances(
  anchorEdges: Readonly<AnchorEdgesInput>,
  wrapEdges: Readonly<AnchorEdgesInput>,
): { resolved?: EffectiveEdgeResult; problem?: AnchorFrameIssue } {
  const values: Record<`${EdgeName}Pt`, number> = {
    topPt: 0,
    rightPt: 0,
    bottomPt: 0,
    leftPt: 0,
  };
  const sources = {} as Record<EdgeName, 'anchor' | 'wrap' | 'implicit-zero'>;
  for (const edge of EDGE_NAMES) {
    const childStatus = edgeStatus(wrapEdges, edge);
    const parentStatus = edgeStatus(anchorEdges, edge);
    const selected = childStatus === 'valid'
      ? { status: childStatus, value: edgeValue(wrapEdges, edge), source: 'wrap' as const }
      : childStatus === 'invalid'
        ? { status: childStatus, value: edgeValue(wrapEdges, edge), source: 'wrap' as const }
        : parentStatus === 'valid'
          ? { status: parentStatus, value: edgeValue(anchorEdges, edge), source: 'anchor' as const }
          : parentStatus === 'invalid'
            ? { status: parentStatus, value: edgeValue(anchorEdges, edge), source: 'anchor' as const }
            : { status: 'missing' as const, value: null, source: 'implicit-zero' as const };
    if (
      selected.status === 'invalid'
      || (selected.status === 'valid' && (!finite(selected.value) || selected.value < 0))
    ) {
      return {
        problem: issue(
          'invalid-distance',
          `${selected.source === 'wrap' ? 'wrap.distances' : 'anchorDistances'}.${edge}`,
          'wrap distance must be finite and non-negative',
        ),
      };
    }
    values[`${edge}Pt`] = selected.status === 'missing' ? 0 : selected.value as number;
    sources[edge] = selected.source;
  }
  return { resolved: { values, sources } };
}

function expandBounds(
  rect: Readonly<AnchorFrameRect>,
  edges: Readonly<AnchorEffectiveEdges>,
): AnchorFrameRect | null {
  const expanded = {
    xPt: rect.xPt - edges.leftPt,
    yPt: rect.yPt - edges.topPt,
    widthPt: rect.widthPt + edges.leftPt + edges.rightPt,
    heightPt: rect.heightPt + edges.topPt + edges.bottomPt,
  };
  return rectIsValid(expanded) ? expanded : null;
}

function polygonGeometry(
  acquisition: Readonly<AnchorAcquisitionInput>,
  frame: Readonly<AnchorFrameRect>,
): { polygon?: NonNullable<AnchorWrapGeometry['polygon']>; bounds?: AnchorFrameRect; problem?: AnchorFrameIssue } {
  const polygon = acquisition.wrap.polygon;
  if (
    polygon === null
    || polygon.invalidPointCount !== 0
    || polygon.coordinateSpace.width !== FIXED_WRAP_COORDINATE_SIZE
    || polygon.coordinateSpace.height !== FIXED_WRAP_COORDINATE_SIZE
    || polygon.points.length < 3
  ) {
    return {
      problem: issue(
        'invalid-wrap-polygon',
        'wrap.polygon',
        'tight and through wrapping require a valid fixed 21600 by 21600 polygon',
      ),
    };
  }
  const points: { xPt: number; yPt: number }[] = [];
  for (const [index, point] of polygon.points.entries()) {
    if (!finite(point.x) || !finite(point.y)) {
      return {
        problem: issue(
          'invalid-wrap-polygon',
          `wrap.polygon.points.${index}`,
          'polygon coordinates must be finite',
        ),
      };
    }
    points.push({
      xPt: frame.xPt + (point.x / FIXED_WRAP_COORDINATE_SIZE) * frame.widthPt,
      yPt: frame.yPt + (point.y / FIXED_WRAP_COORDINATE_SIZE) * frame.heightPt,
    });
  }
  const xs = points.map((point) => point.xPt);
  const ys = points.map((point) => point.yPt);
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const minY = Math.min(...ys);
  const maxY = Math.max(...ys);
  return {
    polygon: { edited: polygon.edited, points },
    bounds: { xPt: minX, yPt: minY, widthPt: maxX - minX, heightPt: maxY - minY },
  };
}

function transformMetadata(
  group: AnchorAcquisitionInput['group'],
): AnchorTransformMetadata {
  return {
    coordinateSpace: 'anchor-frame',
    groupApplication: 'parser-resolved-child-frame',
    group: group === null ? null : {
      childSourceId: group.childSourceId,
      sourceIndex: group.sourceIndex,
      sourceCount: group.sourceCount,
      transformChain: group.transformChain.map((transform) => ({ ...transform })),
      childTransform: group.childTransform === null ? null : { ...group.childTransform },
      resolvedChildFrame: { ...group.resolvedChildFrame },
    },
  };
}

function frozenResult(result: AnchorFrameResult): AnchorFrameResult {
  return snapshotPlainData(result, 'anchor frame result') as AnchorFrameResult;
}

/**
 * Resolves only retained point-space anchor geometry. It deliberately does not
 * choose flow ownership, collision placement, z-order, paint transforms, or a
 * compatibility fallback. See ECMA-376 Part 1 §§20.4.2.3, .6-.20 and
 * §§20.4.3.1-.7; implementation-note evidence is owned by the compatibility
 * decision module.
 */
export function resolveAnchorFrame(input: AnchorFrameInput): AnchorFrameResult {
  const { acquisition, frames } = input;
  // CT_Anchor makes these authoring controls required even when a particular
  // layout stage does not consume their values. Rejecting incomplete facts at
  // acquisition prevents a drawable geometry from legitimizing malformed XML.
  for (const field of [
    'relativeHeight',
    'behindDoc',
    'locked',
    'layoutInCell',
    'allowOverlap',
  ] as const) {
    const status = acquisition.behavior[`${field}Status`];
    const value = acquisition.behavior[field];
    if (status === 'valid' && value !== null) continue;
    const missing = status === 'missing';
    const problem = issue(
      missing ? 'missing-required-behavior' : 'invalid-required-behavior',
      `behavior.${field}`,
      `CT_Anchor requires a ${field} value`,
    );
    return frozenResult({
      status: 'unsupported',
      occurrenceId: acquisition.occurrenceId,
      axes: {
        horizontal: unsupportedAxis('horizontal', acquisition, problem),
        vertical: unsupportedAxis('vertical', acquisition, problem),
      },
      issues: [problem],
    });
  }
  const issues: AnchorFrameIssue[] = [];
  const width = resolveSize('horizontal', acquisition, frames);
  const height = resolveSize('vertical', acquisition, frames);
  if (width.problem) issues.push(width.problem);
  if (height.problem) issues.push(height.problem);

  let horizontal: { valuePt?: number; diagnostic: AnchorAxisDiagnostic; problem?: AnchorFrameIssue };
  let vertical: { valuePt?: number; diagnostic: AnchorAxisDiagnostic; problem?: AnchorFrameIssue };
  const firstSizeProblem = width.problem ?? height.problem;
  if (firstSizeProblem || !width.resolved || !height.resolved) {
    const problem = firstSizeProblem as AnchorFrameIssue;
    horizontal = { diagnostic: unsupportedAxis('horizontal', acquisition, problem) };
    vertical = { diagnostic: unsupportedAxis('vertical', acquisition, problem) };
  } else if (acquisition.simplePosition.status === 'invalid') {
    const problem = issue(
      'invalid-simple-position',
      'simplePosition.enabled',
      'simplePos enablement is invalid',
    );
    issues.push(problem);
    horizontal = { diagnostic: unsupportedAxis('horizontal', acquisition, problem, true), problem };
    vertical = { diagnostic: unsupportedAxis('vertical', acquisition, problem, true), problem };
  } else if (acquisition.simplePosition.status === 'valid' && acquisition.simplePosition.enabled === true) {
    const pageResult = frameForDirectReference('page', 'horizontal', frames, 'frames.page');
    const x = acquisition.simplePosition.xPt;
    const y = acquisition.simplePosition.yPt;
    if (!pageResult.base || frames.page === null || !rectIsValid(frames.page)) {
      const problem = pageResult.problem ?? issue(
        'invalid-reference-frame',
        'frames.page',
        'simple positioning requires a valid page frame',
      );
      issues.push(problem);
      horizontal = { diagnostic: unsupportedAxis('horizontal', acquisition, problem, true), problem };
      vertical = { diagnostic: unsupportedAxis('vertical', acquisition, problem, true), problem };
    } else if (acquisition.simplePosition.xStatus !== 'valid' || !finite(x)) {
      const invalid = acquisition.simplePosition.xStatus === 'invalid';
      const problem = issue(
        invalid ? 'invalid-simple-position' : 'missing-simple-coordinate',
        'simplePosition.x',
        invalid ? 'simple position x is lexically invalid' : 'simple position x is required',
      );
      issues.push(problem);
      horizontal = { diagnostic: unsupportedAxis('horizontal', acquisition, problem, true), problem };
      vertical = { diagnostic: unsupportedAxis('vertical', acquisition, problem, true), problem };
    } else if (acquisition.simplePosition.yStatus !== 'valid' || !finite(y)) {
      const invalid = acquisition.simplePosition.yStatus === 'invalid';
      const problem = issue(
        invalid ? 'invalid-simple-position' : 'missing-simple-coordinate',
        'simplePosition.y',
        invalid ? 'simple position y is lexically invalid' : 'simple position y is required',
      );
      issues.push(problem);
      horizontal = { diagnostic: unsupportedAxis('horizontal', acquisition, problem, true), problem };
      vertical = { diagnostic: unsupportedAxis('vertical', acquisition, problem, true), problem };
    } else {
      horizontal = simpleAxis('horizontal', x, frames.page);
      vertical = simpleAxis('vertical', y, frames.page);
    }
  } else {
    const x = resolveAxis('horizontal', width.resolved.valuePt, acquisition, frames);
    const y = resolveAxis('vertical', height.resolved.valuePt, acquisition, frames);
    horizontal = { ...x, diagnostic: x.diagnostic as AnchorAxisDiagnostic };
    vertical = { ...y, diagnostic: y.diagnostic as AnchorAxisDiagnostic };
    if (x.problem) issues.push(x.problem);
    if (y.problem) issues.push(y.problem);
  }

  if (
    issues.length > 0
    || !width.resolved
    || !height.resolved
    || horizontal.valuePt === undefined
    || vertical.valuePt === undefined
  ) {
    return frozenResult({
      status: 'unsupported',
      occurrenceId: acquisition.occurrenceId,
      axes: { horizontal: horizontal.diagnostic, vertical: vertical.diagnostic },
      issues,
    });
  }

  const objectFrame: AnchorFrameRect = {
    xPt: horizontal.valuePt,
    yPt: vertical.valuePt,
    widthPt: width.resolved.valuePt,
    heightPt: height.resolved.valuePt,
  };
  const parentEffects = effectEdges(acquisition.parentEffectExtent, 'parentEffectExtent', false);
  if (parentEffects.problem || !parentEffects.values) {
    const problem = parentEffects.problem as AnchorFrameIssue;
    return frozenResult({
      status: 'unsupported',
      occurrenceId: acquisition.occurrenceId,
      axes: { horizontal: horizontal.diagnostic, vertical: vertical.diagnostic },
      issues: [problem],
    });
  }
  const inkBounds = expandBounds(objectFrame, parentEffects.values);
  if (inkBounds === null) {
    const problem = issue(
      'invalid-effect-extent',
      'parentEffectExtent',
      'parent effect extents produce invalid ink bounds',
    );
    return frozenResult({
      status: 'unsupported',
      occurrenceId: acquisition.occurrenceId,
      axes: { horizontal: horizontal.diagnostic, vertical: vertical.diagnostic },
      issues: [problem],
    });
  }
  if (acquisition.wrap.kind === 'missing' || acquisition.wrap.kind === 'invalid') {
    const problem = issue(
      acquisition.wrap.kind === 'missing' ? 'missing-wrap-kind' : 'invalid-wrap-kind',
      'wrap.kind',
      'exactly one valid wrap kind is required',
    );
    return frozenResult({
      status: 'unsupported',
      occurrenceId: acquisition.occurrenceId,
      axes: { horizontal: horizontal.diagnostic, vertical: vertical.diagnostic },
      issues: [problem],
    });
  }
  const distances = effectiveDistances(acquisition.anchorDistances, acquisition.wrap.distances);
  if (distances.problem || !distances.resolved) {
    return frozenResult({
      status: 'unsupported',
      occurrenceId: acquisition.occurrenceId,
      axes: { horizontal: horizontal.diagnostic, vertical: vertical.diagnostic },
      issues: [distances.problem as AnchorFrameIssue],
    });
  }
  const sideRequired = acquisition.wrap.kind === 'square'
    || acquisition.wrap.kind === 'tight'
    || acquisition.wrap.kind === 'through';
  const side = sideRequired
    && ['bothSides', 'left', 'right', 'largest'].includes(acquisition.wrap.side ?? '')
    ? acquisition.wrap.side as AnchorWrapGeometry['side']
    : null;
  if (sideRequired && side === null) {
    const problem = issue(
      'invalid-wrap-side',
      'wrap.side',
      'square, tight, and through wrapping require an authored wrap side',
    );
    return frozenResult({
      status: 'unsupported',
      occurrenceId: acquisition.occurrenceId,
      axes: { horizontal: horizontal.diagnostic, vertical: vertical.diagnostic },
      issues: [problem],
    });
  }

  let wrapEffect = parentEffects.values;
  let wrapEffectSource: AnchorWrapGeometry['effectExtentSource'] = EDGE_NAMES.some(
    (edge) => edgeStatus(acquisition.parentEffectExtent, edge) !== 'missing',
  ) ? 'parent' : 'none';
  if (acquisition.wrap.effectExtent !== null) {
    const childEffect = effectEdges(acquisition.wrap.effectExtent, 'wrap.effectExtent', true);
    if (childEffect.problem || !childEffect.values) {
      return frozenResult({
        status: 'unsupported',
        occurrenceId: acquisition.occurrenceId,
        axes: { horizontal: horizontal.diagnostic, vertical: vertical.diagnostic },
        issues: [childEffect.problem as AnchorFrameIssue],
      });
    }
    wrapEffect = childEffect.values;
    wrapEffectSource = 'wrap-child';
  }

  let polygon: AnchorWrapGeometry['polygon'] = null;
  let coordinateSpace: AnchorWrapGeometry['coordinateSpace'] = null;
  let wrapBase: AnchorFrameRect | null = null;
  if (acquisition.wrap.kind === 'tight' || acquisition.wrap.kind === 'through') {
    const mapped = polygonGeometry(acquisition, objectFrame);
    if (mapped.problem || !mapped.polygon || !mapped.bounds) {
      return frozenResult({
        status: 'unsupported',
        occurrenceId: acquisition.occurrenceId,
        axes: { horizontal: horizontal.diagnostic, vertical: vertical.diagnostic },
        issues: [mapped.problem as AnchorFrameIssue],
      });
    }
    polygon = mapped.polygon;
    coordinateSpace = { width: 21600, height: 21600 };
    wrapBase = mapped.bounds;
    // Polygon wrapping is defined by the polygon itself, not effectExtent.
    wrapEffect = { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 };
    wrapEffectSource = 'none';
  } else if (acquisition.wrap.kind !== 'none') {
    wrapBase = expandBounds(objectFrame, wrapEffect);
    if (wrapBase === null) {
      const problem = issue(
        'invalid-effect-extent',
        'wrap.effectExtent',
        'wrapping effect extents produce invalid bounds',
      );
      return frozenResult({
        status: 'unsupported',
        occurrenceId: acquisition.occurrenceId,
        axes: { horizontal: horizontal.diagnostic, vertical: vertical.diagnostic },
        issues: [problem],
      });
    }
  }
  const wrapBounds = wrapBase === null
    ? null
    : expandBounds(wrapBase, distances.resolved.values);
  if (wrapBase !== null && wrapBounds === null) {
    const problem = issue('invalid-distance', 'wrap.distances', 'distances produce invalid bounds');
    return frozenResult({
      status: 'unsupported',
      occurrenceId: acquisition.occurrenceId,
      axes: { horizontal: horizontal.diagnostic, vertical: vertical.diagnostic },
      issues: [problem],
    });
  }

  return frozenResult({
    status: 'resolved',
    occurrenceId: acquisition.occurrenceId,
    axes: { horizontal: horizontal.diagnostic, vertical: vertical.diagnostic },
    issues: [],
    geometry: {
      objectFrame,
      inkBounds,
      wrapBounds,
      size: {
        horizontal: width.resolved.diagnostic,
        vertical: height.resolved.diagnostic,
      },
      parentEffectExtent: parentEffects.values,
      wrap: {
        kind: acquisition.wrap.kind,
        side,
        distances: distances.resolved.values,
        distanceSources: distances.resolved.sources,
        effectExtent: wrapEffect,
        effectExtentSource: wrapEffectSource,
        coordinateSpace,
        polygon,
      },
      transform: transformMetadata(acquisition.group),
    },
  });
}

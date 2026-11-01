// Compatibility facade for legacy DOCX modules. The layout-owned implementation
// lives under layout/ so canonical layout code never depends on a root legacy
// module while existing internal import paths remain stable during migration.
export {
  MIN_LINE_GAP,
  floatOverlapsColumnX,
  isWrapFloat,
  normalizeWrapSide,
  prepareFloatWrap,
  rectsOverlap,
  resolveLineFloatWindow,
  computePreparedLineFloatWindow,
  computePreparedLineFloatWindowWithDiagnostics,
  skipPastTopAndBottom,
} from './layout/float-wrap.js';

export {
  FLOAT_OVERLAP_EPS,
  FLOAT_PAGE_RIGHT_SLACK,
} from './layout/floats.js';


export type {
  FloatRect,
  Gap,
  LineFloatReference,
  LineFloatSweepDiagnostics,
  LineWindowAdmission,
  PreparedFloatWrap,
  WrapSide,
} from './layout/float-wrap.js';

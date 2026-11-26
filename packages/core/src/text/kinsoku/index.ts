// ECMA-376 §17.15.1.58–.60 Japanese line-breaking (kinsoku) engine.
// Package-agnostic: operates on code points and split indices only.
export type { KinsokuRules } from './rules.js';
export { resolveKinsokuRules, DEFAULT_KINSOKU_RULES } from './rules.js';
export { kinsokuAdjustedSplit } from './split.js';

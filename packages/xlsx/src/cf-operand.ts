import type { Cell } from './types.js';

// ────────────────────────────────────────────────────────────────
// Conditional-formatting operand decoding (no formula evaluation)
//
// The operands of a `cellIs` rule and the activity condition of a
// colorScale / dataBar / iconSet rule ([MS-XLSX] 2.6.27) are formulas. A
// guessed value could stop lower-priority rules (ECMA-376 §18.3.1.10
// stopIfTrue), and this operand decoder reads cached values instead of
// evaluating formulas. So an operand is decoded in exactly two cases:
//
//   - a plain numeric literal (optional sign and exponent) or "string"
//     literal (`""` escapes a quote);
//   - a single A1 cell reference, relative or `$`-absolute, shifted by the
//     evaluated cell's offset from the rule range's top-left anchor. Its
//     value is that cell's cached value as the workbook stores it: number,
//     string, boolean or blank, kept distinct. A reference is a lookup of a
//     stored value, not evaluation; an error cell is not decoded.
//
// Anything else (`0+10`, functions, names, ranges, sheet-qualified
// references) is not decodable, and callers treat the rule as not matching:
// it neither formats nor stops. `decodeCfOperandResult` keeps the reasons
// apart: an operand outside this grammar (or a relative reference shifted
// past the grid edge) is outside this decoder's admitted domain and the
// renderer reports it (#1547); a cached Excel error is a normal no-match.
// ────────────────────────────────────────────────────────────────

/** A decoded operand; `null` is a blank cell. */
export type CfOperandValue = number | string | boolean | null;

export interface CfOperandContext {
  row: number;
  col: number;
  anchorRow: number;
  anchorCol: number;
  cellIndex: ReadonlyMap<string, Cell>;
}

/** Tagged decode. `unsupported` reasons: `syntax` (not a literal or a
 *  single in-grid A1 reference), `outOfGrid` (a relative reference shifted
 *  past the sheet edge), `unresolvedValue` (an unresolved shared-string
 *  reference reached the decoder). `error` is a cached Excel error value. */
export type CfOperandDecode =
  | { readonly kind: 'value'; readonly value: CfOperandValue }
  | { readonly kind: 'unsupported'; readonly reason: 'syntax' | 'outOfGrid' | 'unresolvedValue' }
  | { readonly kind: 'error' };

const UNSUPPORTED_SYNTAX = Object.freeze({ kind: 'unsupported', reason: 'syntax' } as const);
const UNSUPPORTED_OUT_OF_GRID = Object.freeze({ kind: 'unsupported', reason: 'outOfGrid' } as const);
const UNSUPPORTED_VALUE = Object.freeze({ kind: 'unsupported', reason: 'unresolvedValue' } as const);
const CACHED_ERROR = Object.freeze({ kind: 'error' } as const);
const BLANK = Object.freeze({ kind: 'value', value: null } as const);

const NUMERIC_LITERAL = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;
const STRING_LITERAL = /^"(?:[^"]|"")*"$/;
const CELL_REFERENCE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/;
const MAX_ROW = 1_048_576;
const MAX_COL = 16_384;

/** The literal's value, or `undefined` when `formula` is not a plain
 *  numeric or string literal. */
export function decodeCfLiteral(formula: string): number | string | undefined {
  const t = formula.trim();
  if (STRING_LITERAL.test(t)) return t.slice(1, -1).replace(/""/g, '"');
  if (NUMERIC_LITERAL.test(t)) {
    const n = Number(t);
    return Number.isFinite(n) ? n : undefined;
  }
  return undefined;
}

/** Decode an operand into a tagged result (see {@link CfOperandDecode}). */
export function decodeCfOperandResult(formula: string, ctx: CfOperandContext): CfOperandDecode {
  const literal = decodeCfLiteral(formula);
  if (literal !== undefined) return { kind: 'value', value: literal };
  const m = CELL_REFERENCE.exec(formula.trim());
  if (!m) return UNSUPPORTED_SYNTAX;
  let refCol = 0;
  for (const ch of m[2].toUpperCase()) refCol = refCol * 26 + (ch.charCodeAt(0) - 64);
  const refRow = Number(m[4]);
  // Outside the grid the spelling is a name, not a cell (not decoded).
  if (refRow < 1 || refRow > MAX_ROW || refCol > MAX_COL) return UNSUPPORTED_SYNTAX;
  const row = m[3] === '$' ? refRow : refRow + (ctx.row - ctx.anchorRow);
  const col = m[1] === '$' ? refCol : refCol + (ctx.col - ctx.anchorCol);
  // This decoder admits only references whose shifted coordinates stay in
  // the worksheet grid. No Office wrap/error behaviour is inferred here.
  if (row < 1 || row > MAX_ROW || col < 1 || col > MAX_COL) return UNSUPPORTED_OUT_OF_GRID;
  const value = ctx.cellIndex.get(`${row}:${col}`)?.value;
  if (!value) return BLANK;
  switch (value.type) {
    case 'empty': return BLANK;
    case 'number': return { kind: 'value', value: value.number };
    case 'text': return { kind: 'value', value: value.text };
    case 'bool': return { kind: 'value', value: value.bool };
    case 'error': return CACHED_ERROR;
    default: return UNSUPPORTED_VALUE;
  }
}

/** Compatibility wrapper: the operand's value, or `undefined` when it is
 *  not decodable (any non-value tag). */
export function decodeCfOperand(formula: string, ctx: CfOperandContext): CfOperandValue | undefined {
  const result = decodeCfOperandResult(formula, ctx);
  return result.kind === 'value' ? result.value : undefined;
}

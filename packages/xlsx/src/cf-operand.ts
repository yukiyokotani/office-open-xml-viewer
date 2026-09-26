import type { Cell } from './types.js';

// ────────────────────────────────────────────────────────────────
// Conditional-formatting operand decoding (no formula evaluation)
//
// The operands of a `cellIs` rule and the activity condition of a
// colorScale / dataBar / iconSet rule ([MS-XLSX] 2.6.27) are formulas. A
// guessed value could stop lower-priority rules (ECMA-376 §18.3.1.10
// stopIfTrue), and this library does not evaluate formulas: it reads cached
// values. So an operand is decoded in exactly two cases:
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
// it neither formats nor stops.
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

/** The operand's value, or `undefined` when it is neither a literal nor a
 *  single-cell reference to a cached value. */
export function decodeCfOperand(formula: string, ctx: CfOperandContext): CfOperandValue | undefined {
  const literal = decodeCfLiteral(formula);
  if (literal !== undefined) return literal;
  const m = CELL_REFERENCE.exec(formula.trim());
  if (!m) return undefined;
  let refCol = 0;
  for (const ch of m[2].toUpperCase()) refCol = refCol * 26 + (ch.charCodeAt(0) - 64);
  const refRow = Number(m[4]);
  if (refRow < 1 || refRow > MAX_ROW || refCol > MAX_COL) return undefined;
  const row = m[3] === '$' ? refRow : refRow + (ctx.row - ctx.anchorRow);
  const col = m[1] === '$' ? refCol : refCol + (ctx.col - ctx.anchorCol);
  // Excel wraps a relative reference shifted past the sheet edge; not
  // decoded.
  if (row < 1 || row > MAX_ROW || col < 1 || col > MAX_COL) return undefined;
  const value = ctx.cellIndex.get(`${row}:${col}`)?.value;
  if (!value) return null;
  switch (value.type) {
    case 'empty': return null;
    case 'number': return value.number;
    case 'text': return value.text;
    case 'bool': return value.bool;
    default: return undefined;
  }
}

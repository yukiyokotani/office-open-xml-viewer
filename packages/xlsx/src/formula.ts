import { excelSerialToUtcDate, utcDateToExcelSerial } from '@silurus/ooxml-core';
import type { Cell, DefinedName } from './types.js';
// OOXML's fixed address grammar bounds belong to lexical reference parsing,
// independently of renderer geometry and its shared layout runtime.
const MAX_WORKSHEET_ROW = 1_048_576;
const MAX_WORKSHEET_COL = 16_384;

// ────────────────────────────────────────────────────────────────
// Formula evaluator (conditional-formatting `expression` rules)
//
// Handles the narrow subset of Excel formulas used by CF expression rules:
// numeric/boolean literals, cell references (A1-style, with $ absolute
// markers), defined-name resolution, comparison/arithmetic operators, error
// values, and the functions listed in `FUNCTION_ARITY`. An expression is
// parsed completely before it is evaluated against cached cell values.
// Formula strings embed relative references that shift based on
// the evaluation cell's offset from an anchor cell:
//   - CF formulas use the top-left of the rule's `sqref` as anchor
//   - Workbook-level defined names are anchored at A1 (row 1, col 1)
// Column letters outside the defined-name anchor case are also shifted by
// the (col - anchorCol) delta; rows similarly. `$` markers pin the coord.
// ────────────────────────────────────────────────────────────────

interface EvalCtx {
  row: number;
  col: number;
  anchorRow: number;
  anchorCol: number;
  cellIndex: ReadonlyMap<string, Cell>;
  /** In-scope names keyed by spelling; see `lookupDefinedName`. */
  definedNames: ReadonlyMap<string, DefinedName>;
  /** Recursion guard for nested defined-name resolution. */
  depth: number;
}

/** An ECMA-376 §18.17 error value (`#DIV/0!`, `#N/A`, …), from an error
 *  literal, a cell's cached error or a function/operator domain failure. It
 *  is a value, so IFERROR and the IS* functions can inspect it; every other
 *  consumer coerces through `toNum`/`toBool`/`toStr`, which propagate it. */
class FormulaErrorValue {
  constructor(readonly code: string) {}
}

type EvalScalar = number | boolean | string | null | FormulaErrorValue;
type EvalValue = EvalScalar | EvalScalar[];
type PlainScalar = Exclude<EvalScalar, FormulaErrorValue>;

type FailureKind = 'unsupported' | 'invalid' | 'error';

/** Internal CF boundary, not a claim of complete formula support (#1547).
 * ECMA-376 §18.17.2 requires the whole expression to be consumed. Unsupported
 * syntax/functions/names and invalid expressions are never numeric placeholders:
 * NOT, comparison or arithmetic could turn 0 into TRUE and make stopIfTrue
 * suppress valid lower-priority rules (§18.3.1.10). An expression whose value
 * is an error value is `error`: a CF formula applies only when it is TRUE.
 */
export type FormulaEvaluation =
  | { kind: 'value'; value: EvalValue }
  | { kind: FailureKind };

class FormulaFailure extends Error {
  constructor(readonly kind: FailureKind, readonly code?: string) {
    super(`CF formula ${kind}`);
  }
}

/** Flatten nested scalars and arrays to a flat list of scalars. */
function flatten(v: EvalValue): EvalScalar[] {
  return Array.isArray(v) ? v : [v];
}

/** Unwrap an array value to its first scalar element (Excel's intersection
 *  behavior is not modeled; we collapse ranges to the first cell when a
 *  scalar is required). */
function toScalar(v: EvalValue): EvalScalar {
  return Array.isArray(v) ? (v[0] ?? 0) : v;
}

/** Scalar value of an operand, propagating an error value (the operators
 *  and the functions modeled here return the first error operand). */
function operand(v: EvalValue): PlainScalar {
  const s = toScalar(v);
  if (s instanceof FormulaErrorValue) throw new FormulaFailure('error', s.code);
  return s;
}

/** SUM/MIN/MAX/AVERAGE/AND/OR propagate an error anywhere in their
 *  arguments, including inside a referenced range. */
function flattenPropagatingErrors(args: EvalValue[]): PlainScalar[] {
  return args.flatMap(flatten).map(v => operand(v));
}

const MAX_DEFINED_NAME_DEPTH = 8;
// Library admission policy, not an Excel grammar limit. Bound recursive
// descent before an untrusted CF expression can exhaust the JS call stack.
// The budget covers grouping, unary operators and defined-name expansion,
// and therefore also the depth of the tree that `evaluate` walks.
const MAX_FORMULA_PARSE_DEPTH = 128;
// Library resource policy: retain at most 64 Ki UTF-16 source units from
// expanded defined-name bodies per CF expression. A depth guard alone does
// not bound a branching name graph's exponentially duplicated AST. Debit
// the shared budget before tokenizing each body; reject the entire expression
// instead of truncating its meaning. This does not cap the outer formula.
const MAX_EXPANDED_NAME_SOURCE_UNITS = 64 * 1024;
interface ParseBudget {
  depth: number;
  remainingNameSourceUnits: number;
}

export function evalFormulaToBool(formula: string, ctx: EvalCtx): boolean {
  const result = evaluateFormula(formula, ctx);
  if (result.kind !== 'value') return false;
  try {
    return toBool(result.value);
  } catch (error) {
    // A range whose first cell is an error value.
    if (error instanceof FormulaFailure) return false;
    throw error;
  }
}

export function evaluateFormula(formula: string, ctx: EvalCtx): FormulaEvaluation {
  try {
    const value = evalFormula(formula, ctx);
    if (value instanceof FormulaErrorValue) return { kind: 'error' };
    return { kind: 'value', value };
  } catch (error) {
    if (error instanceof FormulaFailure) return { kind: error.kind };
    // Do not silently misclassify a programming/resource failure as no match.
    throw error;
  }
}

function toBool(v: EvalValue): boolean {
  const s = operand(v);
  if (typeof s === 'boolean') return s;
  if (typeof s === 'number') return s !== 0;
  if (typeof s === 'string') return s.length > 0 && s.toUpperCase() !== 'FALSE';
  return false;
}

/** Number operand for arithmetic, unary operators and number-typed function
 * arguments (§18.17.7: an incompatible argument type is #VALUE!). Each failure
 * is an Excel error value, never a placeholder 0: NOT, a comparison or
 * stopIfTrue could otherwise turn the error into a match. Comparison operands
 * do not use this conversion; see applyCmp. */
function toNum(v: EvalValue): number {
  const s = operand(v);
  if (typeof s === 'number') {
    // §18.17.2.6 numbers are real numbers. An IEEE Infinity or NaN can only
    // come from an overflowed or undefined intermediate result, which §18.17.3
    // classifies as #NUM!. Operator results are not range-checked themselves.
    if (!Number.isFinite(s)) throw new FormulaFailure('error', '#NUM!');
    return s;
  }
  if (typeof s === 'boolean') return s ? 1 : 0;
  if (s == null) return 0;
  // §18.17.3 #VALUE!: text is an incompatible operand unless converted.
  const n = parseNumericText(s);
  if (n === null) throw new FormulaFailure('error', '#VALUE!');
  return n;
}

/** Library policy, not an Excel rule: §18.17.2.6 leaves text-to-number
 * conversion implementation-defined. Only the locale-independent §18.17.2.1
 * numeric-constant spelling converts, with an optional sign and surrounding
 * U+0020 spaces; underflow reads as 0, as the literal tokenizer does. Other
 * spellings Excel may convert through locale or number formats (percent,
 * currency, grouping, dates, times) fail as #VALUE!, as does a numeric
 * prefix ("1x") or overflow, rather than as a guess. */
function parseNumericText(s: string): number | null {
  if (!/^ *[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)? *$/u.test(s)) return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

function toStr(v: EvalValue): string {
  const s = operand(v);
  if (s == null) return '';
  if (typeof s === 'boolean') return s ? 'TRUE' : 'FALSE';
  return String(s);
}

function formulaError(code: string): FormulaErrorValue {
  return new FormulaErrorValue(code);
}

interface Tok {
  kind: 'num' | 'str' | 'op' | 'lparen' | 'rparen' | 'comma' | 'ref' | 'name' | 'bool' | 'colon' | 'error';
  text: string;
  /** For 'ref': pre-parsed reference. */
  ref?: { colAbs: boolean; col: number; rowAbs: boolean; row: number };
}

const OP_CHARS = new Set(['<', '>', '=', '+', '-', '*', '/', '&', '^', '%']);

function tokenize(formula: string): Tok[] {
  const toks: Tok[] = [];
  let i = 0;
  const s = formula;
  while (i < s.length) {
    const c = s[i];
    if (c === ' ' || c === '\t' || c === '\n' || c === '\r') { i++; continue; }
    if (c === '(') { toks.push({ kind: 'lparen', text: c }); i++; continue; }
    if (c === ')') { toks.push({ kind: 'rparen', text: c }); i++; continue; }
    if (c === ',') { toks.push({ kind: 'comma', text: c }); i++; continue; }
    if (c === ':') { toks.push({ kind: 'colon', text: c }); i++; continue; }
    if (c === '"') {
      let j = i + 1; let buf = '';
      while (j < s.length) {
        if (s[j] === '"' && s[j + 1] === '"') { buf += '"'; j += 2; continue; }
        if (s[j] === '"') break;
        buf += s[j]; j++;
      }
      if (j === s.length) throw new FormulaFailure('invalid');
      toks.push({ kind: 'str', text: buf });
      i = j + 1;
      continue;
    }
    if (c === '#') {
      // ECMA-376 §18.17.2 error literals (e.g. `#REF!` left by a deleted
      // reference). Skipping the `#` would read `REF` as an unknown name
      // worth 0 and let the rule match.
      const m = /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA)/u.exec(s.slice(i));
      if (!m) throw new FormulaFailure('unsupported');
      toks.push({ kind: 'error', text: m[0] });
      i += m[0].length;
      continue;
    }
    if ((c >= '0' && c <= '9') || (c === '.' && /[0-9]/u.test(s[i + 1] ?? ''))) {
      // §18.17.2.1 numerical-constant: retain decimal/exponent spellings
      // while rejecting adjacent numbers rather than accepting a prefix.
      const m = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/u.exec(s.slice(i));
      if (!m || !Number.isFinite(Number(m[0]))) throw new FormulaFailure('invalid');
      toks.push({ kind: 'num', text: m[0] });
      i += m[0].length;
      continue;
    }
    if (OP_CHARS.has(c)) {
      // Multi-char operators: <=, >=, <>
      if ((c === '<' || c === '>') && (s[i + 1] === '=' || (c === '<' && s[i + 1] === '>'))) {
        toks.push({ kind: 'op', text: s.slice(i, i + 2) });
        i += 2;
      } else {
        toks.push({ kind: 'op', text: c });
        i++;
      }
      continue;
    }
    // Reference or identifier: may start with $, letters, or letters+digits.
    // Defined names allow letters, digits, '_', '.'; cell refs are
    // `$?[A-Z]+\$?[0-9]+` (case-insensitive) inside the worksheet grid. A
    // name cannot be spelled as a cell reference (§18.2.5, §18.17.5.1), so a
    // spelling inside the grid is a cell and one outside it may be a name.
    // A cell-like spelling directly before '(' is a function name (LOG10).
    if (c === '$' || isIdentStart(c)) {
      let j = i;
      while (j < s.length && (s[j] === '$' || isIdentPart(s[j]))) j++;
      const text = s.slice(i, j);
      i = j;
      const ref = s[i] === '(' ? null : tryParseCellRef(text);
      if (ref) {
        toks.push({ kind: 'ref', text, ref });
      } else {
        const up = text.toUpperCase();
        if ((up === 'TRUE' || up === 'FALSE') && s[i] !== '(') toks.push({ kind: 'bool', text: up });
        else toks.push({ kind: 'name', text });
      }
      continue;
    }
    // Sheet-qualified references, arrays and future syntax are outside this
    // evaluator. Skipping a token would evaluate a different expression.
    throw new FormulaFailure('unsupported');
  }
  return toks;
}

function isIdentStart(c: string): boolean {
  return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c === '_';
}

function isIdentPart(c: string): boolean {
  return isIdentStart(c) || (c >= '0' && c <= '9') || c === '.';
}

function tryParseCellRef(s: string): { colAbs: boolean; col: number; rowAbs: boolean; row: number } | null {
  // $?[A-Z]+\$?[0-9]+
  let i = 0;
  let colAbs = false, rowAbs = false;
  if (s[i] === '$') { colAbs = true; i++; }
  const colStart = i;
  while (i < s.length && s[i] >= 'A' && s[i].toUpperCase() <= 'Z') {
    if (!(s[i] >= 'A' && s[i] <= 'Z') && !(s[i] >= 'a' && s[i] <= 'z')) break;
    i++;
  }
  if (i === colStart) return null;
  const colLetters = s.slice(colStart, i).toUpperCase();
  if (s[i] === '$') { rowAbs = true; i++; }
  const rowStart = i;
  while (i < s.length && s[i] >= '0' && s[i] <= '9') i++;
  if (i === rowStart) return null;
  if (i !== s.length) return null;
  const rowNum = parseInt(s.slice(rowStart, i), 10);
  let col = 0;
  for (let k = 0; k < colLetters.length; k++) {
    col = col * 26 + (colLetters.charCodeAt(k) - 64);
  }
  // The worksheet grid (MAX_WORKSHEET_ROW x MAX_WORKSHEET_COL) bounds cell
  // references. Outside it the spelling is not a cell (§18.2.5, §18.17.5.1),
  // so `Limit1` resolves as a defined name instead of reading a blank that
  // comparisons would treat as 0.
  if (rowNum < 1 || rowNum > MAX_WORKSHEET_ROW || col > MAX_WORKSHEET_COL) return null;
  return { colAbs, col, rowAbs, row: rowNum };
}

type CellRef = { colAbs: boolean; col: number; rowAbs: boolean; row: number };

// The whole expression is parsed into a tree before anything is evaluated, so
// IF evaluates only the selected branch and IFERROR/IS* observe an error
// value produced anywhere inside their argument (see `callFunc`).
type FormulaNode =
  | { t: 'lit'; v: EvalScalar }
  | { t: 'ref'; ref: CellRef }
  | { t: 'range'; a: CellRef; b: CellRef }
  | { t: 'neg' | 'pos'; e: FormulaNode }
  /** A comparison (not chained: a second comparison operator is unsupported). */
  | { t: 'bin'; op: string; l: FormulaNode; r: FormulaNode }
  /** A left-associative `&`, `+`/`-` or `*`/`/` run, kept flat so evaluation
   *  loops over it instead of recursing once per operator. */
  | { t: 'chain'; first: FormulaNode; rest: { op: string; r: FormulaNode }[] }
  | { t: 'call'; name: string; args: FormulaNode[] }
  /** A workbook defined name; its body anchors relative references at A1. */
  | { t: 'name'; body: FormulaNode };

interface Parser {
  toks: Tok[];
  pos: number;
  budget: ParseBudget;
  names: ReadonlyMap<string, DefinedName>;
  nameDepth: number;
}

// §18.17.2.5: defined names are case-insensitive. The caller keys its map by
// the stored spelling, so resolve through ASCII-folded keys. Folding only
// ASCII is library policy: formula tokens are ASCII here, and Unicode case
// equivalence (e.g. KELVIN SIGN vs "k") is not established. A later case
// variant wins, as an exact duplicate already does in the caller's map and
// in internal-hyperlink name resolution (the existing shadowing policy).
//
// The index is built once per caller-owned map, not per evaluated cell. The
// map is treated as fixed after construction; a size change rebuilds the
// index, and a stale key whose definition was removed fails closed.
interface FoldedNameIndex {
  size: number;
  keys: Map<string, string>;
}
const foldedNameIndexes = new WeakMap<ReadonlyMap<string, DefinedName>, FoldedNameIndex>();

function asciiFold(name: string): string {
  return name.replace(/[A-Z]+/gu, (letters) => letters.toLowerCase());
}

function lookupDefinedName(names: ReadonlyMap<string, DefinedName>, spelling: string): DefinedName | undefined {
  let index = foldedNameIndexes.get(names);
  if (!index || index.size !== names.size) {
    const keys = new Map<string, string>();
    for (const key of names.keys()) keys.set(asciiFold(key), key);
    index = { size: names.size, keys };
    foldedNameIndexes.set(names, index);
  }
  const key = index.keys.get(asciiFold(spelling));
  return key === undefined ? undefined : names.get(key);
}

function evalFormula(formula: string, ctx: EvalCtx): EvalValue {
  return evaluate(parseFormula(formula, ctx.definedNames, ctx.depth,
    { depth: 0, remainingNameSourceUnits: MAX_EXPANDED_NAME_SOURCE_UNITS }), ctx);
}

function parseFormula(
  formula: string,
  names: ReadonlyMap<string, DefinedName>,
  nameDepth: number,
  budget: ParseBudget,
): FormulaNode {
  // The stored grammar is an expression; accept the conventional display '='
  // prefix too, without interpreting a second '=' as an empty operand.
  const source = formula.trim();
  const toks = tokenize(source.startsWith('=') ? source.slice(1) : source);
  const p: Parser = { toks, pos: 0, budget, names, nameDepth };
  const node = parseExpr(p);
  if (p.pos !== toks.length) throw new FormulaFailure('unsupported');
  return node;
}

function peek(p: Parser): Tok | undefined { return p.toks[p.pos]; }
function consume(p: Parser): Tok | undefined { return p.toks[p.pos++]; }

function parseExpr(p: Parser): FormulaNode {
  if (p.budget.depth >= MAX_FORMULA_PARSE_DEPTH) throw new FormulaFailure('unsupported');
  p.budget.depth++;
  try {
    return parseCmp(p);
  } finally {
    p.budget.depth--;
  }
}

function parseCmp(p: Parser): FormulaNode {
  const left = parseConcat(p);
  const t = peek(p);
  if (t && t.kind === 'op' && (t.text === '<' || t.text === '>' || t.text === '<=' || t.text === '>=' || t.text === '=' || t.text === '<>')) {
    consume(p);
    return { t: 'bin', op: t.text, l: left, r: parseConcat(p) };
  }
  return left;
}

function parseChain(p: Parser, ops: readonly string[], parseOperand: (p: Parser) => FormulaNode): FormulaNode {
  const first = parseOperand(p);
  const rest: { op: string; r: FormulaNode }[] = [];
  while (true) {
    const t = peek(p);
    if (!t || t.kind !== 'op' || !ops.includes(t.text)) break;
    consume(p);
    rest.push({ op: t.text, r: parseOperand(p) });
  }
  return rest.length ? { t: 'chain', first, rest } : first;
}

function parseConcat(p: Parser): FormulaNode { return parseChain(p, ['&'], parseAdd); }
function parseAdd(p: Parser): FormulaNode { return parseChain(p, ['+', '-'], parseMul); }
function parseMul(p: Parser): FormulaNode { return parseChain(p, ['*', '/'], parseUnary); }

function applyBinary(op: string, a: EvalValue, b: EvalValue): EvalValue {
  // Operands are coerced left to right, so the left operand's error wins.
  switch (op) {
    case '&': return toStr(a) + toStr(b);
    case '+': return toNum(a) + toNum(b);
    case '-': return toNum(a) - toNum(b);
    case '*': return toNum(a) * toNum(b);
    case '/': {
      const n = toNum(a);
      const d = toNum(b);
      // ECMA-376 §18.17 error values: #DIV/0! is the result of dividing by
      // zero (a blank divisor reads as 0), never a numeric placeholder.
      return d === 0 ? formulaError('#DIV/0!') : n / d;
    }
    default:  return applyCmp(op, operand(a), operand(b));
  }
}

/** Unsettled (#1547): §18.17 does not define how comparison operators convert
 * or order mixed-type, blank or text operands, and no Excel evidence is
 * recorded here. Until it is, this keeps the evaluator's earlier rule rather
 * than adopting either arithmetic's #VALUE! conversion or a guessed Excel
 * order: when neither operand is text without a numeric prefix, both compare
 * as numbers; otherwise both compare as case-sensitive strings, a blank as "".
 * Known unverified consequences include "1"=1, TRUE=1 and "2024 Q1"="2024 Q2"
 * reading TRUE. */
function applyCmp(op: string, a: PlainScalar, b: PlainScalar): boolean {
  const an = typeof a === 'string' && isNaN(parseFloat(a)) ? null : comparisonNumber(a);
  const bn = typeof b === 'string' && isNaN(parseFloat(b)) ? null : comparisonNumber(b);
  if (an !== null && bn !== null) {
    switch (op) {
      case '<':  return an <  bn;
      case '>':  return an >  bn;
      case '<=': return an <= bn;
      case '>=': return an >= bn;
      case '=':  return an === bn;
      case '<>': return an !== bn;
    }
  }
  const sa = String(a ?? ''); const sb = String(b ?? '');
  switch (op) {
    case '<':  return sa <  sb;
    case '>':  return sa >  sb;
    case '<=': return sa <= sb;
    case '>=': return sa >= sb;
    case '=':  return sa === sb;
    case '<>': return sa !== sb;
  }
  return false;
}

/** The evaluator's earlier lenient number reading, kept only for applyCmp's
 * unsettled rule: a text prefix ("1x" → 1), 0 for other text, and no error.
 * Do not use it for arithmetic or function arguments; they use toNum. */
function comparisonNumber(v: EvalValue): number {
  const s = operand(v);
  if (typeof s === 'number') return s;
  if (typeof s === 'boolean') return s ? 1 : 0;
  if (s == null) return 0;
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}


function parseUnary(p: Parser): FormulaNode {
  if (p.budget.depth >= MAX_FORMULA_PARSE_DEPTH) throw new FormulaFailure('unsupported');
  p.budget.depth++;
  try {
    const t = peek(p);
    if (t && t.kind === 'op' && t.text === '-') { consume(p); return { t: 'neg', e: parseUnary(p) }; }
    if (t && t.kind === 'op' && t.text === '+') { consume(p); return { t: 'pos', e: parseUnary(p) }; }
    return parsePrimary(p);
  } finally {
    p.budget.depth--;
  }
}

function parsePrimary(p: Parser): FormulaNode {
  const t = consume(p);
  if (!t) throw new FormulaFailure('invalid');
  if (t.kind === 'num') return { t: 'lit', v: parseFloat(t.text) };
  if (t.kind === 'str') return { t: 'lit', v: t.text };
  if (t.kind === 'bool') return { t: 'lit', v: t.text === 'TRUE' };
  if (t.kind === 'error') return { t: 'lit', v: formulaError(t.text) };
  if (t.kind === 'lparen') {
    const v = parseExpr(p);
    const next = consume(p);
    if (!next || next.kind !== 'rparen') throw new FormulaFailure('invalid');
    return v;
  }
  if (t.kind === 'ref') {
    // Range: `A1:B5` — resolve as array of cell values.
    if (peek(p)?.kind === 'colon') {
      consume(p);
      const right = consume(p);
      if (right?.kind !== 'ref' || !right.ref) throw new FormulaFailure('invalid');
      return { t: 'range', a: t.ref!, b: right.ref };
    }
    return { t: 'ref', ref: t.ref! };
  }
  if (t.kind === 'name') {
    // Function call: NAME(args)
    if (peek(p)?.kind === 'lparen') {
      consume(p);
      const args: FormulaNode[] = [];
      if (peek(p)?.kind !== 'rparen') {
        args.push(parseExpr(p));
        while (peek(p)?.kind === 'comma') {
          consume(p);
          args.push(parseExpr(p));
        }
      }
      const next = consume(p);
      if (!next || next.kind !== 'rparen') throw new FormulaFailure('invalid');
      const name = t.text.toUpperCase();
      checkArity(name, args.length);
      return { t: 'call', name, args };
    }
    // Library admission is whole-expression, independently of IF's lazy
    // value evaluation: an unsupported name/body cannot be hidden in the
    // branch not taken and then incorrectly suppress another CF rule.
    const dn = lookupDefinedName(p.names, t.text);
    if (!dn || p.nameDepth >= MAX_DEFINED_NAME_DEPTH) throw new FormulaFailure('unsupported');
    if (dn.formula.length > p.budget.remainingNameSourceUnits) throw new FormulaFailure('unsupported');
    p.budget.remainingNameSourceUnits -= dn.formula.length;
    // Strip `SheetName!` prefix if present; keep just the ref body.
    return { t: 'name', body: parseFormula(stripSheetPrefix(dn.formula), p.names, p.nameDepth + 1, p.budget) };
  }
  throw new FormulaFailure('invalid');
}

function evaluate(n: FormulaNode, ctx: EvalCtx): EvalValue {
  try {
    return evaluateNode(n, ctx);
  } catch (error) {
    // Excel errors are values at every expression boundary, including a
    // compound operator expression. ISNUMBER/COUNT must inspect them just
    // as they inspect a cached error cell. Only operand coercion propagates
    // them; library admission and programming failures remain exceptions.
    if (error instanceof FormulaFailure && error.kind === 'error') {
      return formulaError(error.code ?? '#VALUE!');
    }
    throw error;
  }
}

function evaluateNode(n: FormulaNode, ctx: EvalCtx): EvalValue {
  switch (n.t) {
    case 'lit': return n.v;
    case 'ref': return resolveRef(n.ref, ctx);
    case 'range': return resolveRange(n.a, n.b, ctx);
    case 'neg': return -toNum(evaluate(n.e, ctx));
    case 'pos': return toNum(evaluate(n.e, ctx));
    case 'bin': return applyBinary(n.op, evaluate(n.l, ctx), evaluate(n.r, ctx));
    case 'chain': {
      let acc = evaluate(n.first, ctx);
      for (const { op, r } of n.rest) acc = applyBinary(op, acc, evaluate(r, ctx));
      return acc;
    }
    case 'call': return callFunc(n.name, n.args, ctx);
    // Workbook-level defined names anchor at A1 for relative-ref shifts.
    case 'name': return evaluate(n.body, { ...ctx, anchorRow: 1, anchorCol: 1 });
  }
}

// Argument counts of the supported functions, from their signatures in
// ECMA-376 Part 1 §18.17.7 (IFS and CONCAT are later Excel functions with the
// same documented form). A call outside them is not a valid formula, so it is
// invalid rather than evaluated with a missing argument read as blank (which
// made NOT() TRUE). The 255-argument ceiling is library admission policy,
// not a normative maximum: §18.17.2 encourages support for at least 255.
// IF requires its logical test and value-if-true slot (§18.17.7.147);
// empty argument slots remain outside this evaluator's grammar. Names not
// listed are not implemented and fail as
// unsupported during whole-expression admission, including an unused IF branch.
const FUNCTION_ARITY: Readonly<Record<string, readonly [min: number, max: number]>> = {
  AND: [1, 255], OR: [1, 255], NOT: [1, 1], IF: [2, 3], IFERROR: [2, 2], IFS: [2, 254],
  TRUE: [0, 0], FALSE: [0, 0],
  ISBLANK: [1, 1], ISNUMBER: [1, 1], ISTEXT: [1, 1], ISNONTEXT: [1, 1],
  ISERROR: [1, 1], ISERR: [1, 1], ISNA: [1, 1], ISLOGICAL: [1, 1],
  ROUNDDOWN: [2, 2], ROUNDUP: [2, 2], ROUND: [2, 2], INT: [1, 1], TRUNC: [1, 2],
  CEILING: [2, 2], FLOOR: [2, 2], MOD: [2, 2], POWER: [2, 2], SQRT: [1, 1],
  ABS: [1, 1], SIGN: [1, 1], EXP: [1, 1], LN: [1, 1], LOG10: [1, 1],
  MIN: [1, 255], MAX: [1, 255], SUM: [1, 255], AVERAGE: [1, 255],
  COUNT: [1, 255], COUNTA: [1, 255], COUNTBLANK: [1, 1],
  COUNTIF: [2, 2], SUMIF: [2, 3], AVERAGEIF: [2, 3],
  LEN: [1, 1], LEFT: [1, 2], RIGHT: [1, 2], MID: [3, 3], UPPER: [1, 1], LOWER: [1, 1],
  TRIM: [1, 1], EXACT: [2, 2], FIND: [2, 3], SEARCH: [2, 3],
  CONCATENATE: [1, 255],
  // Later Excel CONCAT syntax admits at most 253 text arguments:
  // https://support.microsoft.com/en-us/excel/functions/concat-function
  CONCAT: [1, 253], T: [1, 1], N: [1, 1], VALUE: [1, 1],
  ROW: [0, 1], COLUMN: [0, 1],
  TODAY: [0, 0], NOW: [0, 0], DATE: [3, 3], YEAR: [1, 1], MONTH: [1, 1], DAY: [1, 1],
  WEEKDAY: [1, 2],
};

function checkArity(name: string, count: number): void {
  const arity = FUNCTION_ARITY[name];
  if (!arity) throw new FormulaFailure('unsupported');
  // IFS takes condition/value pairs.
  if (count < arity[0] || count > arity[1] || (name === 'IFS' && count % 2 !== 0)) {
    throw new FormulaFailure('invalid');
  }
}

function stripSheetPrefix(formula: string): string {
  // Match `'Sheet Name'!ref` or `SheetName!ref`. Only the leading reference
  // prefix is stripped; we don't need cross-sheet lookups because defined
  // names here point to cells on the active sheet.
  const m = formula.match(/^(?:'[^']*'|[A-Za-z_][A-Za-z0-9_.]*)!(.*)$/);
  return m ? m[1] : formula;
}

/** Sheet coordinates of a reference evaluated at `ctx`'s cell. */
function refCoord(ref: CellRef, ctx: EvalCtx): { row: number; col: number } {
  return {
    col: ref.colAbs ? ref.col : ref.col + (ctx.col - ctx.anchorCol),
    row: ref.rowAbs ? ref.row : ref.row + (ctx.row - ctx.anchorRow),
  };
}

function resolveRef(ref: CellRef, ctx: EvalCtx): EvalScalar {
  const { row, col } = refCoord(ref, ctx);
  return cellValueToEval(ctx.cellIndex.get(`${row}:${col}`));
}

// Library resource policy, not an Excel limit: bound the per-cell work of a
// pathological range. A larger range is unsupported rather than truncated,
// because a partial range silently changes COUNTIF/SUM/... results.
const MAX_RANGE_CELLS = 4096;

interface CellRect { r1: number; c1: number; r2: number; c2: number }

function rangeRect(a: CellRef, b: CellRef, ctx: EvalCtx): CellRect {
  const pa = refCoord(a, ctx);
  const pb = refCoord(b, ctx);
  return {
    r1: Math.min(pa.row, pb.row), c1: Math.min(pa.col, pb.col),
    r2: Math.max(pa.row, pb.row), c2: Math.max(pa.col, pb.col),
  };
}

/** Cached values of `rect` in row-major order. */
function readRect({ r1, c1, r2, c2 }: CellRect, ctx: EvalCtx): EvalScalar[] {
  if ((r2 - r1 + 1) * (c2 - c1 + 1) > MAX_RANGE_CELLS) throw new FormulaFailure('unsupported');
  const out: EvalScalar[] = [];
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      out.push(cellValueToEval(ctx.cellIndex.get(`${r}:${c}`)));
    }
  }
  return out;
}

function resolveRange(a: CellRef, b: CellRef, ctx: EvalCtx): EvalScalar[] {
  return readRect(rangeRect(a, b, ctx), ctx);
}

/** The cell rectangle a reference node denotes at `ctx`, or null for any
 *  other expression. A defined name's body anchors at A1, as in `evaluate`. */
function referenceRect(n: FormulaNode, ctx: EvalCtx): CellRect | null {
  switch (n.t) {
    case 'ref': { const { row, col } = refCoord(n.ref, ctx); return { r1: row, c1: col, r2: row, c2: col }; }
    case 'range': return rangeRect(n.a, n.b, ctx);
    case 'name': return referenceRect(n.body, { ...ctx, anchorRow: 1, anchorCol: 1 });
    default: return null;
  }
}

/** SUMIF sum_range / AVERAGEIF average_range. Microsoft's SUMIF and
 * AVERAGEIF documentation: the supplied range need not match `range` in size
 * or shape; the cells used start at its top-left cell and take `range`'s
 * dimensions (A1:B4 with C1:C2 reads C1:D4). Only that rectangle is read.
 * https://support.microsoft.com/en-us/excel/functions/sumif-function
 * https://support.microsoft.com/en-us/excel/functions/averageif-function
 * Only direct references (including reference-valued defined names) retain
 * rectangle provenance here. Computed expressions keep their existing
 * evaluate/flatten pairing; this preserves IF/IFS/IFERROR selection and lazy
 * branch behavior without re-evaluating a condition to infer a rectangle.
 * Library policy, without Excel evidence: a resized rectangle outside the
 * worksheet grid is unsupported. */
function resizedTargetValues(rangeNode: FormulaNode, targetNode: FormulaNode, ctx: EvalCtx): EvalScalar[] | null {
  const range = referenceRect(rangeNode, ctx);
  const target = referenceRect(targetNode, ctx);
  if (!range || !target) return null;
  const r2 = target.r1 + (range.r2 - range.r1);
  const c2 = target.c1 + (range.c2 - range.c1);
  if (target.r1 < 1 || target.c1 < 1 || r2 > MAX_WORKSHEET_ROW || c2 > MAX_WORKSHEET_COL) {
    throw new FormulaFailure('unsupported');
  }
  return readRect({ r1: target.r1, c1: target.c1, r2, c2 }, ctx);
}

function cellValueToEval(cell: Cell | undefined): EvalScalar {
  // An empty / missing cell is *not* the same as 0. CF expressions like
  // `=$C5=0` or `NOT(ISBLANK($C5))` will match a missing cell if we return
  // 0 here, which is exactly the bug that turned C5-C8 beige on sample-10.
  // Return null and let the arithmetic / comparison operators coerce as
  // needed (null+0 → 0, null="" → true, null=0 → false). See ECMA-376
  // §18.18.62 and the actual Excel evaluation behaviour.
  if (!cell) return null;
  switch (cell.value.type) {
    case 'number': return cell.value.number;
    case 'bool':   return cell.value.bool;
    case 'text':   return cell.value.text;
    // A cached error is an error value, not a blank: `=$A2=0` must not match
    // an `#N/A` cell, and ISNA/ISERR must see its code.
    case 'error':  return formulaError(cell.value.error);
    case 'empty':
    default:       return null;
  }
}

function callFunc(name: string, argNodes: FormulaNode[], ctx: EvalCtx): EvalValue {
  const arg = (i: number) => evaluate(argNodes[i], ctx);
  // Functions that must not evaluate every argument first.
  switch (name) {
    // IF returns the selected branch only (§18.17.7.147): an error in a branch
    // not taken, e.g. the division in IF(B1=0,FALSE,A1/B1>1), does not
    // become the result. An error in a condition does.
    case 'IF':
      if (toBool(arg(0))) return arg(1);
      return argNodes.length > 2 ? arg(2) : false;
    // IFERROR and ISERROR/ISERR/ISNA test for an error value; a blank is not
    // one. ISERR excludes #N/A and ISNA accepts only #N/A.
    case 'IFERROR': {
      const v = toScalar(arg(0));
      // §18.17.7.148: an empty-cell argument is treated as empty text,
      // rather than remaining an empty-cell value for ISBLANK to observe.
      return (v instanceof FormulaErrorValue ? arg(1) : v) ?? '';
    }
    case 'ISERROR':    return toScalar(arg(0)) instanceof FormulaErrorValue;
    case 'ISERR':      { const v = toScalar(arg(0)); return v instanceof FormulaErrorValue && v.code !== '#N/A'; }
    case 'ISNA':       { const v = toScalar(arg(0)); return v instanceof FormulaErrorValue && v.code === '#N/A'; }
    // ROW()/COLUMN() are the evaluated cell; with a reference argument they
    // are that reference's row/column. Other argument forms are not modeled.
    case 'ROW':
    case 'COLUMN': {
      const node = argNodes[0];
      const at = node === undefined ? ctx : node.t === 'ref' ? refCoord(node.ref, ctx) : null;
      if (!at) throw new FormulaFailure('unsupported');
      return name === 'ROW' ? at.row : at.col;
    }
    // Direct target references are resized; computed targets retain evaluation.
    case 'SUMIF':
    case 'AVERAGEIF': {
      const source = flatten(arg(0));
      const criteria = arg(1);
      const target = argNodes.length > 2
        ? resizedTargetValues(argNodes[0], argNodes[2], ctx) ?? flatten(arg(2))
        : null;
      return name === 'SUMIF' ? sumIf(source, criteria, target) : averageIf(source, criteria, target);
    }
  }
  const args = argNodes.map(n => evaluate(n, ctx));
  switch (name) {
    // Preserve this evaluator's eager IFS argument/error behavior. IF's
    // specified lazy rule does not establish the later IFS function's
    // evaluation strategy; changing it needs separate Excel evidence.
    case 'IFS':
      flattenPropagatingErrors(args);
      for (let i = 0; i < args.length; i += 2) {
        if (toBool(args[i])) return args[i + 1];
      }
      return formulaError('#N/A');
    // ── Logic ───────────────────────────────────────────────────────────────
    case 'AND':        return flattenPropagatingErrors(args).every(a => toBool(a));
    case 'OR':         return flattenPropagatingErrors(args).some(a => toBool(a));
    case 'NOT':        return !toBool(args[0]);
    case 'TRUE':       return true;
    case 'FALSE':      return false;
    // ── Type checks ─────────────────────────────────────────────────────────
    // ISBLANK tests for an empty cell: an empty-text value is not one.
    case 'ISBLANK':    return toScalar(args[0]) === null;
    case 'ISNUMBER':   return typeof toScalar(args[0]) === 'number';
    case 'ISTEXT':     return typeof toScalar(args[0]) === 'string';
    case 'ISNONTEXT':  return typeof toScalar(args[0]) !== 'string';
    case 'ISLOGICAL':  return typeof toScalar(args[0]) === 'boolean';
    // ── Rounding / math ─────────────────────────────────────────────────────
    case 'ROUNDDOWN': {
      const n = toNum(args[0]); const d = toNum(args[1]);
      const p = Math.pow(10, d);
      return (n >= 0 ? Math.floor(n * p) : Math.ceil(n * p)) / p;
    }
    case 'ROUNDUP': {
      const n = toNum(args[0]); const d = toNum(args[1]);
      const p = Math.pow(10, d);
      return (n >= 0 ? Math.ceil(n * p) : Math.floor(n * p)) / p;
    }
    case 'ROUND':      return roundHalfAwayFromZero(toNum(args[0]), toNum(args[1]));
    case 'INT':        return Math.floor(toNum(args[0]));
    case 'TRUNC':      { const n = toNum(args[0]); const d = toNum(args[1] ?? 0); const p = Math.pow(10, d); return (n >= 0 ? Math.floor(n * p) : Math.ceil(n * p)) / p; }
    case 'CEILING':    { const n = toNum(args[0]); const sig = toNum(args[1]); return sig === 0 ? 0 : Math.ceil(n / sig) * sig; }
    case 'FLOOR':      { const n = toNum(args[0]); const sig = toNum(args[1]); return sig === 0 ? 0 : Math.floor(n / sig) * sig; }
    // Domain failures are §18.17 error values (#DIV/0!, #NUM!), not blanks
    // that a comparison would read as 0.
    case 'MOD':        { const a = toNum(args[0]); const b = toNum(args[1]); return b === 0 ? formulaError('#DIV/0!') : a - Math.floor(a / b) * b; }
    case 'POWER':      return Math.pow(toNum(args[0]), toNum(args[1]));
    case 'SQRT':       { const n = toNum(args[0]); return n < 0 ? formulaError('#NUM!') : Math.sqrt(n); }
    case 'ABS':        return Math.abs(toNum(args[0]));
    case 'SIGN':       { const n = toNum(args[0]); return n > 0 ? 1 : n < 0 ? -1 : 0; }
    case 'EXP':        return Math.exp(toNum(args[0]));
    case 'LN':         { const n = toNum(args[0]); return n <= 0 ? formulaError('#NUM!') : Math.log(n); }
    case 'LOG10':      { const n = toNum(args[0]); return n <= 0 ? formulaError('#NUM!') : Math.log10(n); }
    // ── Aggregates ──────────────────────────────────────────────────────────
    case 'MIN':        { const ns = flattenPropagatingErrors(args).filter(v => typeof v === 'number') as number[]; return ns.length ? Math.min(...ns) : 0; }
    case 'MAX':        { const ns = flattenPropagatingErrors(args).filter(v => typeof v === 'number') as number[]; return ns.length ? Math.max(...ns) : 0; }
    case 'SUM':        return flattenPropagatingErrors(args).reduce<number>((s, v) => s + (typeof v === 'number' ? v : 0), 0);
    case 'AVERAGE':    { const ns = flattenPropagatingErrors(args).filter(v => typeof v === 'number') as number[]; return ns.length ? ns.reduce((s, v) => s + v, 0) / ns.length : formulaError('#DIV/0!'); }
    case 'COUNT':      return args.flatMap(flatten).filter(v => typeof v === 'number').length;
    case 'COUNTA':     return args.flatMap(flatten).filter(v => v != null && v !== '').length;
    case 'COUNTBLANK': return args.flatMap(flatten).filter(v => v == null || v === '').length;
    case 'COUNTIF':    return countIf(flatten(args[0]), args[1]);
    // ── Text ────────────────────────────────────────────────────────────────
    case 'LEN':        return toStr(args[0]).length;
    case 'LEFT':       return toStr(args[0]).slice(0, Math.max(0, toNum(args[1] ?? 1)));
    case 'RIGHT':      { const s = toStr(args[0]); const n = Math.max(0, toNum(args[1] ?? 1)); return n >= s.length ? s : s.slice(s.length - n); }
    case 'MID':        { const s = toStr(args[0]); const start = Math.max(1, toNum(args[1])) - 1; const len = Math.max(0, toNum(args[2])); return s.slice(start, start + len); }
    case 'UPPER':      return toStr(args[0]).toUpperCase();
    case 'LOWER':      return toStr(args[0]).toLowerCase();
    case 'TRIM':       return toStr(args[0]).replace(/\s+/g, ' ').trim();
    case 'EXACT':      return toStr(args[0]) === toStr(args[1]);
    // A text not found is #VALUE! (so ISNUMBER(SEARCH(...)) stays FALSE).
    case 'FIND':       { const needle = toStr(args[0]); const hay = toStr(args[1]); const start = Math.max(1, toNum(args[2] ?? 1)) - 1; const idx = hay.indexOf(needle, start); return idx < 0 ? formulaError('#VALUE!') : idx + 1; }
    case 'SEARCH':     { const needle = toStr(args[0]).toLowerCase(); const hay = toStr(args[1]).toLowerCase(); const start = Math.max(1, toNum(args[2] ?? 1)) - 1; const idx = hay.indexOf(needle, start); return idx < 0 ? formulaError('#VALUE!') : idx + 1; }
    case 'CONCATENATE':
    case 'CONCAT':     return args.flatMap(flatten).map(v => toStr(v)).join('');
    case 'T':          { const s = operand(args[0]); return typeof s === 'string' ? s : ''; }
    case 'N':          { const s = operand(args[0]); return typeof s === 'number' ? s : typeof s === 'boolean' ? (s ? 1 : 0) : 0; }
    case 'VALUE':      return toNum(args[0]);
    // ── Date / time ─────────────────────────────────────────────────────────
    case 'TODAY':      return todaySerial();
    case 'NOW':        return nowSerial();
    case 'DATE':       return dateToSerial(toNum(args[0]), toNum(args[1]), toNum(args[2]));
    case 'YEAR':       return serialToDate(toNum(args[0])).y;
    case 'MONTH':      return serialToDate(toNum(args[0])).m;
    case 'DAY':        return serialToDate(toNum(args[0])).d;
    case 'WEEKDAY':    {
      // return type 1 (Sun=1..Sat=7) default; type 2 = Mon=1..Sun=7; type 3 = Mon=0..Sun=6.
      const d = serialToJsDate(toNum(args[0]));
      const jsDow = d.getUTCDay(); // Sun=0..Sat=6
      const rt = toNum(args[1] ?? 1);
      if (rt === 2) return jsDow === 0 ? 7 : jsDow;
      if (rt === 3) return jsDow === 0 ? 6 : jsDow - 1;
      return jsDow + 1;
    }
    default:
      throw new FormulaFailure('unsupported');
  }
}

/** §18.17.7.278 ROUND. Normative: digits 0–4 round down and 5–9 round up by
 * magnitude, so ties round away from zero (ROUND(-1.475,2) = -1.48).
 * Library policy: the digits are x's shortest round-trip decimal spelling.
 * Round that spelling directly, converting back to a double only once:
 * binary products can lose a written tie, while exponent-shifting through
 * intermediate doubles can change x even when no decimal digit is discarded.
 * Unsupported: fractional number-digits has no rule established here.
 * Callers pass finite operands; a result outside the finite-double range is
 * the §18.17.3 #NUM! error. */
function roundHalfAwayFromZero(x: number, digits: number): number {
  if (!Number.isInteger(digits)) throw new FormulaFailure('unsupported');
  if (x === 0) return 0;
  // Finite doubles' shortest spellings fit in decimal exponents -324..308
  // with at most 17 significant digits. This resource bound covers every
  // place that can change them and keeps exponents in integer spelling.
  const d = Math.max(-400, Math.min(400, digits));
  const [mantissa, exponent = '0'] = String(Math.abs(x)).split('e');
  const point = mantissa.indexOf('.');
  const decimalDigits = mantissa.replace('.', '');
  const integerPlaces = (point === -1 ? mantissa.length : point) + Number(exponent);
  const retainedPlaces = integerPlaces + d;
  // No discarded digit: preserve the original double exactly.
  if (retainedPlaces >= decimalDigits.length) return x;
  if (retainedPlaces < 0) return 0;
  const retained = retainedPlaces === 0 ? 0n : BigInt(decimalDigits.slice(0, retainedPlaces));
  const rounded = retained + (decimalDigits[retainedPlaces]! >= '5' ? 1n : 0n);
  const value = Number(`${rounded}e${-d}`);
  if (!Number.isFinite(value)) throw new FormulaFailure('error', '#NUM!');
  return value === 0 ? 0 : Math.sign(x) * value;
}

function countIf(source: EvalScalar[], criteria: EvalValue): number {
  const pred = makeCriteriaPredicate(criteria);
  let n = 0;
  for (const v of source) if (pred(v)) n++;
  return n;
}

function sumIf(source: EvalScalar[], criteria: EvalValue, sumRange: EvalScalar[] | null): number {
  const pred = makeCriteriaPredicate(criteria);
  const target = sumRange ?? source;
  let sum = 0;
  for (let i = 0; i < source.length; i++) {
    if (pred(source[i])) {
      const t = target[i];
      if (typeof t === 'number') sum += t;
    }
  }
  return sum;
}

// Microsoft AVERAGEIF documentation: an empty average_range cell is excluded,
// so a matched blank (null) is not in the denominator while a matched 0 is,
// and no remaining matched cell is #DIV/0!.
// https://support.microsoft.com/en-us/excel/functions/averageif-function
// A direct-reference average_range is resized to `source`'s dimensions;
// computed expressions retain prior index pairing (see `resizedTargetValues`).
// Unresolved, kept as before without Excel evidence: an index past a shorter
// computed target is counted instead of treated as a blank. A matched text,
// boolean or error target adds nothing to the SUMIF-style numerator but counts.
function averageIf(source: EvalScalar[], criteria: EvalValue, averageRange: EvalScalar[] | null): EvalScalar {
  const pred = makeCriteriaPredicate(criteria);
  const target = averageRange ?? source;
  let sum = 0;
  let count = 0;
  for (let i = 0; i < source.length; i++) {
    if (pred(source[i])) {
      const t = target[i];
      if (t === null) continue;
      if (typeof t === 'number') sum += t;
      count++;
    }
  }
  return count === 0 ? formulaError('#DIV/0!') : sum / count;
}

/** Build a predicate matching Excel's COUNTIF/SUMIF criteria syntax:
 *  a bare value (exact match) or a string like ">5", "<>foo", "=100". */
function makeCriteriaPredicate(criteria: EvalValue): (v: EvalScalar) => boolean {
  const raw = toScalar(criteria);
  // An error-valued criterion is not modeled.
  if (raw instanceof FormulaErrorValue) throw new FormulaFailure('unsupported');
  if (typeof raw !== 'string') {
    const rn = typeof raw === 'number' ? raw : null;
    return (v) => {
      if (rn !== null && typeof v === 'number') return v === rn;
      return v === raw;
    };
  }
  const m = raw.match(/^(<=|>=|<>|<|>|=)(.*)$/);
  const op = m ? m[1] : '=';
  const rhsStr = m ? m[2] : raw;
  const rhsNum = rhsStr.trim() === '' ? NaN : parseFloat(rhsStr);
  const rhsIsNum = !isNaN(rhsNum) && /^-?\d+(\.\d+)?$/.test(rhsStr.trim());
  return (v) => {
    if (rhsIsNum && typeof v === 'number') {
      switch (op) {
        case '<':  return v <  rhsNum;
        case '>':  return v >  rhsNum;
        case '<=': return v <= rhsNum;
        case '>=': return v >= rhsNum;
        case '<>': return v !== rhsNum;
        default:   return v === rhsNum;
      }
    }
    const sv = v == null ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE')
      : v instanceof FormulaErrorValue ? v.code : String(v);
    switch (op) {
      case '<>': return sv !== rhsStr;
      case '<':  return sv <  rhsStr;
      case '>':  return sv >  rhsStr;
      case '<=': return sv <= rhsStr;
      case '>=': return sv >= rhsStr;
      default:   return sv === rhsStr;
    }
  };
}

// Serial ↔ calendar-date conversions are delegated to the shared core
// `excel-date` module (ECMA-376 §18.17.4.1) so the 1900 Lotus leap-year-bug
// compensation and 1900/1904 epoch selection live in one place. Before this
// the formula engine carried its own `EXCEL_EPOCH_OFFSET = 25569` arithmetic,
// which had no leap-bug correction (serials < 60 read the wrong calendar day)
// and no 1904 support.
//
// This evaluator exists only for conditional-formatting formulas (see
// conditional-format.ts); cell values are never recalculated and are rendered
// from their cached `<v>` (number-format.ts). It always operates in the 1900
// date system, so TODAY()/NOW() inside a CF formula yield 1900-system serials
// and every core call below passes date1904=false. Threading the workbook's
// date system into CF evaluation is part of completing that evaluator (#1547).

function todaySerial(): number {
  const d = new Date();
  const utcMid = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  return utcDateToExcelSerial(utcMid, false);
}

function nowSerial(): number {
  return utcDateToExcelSerial(new Date(Date.now()), false);
}

function dateToSerial(y: number, m: number, d: number): number {
  // Excel rolls over out-of-range months/days (e.g. DATE(2019, 13, 1) = Jan 2020).
  // `Date.UTC` performs the same normalization. Whole days only (no time), so
  // floor to drop any sub-day residue the leap-bug boundary could introduce.
  return Math.floor(utcDateToExcelSerial(new Date(Date.UTC(y, m - 1, d)), false));
}

function serialToJsDate(serial: number): Date {
  return excelSerialToUtcDate(Math.floor(serial), false);
}

function serialToDate(serial: number): { y: number; m: number; d: number } {
  const d = serialToJsDate(serial);
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate() };
}

import { excelSerialToUtcDate, utcDateToExcelSerial } from '@silurus/ooxml-core';
import type { Cell, DefinedName } from './types.js';

// ────────────────────────────────────────────────────────────────
// Formula evaluator (conditional-formatting `expression` rules)
//
// Handles the narrow subset of Excel formulas used by CF expression rules:
// numeric/boolean literals, cell references (A1-style, with $ absolute
// markers), defined-name resolution, comparison/arithmetic operators, and
// a handful of functions (AND, OR, NOT, IF, ROUNDDOWN, ROUND, ROUNDUP,
// ISBLANK). Formula strings embed relative references that shift based on
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
  definedNames: Map<string, DefinedName>;
  /** Recursion guard for nested defined-name resolution. */
  depth: number;
}

export type EvalScalar = number | boolean | string | null;
type EvalValue = EvalScalar | EvalScalar[];

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

const MAX_DEFINED_NAME_DEPTH = 8;

export function evalFormulaToBool(formula: string, ctx: EvalCtx): boolean {
  try {
    const v = evalFormula(formula, ctx);
    return toBool(v);
  } catch {
    return false;
  }
}

// Exact evaluation for places where a wrong value must not be guessed: the
// operands of a `cellIs` rule and the activity condition of a scale rule,
// both of which can stop lower-priority rules (§18.3.1.10 stopIfTrue).
// The lenient evaluator above approximates (an unknown function or name is
// 0, a non-numeric string coerces through parseFloat, x/0 is 0, unknown
// characters are skipped). In strict mode each of those is instead
// "unevaluable", and the caller treats the rule as not matching: literals,
// cell references, defined names, + - * / & comparisons and a small set of
// functions whose implementation here is exact are the supported subset.
let strict = false;

// Date functions are excluded: the context does not carry the workbook's
// date system (date1904), and DATE's 0-1899 year offset is not modelled.
const STRICT_FUNCTIONS = new Set(['AND', 'OR', 'NOT', 'IF', 'TRUE', 'FALSE', 'ABS', 'INT']);

class Unevaluable extends Error {}

/** The formula's scalar value, or `undefined` when the formula falls
 *  outside the exactly evaluated subset or yields an Excel error. */
export function evalFormulaStrict(formula: string, ctx: EvalCtx): EvalScalar | undefined {
  const previous = strict;
  strict = true;
  try {
    const v = evalFormula(formula, ctx);
    if (Array.isArray(v)) return undefined;
    return v;
  } catch {
    return undefined;
  } finally {
    strict = previous;
  }
}

function unevaluable(what: string): never {
  throw new Unevaluable(what);
}

function toBool(v: EvalValue): boolean {
  if (strict && Array.isArray(v)) unevaluable('array as scalar');
  const s = toScalar(v);
  if (typeof s === 'boolean') return s;
  if (typeof s === 'number') return s !== 0;
  if (strict && typeof s === 'string') {
    // Excel: only "TRUE"/"FALSE" convert to a logical; other text is #VALUE!.
    const up = s.toUpperCase();
    if (up === 'TRUE' || up === 'FALSE') return up === 'TRUE';
    unevaluable('text as logical');
  }
  if (typeof s === 'string') return s.length > 0 && s.toUpperCase() !== 'FALSE';
  return false;
}

function toNum(v: EvalValue): number {
  if (strict && Array.isArray(v)) unevaluable('array as scalar');
  const s = toScalar(v);
  if (typeof s === 'number') return s;
  if (typeof s === 'boolean') return s ? 1 : 0;
  if (s == null) return 0;
  if (strict) {
    // Excel converts numeric text in arithmetic; other text is #VALUE!.
    const t = String(s).trim();
    const n = t === '' ? NaN : Number(t);
    if (!Number.isFinite(n)) unevaluable('text as number');
    return n;
  }
  const n = parseFloat(String(s));
  return isNaN(n) ? 0 : n;
}

function toStr(v: EvalValue): string {
  const s = toScalar(v);
  if (s == null) return '';
  if (typeof s === 'boolean') return s ? 'TRUE' : 'FALSE';
  return String(s);
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
      toks.push({ kind: 'str', text: buf });
      i = j + 1;
      continue;
    }
    if (c === '#') {
      // ECMA-376 §18.17.2 error literals (e.g. `#REF!` left by a deleted
      // reference). Skipping the `#` would read `REF` as an unknown name
      // worth 0 and let the rule match.
      const m = /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA)/u.exec(s.slice(i));
      if (!m) throw new Error('unknown error literal');
      toks.push({ kind: 'error', text: m[0] });
      i += m[0].length;
      continue;
    }
    if (c >= '0' && c <= '9') {
      let j = i;
      while (j < s.length && ((s[j] >= '0' && s[j] <= '9') || s[j] === '.')) j++;
      const text = s.slice(i, j);
      // Exponents and repeated points are not tokenized here.
      if (strict && (text.split('.').length > 2 || /[eE]/.test(s[j] ?? ''))) {
        unevaluable('number literal');
      }
      toks.push({ kind: 'num', text });
      i = j;
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
    // `$?[A-Z]+\$?[0-9]+` (case-insensitive).
    if (c === '$' || isIdentStart(c)) {
      let j = i;
      while (j < s.length && (s[j] === '$' || isIdentPart(s[j]))) j++;
      const text = s.slice(i, j);
      i = j;
      const ref = tryParseCellRef(text);
      if (ref) {
        toks.push({ kind: 'ref', text, ref });
      } else {
        const up = text.toUpperCase();
        if (up === 'TRUE' || up === 'FALSE') toks.push({ kind: 'bool', text: up });
        else toks.push({ kind: 'name', text });
      }
      continue;
    }
    // Unknown character — skip (strict: e.g. `!`, `{`, `[` change meaning).
    if (strict) unevaluable(`character ${c}`);
    i++;
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
  return { colAbs, col, rowAbs, row: rowNum };
}

interface Parser {
  toks: Tok[];
  pos: number;
}

function evalFormula(formula: string, ctx: EvalCtx): EvalValue {
  const toks = tokenize(formula);
  const p: Parser = { toks, pos: 0 };
  if (strict && toks.length === 0) unevaluable('empty formula');
  const v = parseExpr(p, ctx);
  // Unconsumed tokens (`%`, `^`, a missing operator) mean a construct this
  // evaluator does not model.
  if (strict && p.pos !== toks.length) unevaluable('trailing tokens');
  return v;
}

function peek(p: Parser): Tok | undefined { return p.toks[p.pos]; }
function consume(p: Parser): Tok | undefined { return p.toks[p.pos++]; }

function parseExpr(p: Parser, ctx: EvalCtx): EvalValue {
  return parseCmp(p, ctx);
}

function parseCmp(p: Parser, ctx: EvalCtx): EvalValue {
  let left = parseConcat(p, ctx);
  const t = peek(p);
  if (t && t.kind === 'op' && (t.text === '<' || t.text === '>' || t.text === '<=' || t.text === '>=' || t.text === '=' || t.text === '<>')) {
    consume(p);
    const right = parseConcat(p, ctx);
    return strict ? applyCmpStrict(t.text, left, right) : applyCmp(t.text, left, right);
  }
  return left;
}

function parseConcat(p: Parser, ctx: EvalCtx): EvalValue {
  let left = parseAdd(p, ctx);
  while (true) {
    const t = peek(p);
    if (!t || t.kind !== 'op' || t.text !== '&') break;
    consume(p);
    const right = parseAdd(p, ctx);
    if (strict) {
      // A non-integer number's text form follows Excel's General format,
      // which String() does not reproduce.
      for (const v of [left, right]) {
        if (Array.isArray(v) || (typeof v === 'number' && !Number.isSafeInteger(v))) unevaluable('& operand');
      }
    }
    left = toStr(left) + toStr(right);
  }
  return left;
}

function applyCmp(op: string, a: EvalValue, b: EvalValue): boolean {
  // Numeric-first comparison; fall back to string compare if either side is
  // a non-numeric string. Matches Excel's behavior for dates (stored as
  // serials) and arithmetic operations.
  const an = typeof a === 'string' && isNaN(parseFloat(a)) ? null : toNum(a);
  const bn = typeof b === 'string' && isNaN(parseFloat(b)) ? null : toNum(b);
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

/** Excel's comparison: text compares case-insensitively, and values of
 *  different types order number < text < logical. An empty cell compares
 *  as 0 against a number, "" against text and FALSE against a logical. */
function applyCmpStrict(op: string, a: EvalValue, b: EvalValue): boolean {
  if (Array.isArray(a) || Array.isArray(b)) unevaluable('array comparison');
  const rank = (v: EvalScalar) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
  const fill = (v: EvalScalar, other: EvalScalar): EvalScalar => {
    if (v != null) return v;
    if (typeof other === 'string') return '';
    if (typeof other === 'boolean') return false;
    return 0;
  };
  const x = fill(a, b);
  const y = fill(b, a);
  let c: number;
  if (rank(x) !== rank(y)) c = rank(x) - rank(y);
  else if (typeof x === 'string') {
    const xs = x.toLowerCase(), ys = (y as string).toLowerCase();
    c = xs < ys ? -1 : xs > ys ? 1 : 0;
  } else {
    const xn = Number(x), yn = Number(y);
    c = xn < yn ? -1 : xn > yn ? 1 : 0;
  }
  switch (op) {
    case '<':  return c < 0;
    case '>':  return c > 0;
    case '<=': return c <= 0;
    case '>=': return c >= 0;
    case '=':  return c === 0;
    default:   return c !== 0;
  }
}

function parseAdd(p: Parser, ctx: EvalCtx): EvalValue {
  let left = parseMul(p, ctx);
  while (true) {
    const t = peek(p);
    if (!t || t.kind !== 'op' || (t.text !== '+' && t.text !== '-')) break;
    consume(p);
    const right = parseMul(p, ctx);
    left = t.text === '+' ? toNum(left) + toNum(right) : toNum(left) - toNum(right);
  }
  return left;
}

function parseMul(p: Parser, ctx: EvalCtx): EvalValue {
  let left = parseUnary(p, ctx);
  while (true) {
    const t = peek(p);
    if (!t || t.kind !== 'op' || (t.text !== '*' && t.text !== '/')) break;
    consume(p);
    const right = parseUnary(p, ctx);
    if (t.text === '*') left = toNum(left) * toNum(right);
    else {
      const rn = toNum(right);
      if (strict && rn === 0) unevaluable('#DIV/0!');
      left = rn === 0 ? 0 : toNum(left) / rn;
    }
  }
  return left;
}

function parseUnary(p: Parser, ctx: EvalCtx): EvalValue {
  const t = peek(p);
  if (t && t.kind === 'op' && t.text === '-') { consume(p); return -toNum(parseUnary(p, ctx)); }
  if (t && t.kind === 'op' && t.text === '+') { consume(p); return toNum(parseUnary(p, ctx)); }
  return parsePrimary(p, ctx);
}

function parsePrimary(p: Parser, ctx: EvalCtx): EvalValue {
  const t = consume(p);
  if (!t) {
    if (strict) unevaluable('missing operand');
    return 0;
  }
  if (t.kind === 'num') return parseFloat(t.text);
  if (t.kind === 'str') return t.text;
  if (t.kind === 'bool') return t.text === 'TRUE';
  // An error value propagates through the operators and functions this
  // evaluator models, so the rule's result is an error: Excel applies a
  // conditional format only when its formula evaluates to TRUE.
  if (t.kind === 'error') throw new Error(t.text);
  if (t.kind === 'lparen') {
    const v = parseExpr(p, ctx);
    const next = consume(p);
    if (!next || next.kind !== 'rparen') throw new Error('missing )');
    return v;
  }
  if (t.kind === 'ref') {
    // Range: `A1:B5` — resolve as array of cell values.
    if (peek(p)?.kind === 'colon') {
      consume(p);
      const right = consume(p);
      if (right?.kind !== 'ref' || !right.ref) throw new Error('range: expected ref after :');
      return resolveRange(t.ref!, right.ref, ctx);
    }
    return resolveRef(t.ref!, ctx);
  }
  if (t.kind === 'name') {
    // Function call: NAME(args)
    if (peek(p)?.kind === 'lparen') {
      consume(p);
      const args: EvalValue[] = [];
      if (peek(p)?.kind !== 'rparen') {
        args.push(parseExpr(p, ctx));
        while (peek(p)?.kind === 'comma') {
          consume(p);
          args.push(parseExpr(p, ctx));
        }
      }
      const next = consume(p);
      if (!next || next.kind !== 'rparen') throw new Error('missing )');
      return callFunc(t.text, args, ctx);
    }
    // Defined-name reference: substitute and evaluate. Strict mode does not:
    // the substitution ignores the name's sheet prefix and scope.
    if (strict) unevaluable(`name ${t.text}`);
    const dn = ctx.definedNames.get(t.text);
    if (dn && ctx.depth < MAX_DEFINED_NAME_DEPTH) {
      // Strip `SheetName!` prefix if present; keep just the ref body.
      const body = stripSheetPrefix(dn.formula);
      // Workbook-level defined names anchor at A1 for relative-ref shifts.
      const inner: EvalCtx = {
        ...ctx,
        anchorRow: 1,
        anchorCol: 1,
        depth: ctx.depth + 1,
      };
      return evalFormula(body, inner);
    }
    return 0;
  }
  if (strict) unevaluable(`token ${t.text}`);
  return 0;
}

function stripSheetPrefix(formula: string): string {
  // Match `'Sheet Name'!ref` or `SheetName!ref`. Only the leading reference
  // prefix is stripped; we don't need cross-sheet lookups because defined
  // names here point to cells on the active sheet.
  const m = formula.match(/^(?:'[^']*'|[A-Za-z_][A-Za-z0-9_.]*)!(.*)$/);
  return m ? m[1] : formula;
}

function resolveRef(
  ref: { colAbs: boolean; col: number; rowAbs: boolean; row: number },
  ctx: EvalCtx,
): EvalScalar {
  const col = ref.colAbs ? ref.col : ref.col + (ctx.col - ctx.anchorCol);
  const row = ref.rowAbs ? ref.row : ref.row + (ctx.row - ctx.anchorRow);
  if (strict && (row < 1 || col < 1 || row > 1_048_576 || col > 16_384)) unevaluable('#REF!');
  const cell = ctx.cellIndex.get(`${row}:${col}`);
  if (strict && cell?.value.type === 'error') unevaluable('error cell');
  return cellValueToEval(cell);
}

function resolveRange(
  a: { colAbs: boolean; col: number; rowAbs: boolean; row: number },
  b: { colAbs: boolean; col: number; rowAbs: boolean; row: number },
  ctx: EvalCtx,
): EvalScalar[] {
  const ac = a.colAbs ? a.col : a.col + (ctx.col - ctx.anchorCol);
  const ar = a.rowAbs ? a.row : a.row + (ctx.row - ctx.anchorRow);
  const bc = b.colAbs ? b.col : b.col + (ctx.col - ctx.anchorCol);
  const br = b.rowAbs ? b.row : b.row + (ctx.row - ctx.anchorRow);
  const c1 = Math.min(ac, bc), c2 = Math.max(ac, bc);
  const r1 = Math.min(ar, br), r2 = Math.max(ar, br);
  const out: EvalScalar[] = [];
  // Cap range size to avoid pathological formulas like A:A (≈1M cells).
  // 4096 cells is plenty for CF use cases.
  const maxCells = 4096;
  for (let r = r1; r <= r2 && out.length < maxCells; r++) {
    for (let c = c1; c <= c2 && out.length < maxCells; c++) {
      out.push(cellValueToEval(ctx.cellIndex.get(`${r}:${c}`)));
    }
  }
  return out;
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
    case 'error':  return null;
    case 'empty':
    default:       return null;
  }
}

function callFunc(nameRaw: string, args: EvalValue[], ctx: EvalCtx): EvalValue {
  const name = nameRaw.toUpperCase();
  if (strict && (!STRICT_FUNCTIONS.has(name) || args.some(Array.isArray)
    || (name === 'IF' && (args.length < 2 || args.length > 3)))) {
    unevaluable(`function ${name}`);
  }
  switch (name) {
    // ── Logic ───────────────────────────────────────────────────────────────
    case 'AND':        return args.flatMap(flatten).every(a => toBool(a));
    case 'OR':         return args.flatMap(flatten).some(a => toBool(a));
    case 'NOT':        return !toBool(args[0]);
    case 'IF':         return toBool(args[0]) ? (args[1] ?? true) : (args[2] ?? false);
    case 'IFERROR':    return args[0] == null ? (args[1] ?? 0) : args[0];
    case 'IFS': {
      for (let i = 0; i + 1 < args.length; i += 2) {
        if (toBool(args[i])) return args[i + 1];
      }
      return null;
    }
    case 'TRUE':       return true;
    case 'FALSE':      return false;
    // ── Type checks ─────────────────────────────────────────────────────────
    case 'ISBLANK':    { const s = toScalar(args[0]); return s == null || s === ''; }
    case 'ISNUMBER':   return typeof toScalar(args[0]) === 'number';
    case 'ISTEXT':     return typeof toScalar(args[0]) === 'string';
    case 'ISNONTEXT':  return typeof toScalar(args[0]) !== 'string';
    case 'ISERROR':
    case 'ISERR':
    case 'ISNA':       return toScalar(args[0]) == null;
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
    case 'ROUND': {
      const n = toNum(args[0]); const d = toNum(args[1]);
      const p = Math.pow(10, d);
      return Math.round(n * p) / p;
    }
    case 'INT':        return Math.floor(toNum(args[0]));
    case 'TRUNC':      { const n = toNum(args[0]); const d = toNum(args[1] ?? 0); const p = Math.pow(10, d); return (n >= 0 ? Math.floor(n * p) : Math.ceil(n * p)) / p; }
    case 'CEILING':    { const n = toNum(args[0]); const sig = toNum(args[1] ?? 1); return sig === 0 ? 0 : Math.ceil(n / sig) * sig; }
    case 'FLOOR':      { const n = toNum(args[0]); const sig = toNum(args[1] ?? 1); return sig === 0 ? 0 : Math.floor(n / sig) * sig; }
    case 'MOD':        { const a = toNum(args[0]); const b = toNum(args[1]); return b === 0 ? null : a - Math.floor(a / b) * b; }
    case 'POWER':      return Math.pow(toNum(args[0]), toNum(args[1]));
    case 'SQRT':       { const n = toNum(args[0]); return n < 0 ? null : Math.sqrt(n); }
    case 'ABS':        return Math.abs(toNum(args[0]));
    case 'SIGN':       { const n = toNum(args[0]); return n > 0 ? 1 : n < 0 ? -1 : 0; }
    case 'EXP':        return Math.exp(toNum(args[0]));
    case 'LN':         { const n = toNum(args[0]); return n <= 0 ? null : Math.log(n); }
    case 'LOG10':      { const n = toNum(args[0]); return n <= 0 ? null : Math.log10(n); }
    // ── Aggregates ──────────────────────────────────────────────────────────
    case 'MIN':        { const ns = args.flatMap(flatten).filter(v => typeof v === 'number') as number[]; return ns.length ? Math.min(...ns) : 0; }
    case 'MAX':        { const ns = args.flatMap(flatten).filter(v => typeof v === 'number') as number[]; return ns.length ? Math.max(...ns) : 0; }
    case 'SUM':        return args.flatMap(flatten).reduce<number>((s, v) => s + (typeof v === 'number' ? v : 0), 0);
    case 'AVERAGE':    { const ns = args.flatMap(flatten).filter(v => typeof v === 'number') as number[]; return ns.length ? ns.reduce((s, v) => s + v, 0) / ns.length : null; }
    case 'COUNT':      return args.flatMap(flatten).filter(v => typeof v === 'number').length;
    case 'COUNTA':     return args.flatMap(flatten).filter(v => v != null && v !== '').length;
    case 'COUNTBLANK': return args.flatMap(flatten).filter(v => v == null || v === '').length;
    case 'COUNTIF':    return countIf(flatten(args[0]), args[1]);
    case 'SUMIF':      return sumIf(flatten(args[0]), args[1], args[2] !== undefined ? flatten(args[2]) : null);
    case 'AVERAGEIF':  {
      const src = flatten(args[0]);
      const sum = sumIf(src, args[1], args[2] !== undefined ? flatten(args[2]) : null);
      const count = countIf(src, args[1]);
      return count === 0 ? null : toNum(sum) / count;
    }
    // ── Text ────────────────────────────────────────────────────────────────
    case 'LEN':        return toStr(args[0]).length;
    case 'LEFT':       return toStr(args[0]).slice(0, Math.max(0, toNum(args[1] ?? 1)));
    case 'RIGHT':      { const s = toStr(args[0]); const n = Math.max(0, toNum(args[1] ?? 1)); return n >= s.length ? s : s.slice(s.length - n); }
    case 'MID':        { const s = toStr(args[0]); const start = Math.max(1, toNum(args[1])) - 1; const len = Math.max(0, toNum(args[2])); return s.slice(start, start + len); }
    case 'UPPER':      return toStr(args[0]).toUpperCase();
    case 'LOWER':      return toStr(args[0]).toLowerCase();
    case 'TRIM':       return toStr(args[0]).replace(/\s+/g, ' ').trim();
    case 'EXACT':      return toStr(args[0]) === toStr(args[1]);
    case 'FIND':       { const needle = toStr(args[0]); const hay = toStr(args[1]); const start = Math.max(1, toNum(args[2] ?? 1)) - 1; const idx = hay.indexOf(needle, start); return idx < 0 ? null : idx + 1; }
    case 'SEARCH':     { const needle = toStr(args[0]).toLowerCase(); const hay = toStr(args[1]).toLowerCase(); const start = Math.max(1, toNum(args[2] ?? 1)) - 1; const idx = hay.indexOf(needle, start); return idx < 0 ? null : idx + 1; }
    case 'CONCATENATE':
    case 'CONCAT':     return args.flatMap(flatten).map(v => v == null ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v)).join('');
    case 'T':          { const s = toScalar(args[0]); return typeof s === 'string' ? s : ''; }
    case 'N':          { const s = toScalar(args[0]); return typeof s === 'number' ? s : typeof s === 'boolean' ? (s ? 1 : 0) : 0; }
    case 'VALUE':      return toNum(args[0]);
    // ── Reference ───────────────────────────────────────────────────────────
    case 'ROW':        return ctx.row;    // no-arg form only (current cell row)
    case 'COLUMN':     return ctx.col;    // no-arg form only (current cell col)
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
      return 0;
  }
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

/** Build a predicate matching Excel's COUNTIF/SUMIF criteria syntax:
 *  a bare value (exact match) or a string like ">5", "<>foo", "=100". */
function makeCriteriaPredicate(criteria: EvalValue): (v: EvalScalar) => boolean {
  const raw = toScalar(criteria);
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
    const sv = v == null ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v);
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
// The formula engine always operates in the 1900 date system: its serials are
// consumed by the number-format layer, whose volatile-recompute path formats
// TODAY()/NOW() against the 1900 epoch regardless of `<workbookPr date1904>`
// (see number-format.ts). So every core call below passes date1904=false.
// Threading the workbook's date system into the formula engine is a separate
// track; this change is purely about retiring the duplicate serial math.

export function todaySerial(): number {
  const d = new Date();
  const utcMid = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  return utcDateToExcelSerial(utcMid, false);
}

export function nowSerial(): number {
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

import type { Cell } from './types.js';

// ────────────────────────────────────────────────────────────────
// Exact evaluation of conditional-formatting operand formulas
//
// Used where a guessed value could stop lower-priority rules
// (ECMA-376 §18.3.1.10 stopIfTrue): the operands of a `cellIs` rule and the
// activity condition of a colorScale / dataBar / iconSet rule ([MS-XLSX]
// 2.6.27). A value is returned only for a grammar whose Excel semantics are
// reproduced exactly here:
//
//   - numeric literals (optional exponent) and "string" literals (`""`
//     escapes a quote);
//   - single-cell A1 references, relative or `$`-absolute, shifted by the
//     evaluation cell's offset from the rule range's top-left anchor. A
//     reference yields the cell's typed value: number, string, boolean or
//     blank (kept distinct; an error cell is unevaluable);
//   - unary minus, and `+ - * /`, on numbers only; every literal and every
//     intermediate result must be finite (an overflow is #NUM! and x/0 is
//     #DIV/0! in Excel, so both are unevaluable);
//   - `&` on strings only;
//   - parentheses for grouping, with Excel's precedence
//     (negation > * / > + - > &).
//
// Everything else is unevaluable, and callers treat the rule as not matching
// (it neither formats nor stops): functions, names, comparisons, `^`, `%`,
// ranges, sheet-qualified or R1C1 references, and any implicit conversion
// (text to number, blank or logical in arithmetic, number to text in `&`).
// The lenient evaluator in formula.ts, used by `expression` rules, is
// separate.
// ────────────────────────────────────────────────────────────────

/** A typed operand value; `null` is a blank cell. */
export type CfExactValue = number | string | boolean | null;

export interface CfExactContext {
  row: number;
  col: number;
  anchorRow: number;
  anchorCol: number;
  cellIndex: ReadonlyMap<string, Cell>;
}

const MAX_ROW = 1_048_576;
const MAX_COL = 16_384;

class Unevaluable extends Error {}

type Token =
  | { kind: 'num'; value: number }
  | { kind: 'str'; value: string }
  | { kind: 'ref'; colAbs: boolean; col: number; rowAbs: boolean; row: number }
  | { kind: 'op'; value: '+' | '-' | '*' | '/' | '&' | '(' | ')' };

const NUMBER = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/;
const REF = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/;

function tokenize(src: string): Token[] {
  const out: Token[] = [];
  let i = 0;
  while (i < src.length) {
    const c = src[i];
    if (c === ' ' || c === '\t' || c === '\r' || c === '\n') { i++; continue; }
    if ('+-*/&()'.includes(c)) {
      out.push({ kind: 'op', value: c as '+' });
      i++;
      continue;
    }
    if (c === '"') {
      let j = i + 1;
      let text = '';
      for (;;) {
        if (j >= src.length) throw new Unevaluable('unclosed string');
        if (src[j] === '"') {
          if (src[j + 1] === '"') { text += '"'; j += 2; continue; }
          break;
        }
        text += src[j++];
      }
      out.push({ kind: 'str', value: text });
      i = j + 1;
      continue;
    }
    const rest = src.slice(i);
    const num = NUMBER.exec(rest);
    if (num) {
      const value = Number(num[0]);
      if (!Number.isFinite(value)) throw new Unevaluable('#NUM!');
      out.push({ kind: 'num', value });
      i += num[0].length;
    } else {
      const ref = REF.exec(rest);
      if (!ref) throw new Unevaluable(`unsupported at ${i}`);
      let col = 0;
      for (const ch of ref[2].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
      out.push({
        kind: 'ref', colAbs: ref[1] === '$', col, rowAbs: ref[3] === '$', row: Number(ref[4]),
      });
      i += ref[0].length;
    }
    // A literal or reference directly followed by a name character is
    // something else (`1E`, `A1B`, a function name such as `LOG10(`
    // continues with `(`, which the parser rejects).
    if (/^[A-Za-z0-9_.$!:]/.test(src.slice(i))) throw new Unevaluable('malformed token');
  }
  return out;
}

class Parser {
  private pos = 0;
  constructor(private readonly toks: Token[], private readonly ctx: CfExactContext) {}

  parse(): CfExactValue {
    if (this.toks.length === 0) throw new Unevaluable('empty');
    const v = this.concat();
    if (this.pos !== this.toks.length) throw new Unevaluable('trailing tokens');
    return v;
  }

  private peekOp(...ops: string[]): string | undefined {
    const t = this.toks[this.pos];
    return t?.kind === 'op' && ops.includes(t.value) ? t.value : undefined;
  }

  private concat(): CfExactValue {
    let left = this.additive();
    while (this.peekOp('&')) {
      this.pos++;
      const right = this.additive();
      if (typeof left !== 'string' || typeof right !== 'string') throw new Unevaluable('& operand');
      left = left + right;
    }
    return left;
  }

  private additive(): CfExactValue {
    let left = this.multiplicative();
    for (let op = this.peekOp('+', '-'); op; op = this.peekOp('+', '-')) {
      this.pos++;
      left = arithmetic(op, left, this.multiplicative());
    }
    return left;
  }

  private multiplicative(): CfExactValue {
    let left = this.unary();
    for (let op = this.peekOp('*', '/'); op; op = this.peekOp('*', '/')) {
      this.pos++;
      left = arithmetic(op, left, this.unary());
    }
    return left;
  }

  private unary(): CfExactValue {
    if (this.peekOp('-')) {
      this.pos++;
      const v = this.unary();
      if (typeof v !== 'number') throw new Unevaluable('negation operand');
      return -v;
    }
    return this.primary();
  }

  private primary(): CfExactValue {
    const t = this.toks[this.pos++];
    if (!t) throw new Unevaluable('missing operand');
    if (t.kind === 'num' || t.kind === 'str') return t.value;
    if (t.kind === 'ref') return this.resolve(t);
    if (t.value === '(') {
      const v = this.concat();
      if (!this.peekOp(')')) throw new Unevaluable('missing )');
      this.pos++;
      return v;
    }
    throw new Unevaluable(`unexpected ${t.value}`);
  }

  private resolve(ref: Extract<Token, { kind: 'ref' }>): CfExactValue {
    if (ref.row < 1 || ref.row > MAX_ROW || ref.col < 1 || ref.col > MAX_COL) {
      throw new Unevaluable('not a cell reference');
    }
    const row = ref.rowAbs ? ref.row : ref.row + (this.ctx.row - this.ctx.anchorRow);
    const col = ref.colAbs ? ref.col : ref.col + (this.ctx.col - this.ctx.anchorCol);
    // Excel wraps a shifted relative reference around the sheet edge; not
    // modelled.
    if (row < 1 || row > MAX_ROW || col < 1 || col > MAX_COL) throw new Unevaluable('#REF!');
    const value = this.ctx.cellIndex.get(`${row}:${col}`)?.value;
    if (!value) return null;
    switch (value.type) {
      case 'empty': return null;
      case 'number': return value.number;
      case 'text': return value.text;
      case 'bool': return value.bool;
      default: throw new Unevaluable(`cell ${value.type}`);
    }
  }
}

function arithmetic(op: string, a: CfExactValue, b: CfExactValue): number {
  if (typeof a !== 'number' || typeof b !== 'number') throw new Unevaluable('arithmetic operand');
  if (op === '/' && b === 0) throw new Unevaluable('#DIV/0!');
  const r = op === '+' ? a + b : op === '-' ? a - b : op === '*' ? a * b : a / b;
  if (!Number.isFinite(r)) throw new Unevaluable('#NUM!');
  return r;
}

/** The formula's typed value, or `undefined` when it is outside the exactly
 *  evaluated grammar above or yields an Excel error. */
export function evalCfExact(formula: string, ctx: CfExactContext): CfExactValue | undefined {
  try {
    return new Parser(tokenize(formula), ctx).parse();
  } catch (e) {
    if (e instanceof Unevaluable) return undefined;
    throw e;
  }
}

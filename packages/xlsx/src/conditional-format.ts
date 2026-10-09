import type { Worksheet, Cell, WorksheetCellRange, CfStop, CfValue, Dxf, CfRule, CellFill, Border, DefinedName } from './types.js';
import { dxfFontToggle } from './dxf-font.js';
import { decodeCfLiteral, decodeCfOperandResult, type CfOperandValue } from './cf-operand.js';
import { evaluateFormula, formulaEvaluationToBool } from './formula.js';
import type { CfDiagnosticCollector, CfDiagnosticKind } from './cf-diagnostics.js';
import { buildCellCoordinateIndex } from './renderer-coordinate-index.js';
import { getWorksheetPolicy } from './worksheet-policy-context.js';

// ────────────────────────────────────────────────────────────────
// Conditional formatting
// ────────────────────────────────────────────────────────────────
export interface CompiledCfRule {
  rule: CfRule;
  sqref: WorksheetCellRange[];
  /** Document position (index into `Worksheet.conditionalFormats` and into
   *  that block's rules), captured before the priority sort. Names the rule
   *  in conditional-formatting diagnostics. */
  blockIndex: number;
  ruleIndex: number;
  scaleMin?: number;
  scaleMax?: number;
  scaleStops?: number[];
  barMin?: number;
  barMax?: number;
  top10Threshold?: number;
  top10IsTop?: boolean;
  avgValue?: number;
  avgIsAbove?: boolean;
  /** Population standard deviation of the sampled range, when the
   *  aboveAverage rule carries a `stdDev` attribute (ECMA-376 §18.3.1.10). */
  avgStdDev?: number;
  iconThresholds?: number[];
  cellIsOperands?: CellIsOperand[];
  /** Threshold admission (#1547): a cfvo the library cannot resolve without
   *  guessing. Such a scale rule is skipped (no paint, no stop) and reported
   *  when the visible evaluation reaches it. */
  thresholdUnsupported?: boolean;
  thresholdInvalid?: boolean;

}

// Protocol with CfRule::Other, fixed four bits; allocated once, not per cell.
const OTHER_FORMULA_PHASES = [[1, 'expression'], [2, 'cellIs'], [4, 'activity'], [8, 'threshold']] as const;

export interface CfContext {
  compiled: CompiledCfRule[];
  worksheet: Worksheet;
  cellIndex: ReadonlyMap<string, Cell>;
  definedNames: Map<string, DefinedName>;
}

export interface CfResult {
  fill?: CellFill;
  fontColor?: string;
  /** Font toggles from the matched rules' dxfs: `false` is an explicit off
   *  that overrides the cell's formatting, `undefined` leaves it. CF is the
   *  top formatting layer, so a renderer takes a defined value over the
   *  cell/table/PivotTable toggle. Observed in Excel's PDF export: cell-style
   *  bold/italic/underline/strike and a built-in table-style header's bold
   *  print off under an explicit-off rule and stay under a rule whose font
   *  omits them. The PivotTable-style case follows the same layering and
   *  was not separately exported. */
  fontBold?: boolean;
  fontItalic?: boolean;
  fontUnderline?: boolean;
  fontStrike?: boolean;
  /** Number format override from a matched CF dxf. Higher-priority rules win
   *  (first match through the rule list). Falls back to the cell's own style
   *  numFmt if unset. */
  numFmt?: { numFmtId: number; formatCode: string | null };
  dataBar?: { color: string; ratio: number; gradient: boolean };
  iconSet?: { name: string; index: number };
  /** Per-edge borders from matched CF rules (merged on top of the cell's base
   *  border). Mostly used by `expression` rules whose dxf only sets borders,
   *  e.g. highlighting today's column in a Gantt chart. */
  border?: Border;
}

function rangeContains(ranges: WorksheetCellRange[], row: number, col: number): boolean {
  for (const r of ranges) {
    if (row >= r.top && row <= r.bottom && col >= r.left && col <= r.right) return true;
  }
  return false;
}

function cellNumericValue(cell: Cell | undefined): number | null {
  if (!cell) return null;
  if (cell.value.type === 'number') return cell.value.number;
  return null;
}

function cellTextValue(cell: Cell | undefined): string | null {
  if (!cell) return null;
  if (cell.value.type === 'text') return cell.value.text;
  return null;
}

function collectNumericValuesInRanges(worksheet: Worksheet, ranges: WorksheetCellRange[]): number[] {
  const out: number[] = [];
  for (const row of worksheet.rows) {
    for (const c of row.cells) {
      if (c.value.type !== 'number') continue;
      if (rangeContains(ranges, c.row, c.col)) out.push(c.value.number);
    }
  }
  return out;
}

/** A resolved cfvo threshold, or the diagnostic kind that keeps the rule from
 *  using a guessed threshold. */
type CfvoResolution = number | CfDiagnosticKind;

/**
 * The numeric parameter of a non-min/max cfvo (ECMA-376 §18.3.1.11 `val`, or
 * the x14 `xm:f` text). Library policy (#1547), not an Office rule: `val` is
 * formula text and thresholds are not formula-evaluated, so only a plain
 * numeric literal is an established value. Any other present text (a
 * reference, `1+1`, a function, `50%`) is `unsupported`; a `formula` cfvo
 * without text is `invalid`. An absent value of another kind keeps the
 * library's existing default. The former `parseFloat` read `1+1` as 1 and a
 * reference as 0, i.e. it invented a threshold.
 */
function cfvoParameter(cfv: CfValue | CfStop): number | undefined | CfDiagnosticKind {
  const text = cfv.value;
  if (text == null || text.trim() === '') {
    return cfv.kind === 'formula' ? 'invalid' : undefined;
  }
  const literal = decodeCfLiteral(text);
  return typeof literal === 'number' ? literal : 'unsupported';
}

function resolveCfvoValue(cfv: CfValue | CfStop, samples: number[]): CfvoResolution {
  // A full-column CF range can exceed the JS argument stack. Scan in place;
  // this keeps the same extrema without spreading the whole range.
  let minv = samples.length ? Infinity : 0;
  let maxv = samples.length ? -Infinity : 0;
  for (const sample of samples) {
    if (sample < minv) minv = sample;
    if (sample > maxv) maxv = sample;
  }
  if (cfv.kind === 'min') return minv;
  if (cfv.kind === 'max') return maxv;
  // Unsupported x14 autoMin/autoMax need a zero clamp, not the standard
  // extrema; unknown types must not acquire the former default numeric 0.
  if (!['num', 'formula', 'percent', 'percentile'].includes(cfv.kind)) return 'unsupported';
  const n = cfvoParameter(cfv);
  if (typeof n === 'string') return n;
  switch (cfv.kind) {
    case 'percent': {
      const p = n ?? 50;
      return minv + (maxv - minv) * (p / 100);
    }
    case 'percentile': {
      if (!samples.length) return 0;
      const sorted = [...samples].sort((a, b) => a - b);
      const p = (n ?? 50) / 100;
      const idx = Math.max(0, Math.min(sorted.length - 1, Math.round(p * (sorted.length - 1))));
      return sorted[idx];
    }
    // `num`, `formula` (a literal) and kinds not modelled here: the literal,
    // or the existing 0 default for an absent value.
    default: return n ?? 0;
  }
}

/** Admit every threshold of a rule, or flag the rule and return undefined. */
function admitThresholds(entry: CompiledCfRule, values: CfvoResolution[]): number[] | undefined {
  let admitted = true;
  for (const value of values) {
    if (value === 'unsupported') {
      entry.thresholdUnsupported = true;
      admitted = false;
    } else if (value === 'invalid') {
      entry.thresholdInvalid = true;
      admitted = false;
    }
  }
  return admitted ? (values as number[]) : undefined;
}

export function compileCf(
  worksheet: Worksheet,
  cellIndex: ReadonlyMap<string, Cell> = createCellIndex(worksheet),
): CfContext {
  const compiled: CompiledCfRule[] = [];
  const definedNames = new Map<string, DefinedName>();
  for (const dn of worksheet.definedNames ?? []) {
    definedNames.set(dn.name, dn);
  }
  const blocks = worksheet.conditionalFormats ?? [];
  for (let blockIndex = 0; blockIndex < blocks.length; blockIndex++) {
    const cf = blocks[blockIndex];
    // New provenance-only blocks must not trigger cached-cell scans. Only
    // existing statistical/scale rules need range samples; reuse their work.
    const needsSamples = cf.rules.some(rule => rule.type !== 'expression' && rule.type !== 'cellIs' && rule.type !== 'other' && !(rule.type === 'dataBar' && rule.extThresholdFormula));
    const samples = needsSamples ? collectNumericValuesInRanges(worksheet, cf.sqref) : [];
    for (let ruleIndex = 0; ruleIndex < cf.rules.length; ruleIndex++) {
      const rule = cf.rules[ruleIndex];
      // Document position is captured here, before the priority sort below.
      const entry: CompiledCfRule = { rule, sqref: cf.sqref, blockIndex, ruleIndex };
      if (rule.type === 'colorScale') {
        entry.scaleStops = admitThresholds(entry, rule.stops.map(s => resolveCfvoValue(s, samples)));
      } else if (rule.type === 'dataBar') {
        if (rule.extThresholdFormula) {
          // This marker represents a refused x14 threshold, including a
          // formula whose body is deliberately not retained. Do not diagnose
          // that omitted body as a malformed standard formula or evaluate it.
          entry.thresholdUnsupported = true;
        } else {
          const bounds = admitThresholds(entry, [
            resolveCfvoValue(rule.min, samples),
            resolveCfvoValue(rule.max, samples),
          ]);
          if (bounds) {
            entry.barMin = bounds[0];
            entry.barMax = bounds[1];
          }
        }
      } else if (rule.type === 'top10') {
        const sorted = [...samples].sort((a, b) => a - b);
        const n = sorted.length;
        if (n > 0) {
          const rank = Math.min(rule.rank, n);
          if (rule.percent) {
            const p = rule.top ? (1 - rank / 100) : (rank / 100);
            const idx = Math.max(0, Math.min(n - 1, Math.round(p * (n - 1))));
            entry.top10Threshold = sorted[idx];
          } else {
            entry.top10Threshold = rule.top ? sorted[Math.max(0, n - rank)] : sorted[Math.min(n - 1, rank - 1)];
          }
          entry.top10IsTop = rule.top;
        }
      } else if (rule.type === 'aboveAverage') {
        if (samples.length > 0) {
          const mean = samples.reduce((a, b) => a + b, 0) / samples.length;
          entry.avgValue = mean;
          entry.avgIsAbove = rule.aboveAverage;
          if (rule.stdDev && rule.stdDev > 0) {
            // ECMA-376 §18.3.1.10: Excel uses the population standard
            // deviation (divide by N, not N-1) for stdDev bands.
            const variance =
              samples.reduce((a, b) => a + (b - mean) * (b - mean), 0) / samples.length;
            entry.avgStdDev = Math.sqrt(variance);
          }
        }
      } else if (rule.type === 'iconSet') {
        entry.iconThresholds = admitThresholds(entry, rule.cfvos.map(cfv => resolveCfvoValue(cfv, samples)));
      } else if (rule.type === 'cellIs') {
        entry.cellIsOperands = rule.formulas.map(cellIsOperand);
      }
      compiled.push(entry);
    }
  }
  // Excel evaluates CF rules in ascending priority (lowest number = highest
  // priority first). For each property (fill/fontColor/border/…) the first
  // matching rule wins, and `stopIfTrue` on a matching rule skips all later
  // rules for that cell (evaluateCf). Match that here by
  // iterating asc and only setting properties that are still unset.
  compiled.sort((a, b) => {
    const pa = (a.rule as { priority: number }).priority ?? 0;
    const pb = (b.rule as { priority: number }).priority ?? 0;
    return pa - pb;
  });
  return { compiled, worksheet, cellIndex, definedNames };
}

function createCellIndex(worksheet: Worksheet): Map<string, Cell> {
  return buildCellCoordinateIndex(worksheet.rows, {
    resource: 'worksheet-cell-index',
    operation: 'index-worksheet-cells',
    limit: getWorksheetPolicy(worksheet).maxCoordinateIndexEntries,
  });
}

/** ECMA-376 §18.18.15 ST_ConditionalFormattingOperator. */
const CELL_IS_OPERATORS: ReadonlySet<string> = new Set([
  'lessThan', 'lessThanOrEqual', 'equal', 'notEqual', 'greaterThanOrEqual',
  'greaterThan', 'between', 'notBetween', 'containsText', 'notContains',
  'beginsWith', 'endsWith',
]);

function cellIsMatch(num: number, operator: string, args: number[]): boolean {
  switch (operator) {
    case 'greaterThan': return num > (args[0] ?? 0);
    case 'greaterThanOrEqual': return num >= (args[0] ?? 0);
    case 'lessThan': return num < (args[0] ?? 0);
    case 'lessThanOrEqual': return num <= (args[0] ?? 0);
    case 'equal': return num === (args[0] ?? 0);
    case 'notEqual': return num !== (args[0] ?? 0);
    case 'between': return num >= (args[0] ?? 0) && num <= (args[1] ?? 0);
    case 'notBetween': return num < (args[0] ?? 0) || num > (args[1] ?? 0);
    default: return false;
  }
}

/** One `cellIs` operand ([MS-XLSX] 2.6.27: a formula, number or cell
 *  reference). A literal is decoded once per rule; anything else is looked
 *  up per cell by `decodeCfOperand` (a single-cell reference to a cached
 *  value) and is otherwise not decodable. `0+10` is never read as 0. */
type CellIsOperand = { literal: CfOperandValue } | { formula: string };

function cellIsOperand(f: string): CellIsOperand {
  const literal = decodeCfLiteral(f);
  return literal !== undefined ? { literal } : { formula: f };
}

function cellIsTextMatch(text: string, operator: string, args: string[]): boolean {
  const a = args[0] ?? '';
  const b = args[1] ?? '';
  const ci = (s: string) => s.toLowerCase();
  switch (operator) {
    case 'equal':         return ci(text) === ci(a);
    case 'notEqual':      return ci(text) !== ci(a);
    case 'containsText':  return ci(text).includes(ci(a));
    case 'notContains':   return !ci(text).includes(ci(a));
    case 'beginsWith':    return ci(text).startsWith(ci(a));
    case 'endsWith':      return ci(text).endsWith(ci(a));
    case 'between':       return ci(text) >= ci(a) && ci(text) <= ci(b);
    case 'notBetween':    return ci(text) <  ci(a) || ci(text) >  ci(b);
    default: return false;
  }
}

function interpolateHex(a: string, b: string, t: number): string {
  const pa = a.replace('#', '');
  const pb = b.replace('#', '');
  const ar = parseInt(pa.slice(0, 2), 16), ag = parseInt(pa.slice(2, 4), 16), ab = parseInt(pa.slice(4, 6), 16);
  const br = parseInt(pb.slice(0, 2), 16), bg = parseInt(pb.slice(2, 4), 16), bb = parseInt(pb.slice(4, 6), 16);
  const r = Math.round(ar + (br - ar) * t);
  const g = Math.round(ag + (bg - ag) * t);
  const bl = Math.round(ab + (bb - ab) * t);
  return `#${r.toString(16).padStart(2, '0').toUpperCase()}${g.toString(16).padStart(2, '0').toUpperCase()}${bl.toString(16).padStart(2, '0').toUpperCase()}`;
}

function colorScaleAt(num: number, stops: CfStop[], stopValues: number[]): string {
  if (!stops.length) return '#FFFFFF';
  if (num <= stopValues[0]) return stops[0].color;
  if (num >= stopValues[stopValues.length - 1]) return stops[stops.length - 1].color;
  for (let i = 1; i < stopValues.length; i++) {
    if (num <= stopValues[i]) {
      const lo = stopValues[i - 1];
      const hi = stopValues[i];
      const t = hi === lo ? 0 : (num - lo) / (hi - lo);
      return interpolateHex(stops[i - 1].color, stops[i].color, t);
    }
  }
  return stops[stops.length - 1].color;
}

function applyDxfToResult(result: CfResult, dxf: Dxf | null | undefined): void {
  if (!dxf) return;
  // First-match-wins (higher priority) for each property. See compileCf.
  // Per ECMA-376 §18.3.1.11, a `<dxf>` is a *differential* format: any child
  // element it contains is an override of the base cell format. So the mere
  // presence of `dxf.fill` means "replace the base fill with this", whatever
  // its patternType / color — including `patternType="none"` (explicit clear)
  // and gradient fills. The paint-site guard (`patternType !== 'none' &&
  // fgColor`) handles whether the result actually paints a color or leaves
  // the cell transparent, so this override stays spec-faithful without
  // second-guessing the fill's shape here.
  if (dxf.fill && !result.fill) result.fill = dxf.fill;
  if (dxf.font?.color && result.fontColor == null) result.fontColor = dxf.font.color;
  // A font toggle is tri-state: an explicit off (`<b val="0"/>`) is a
  // property the rule sets, so it both claims the property against
  // lower-priority rules and turns off bold the cell formatting turns on;
  // an omitted element leaves both alone (§18.8.14-15, §18.8.2).
  result.fontBold ??= dxfFontToggle(dxf, 'bold');
  result.fontItalic ??= dxfFontToggle(dxf, 'italic');
  result.fontUnderline ??= dxfFontToggle(dxf, 'underline');
  result.fontStrike ??= dxfFontToggle(dxf, 'strike');
  if (dxf.numFmt && result.numFmt == null) {
    result.numFmt = {
      numFmtId: dxf.numFmt.numFmtId,
      formatCode: dxf.numFmt.formatCode || null,
    };
  }
  if (dxf.border) {
    // Merge per-edge — higher-priority edges stay; lower-priority edges fill
    // in unset ones. dxf `border` typically sets only the edges the rule
    // cares about (e.g. left+right for a "today" column marker).
    const existing = result.border ?? {} as Border;
    const merged: Border = {
      left:         existing.left         ?? dxf.border.left,
      right:        existing.right        ?? dxf.border.right,
      top:          existing.top          ?? dxf.border.top,
      bottom:       existing.bottom       ?? dxf.border.bottom,
      diagonalUp:   existing.diagonalUp   ?? dxf.border.diagonalUp,
      diagonalDown: existing.diagonalDown ?? dxf.border.diagonalDown,
    };
    result.border = merged;
  }
}

/**
 * The activity condition of a colorScale / dataBar / iconSet rule: its
 * optional formula ([MS-XLSX] 2.6.27 CT_CfRule: "When the formula returns
 * zero, conditional formatting is not displayed. When the formula returns a
 * nonzero value, or is not present, conditional formatting is displayed").
 * Excel reads the SpreadsheetML `<formula>` of these types the same way (see
 * `CfRule`). Relative references anchor at the top-left of the rule's range,
 * as for `expression`. Formulas are not evaluated: only a literal or a
 * single-cell reference to a cached value is decoded
 * (`decodeCfOperandResult`).
 *
 * Library policy (#1547): a condition outside that decoder (or a linked x14
 * activity formula) is `unsupported` and is
 * reported; a cached Excel error is inactive. Numeric-result policy: only a
 * number is established as zero / nonzero; a logical, text or blank result is
 * not converted and stays inactive (the existing matching, not a claim of
 * Excel's conversion). An inactive or unsupported rule neither formats the
 * cell nor stops lower rules.
 */
function scaleRuleActivity(
  formula: string | undefined,
  entry: CompiledCfRule,
  row: number,
  col: number,
  cfCtx: CfContext,
): 'active' | 'inactive' | 'unsupported' {
  if (formula == null) return 'active';
  const anchor = entry.sqref[0];
  if (!anchor) return 'inactive';
  const decoded = decodeCfOperandResult(formula, {
    row, col,
    anchorRow: anchor.top, anchorCol: anchor.left,
    cellIndex: cfCtx.cellIndex,
  });
  if (decoded.kind === 'unsupported') return 'unsupported';
  if (decoded.kind === 'error') return 'inactive';
  return typeof decoded.value === 'number' && decoded.value !== 0 ? 'active' : 'inactive';
}

function hasThresholdFailure(entry: CompiledCfRule): boolean {
  return entry.thresholdUnsupported === true || entry.thresholdInvalid === true;
}

/** Whether a scale rule that reached a numeric cell formats it: its activity
 *  holds and its thresholds were admitted. A refused rule is reported and
 *  neither paints nor stops. Threshold failures are reported only after the
 *  activity is established, so an inactive rule makes no threshold claim. */
function admitScaleRule(
  entry: CompiledCfRule,
  formula: string | undefined,
  row: number,
  col: number,
  cfCtx: CfContext,
  diagnostics: CfDiagnosticCollector | undefined,
): boolean {
  const activity = scaleRuleActivity(formula, entry, row, col, cfCtx);
  if (activity === 'unsupported') {
    diagnostics?.record(entry, 'activity', 'unsupported', row, col);
    return false;
  }
  if (activity === 'inactive') return false;
  if (!hasThresholdFailure(entry)) return true;
  if (entry.thresholdUnsupported) diagnostics?.record(entry, 'threshold', 'unsupported', row, col);
  if (entry.thresholdInvalid) diagnostics?.record(entry, 'threshold', 'invalid', row, col);
  return false;
}

/**
 * Evaluate every rule covering one cell in priority order. `diagnostics` is
 * the operation-local collector of the render that owns this call; the
 * evaluation result is unchanged with or without it. Only rules this call
 * actually reaches are reported; rules after a matching stopIfTrue rule are
 * not reached and make no claim.
 */
export function evaluateCf(
  cell: Cell | undefined,
  row: number,
  col: number,
  cfCtx: CfContext,
  dxfs: Dxf[],
  diagnostics?: CfDiagnosticCollector,
): CfResult {
  const result: CfResult = {};
  if (!cfCtx.compiled.length) return result;
  for (const entry of cfCtx.compiled) {
    if (!rangeContains(entry.sqref, row, col)) continue;
    const rule = entry.rule;
    const numVal = cellNumericValue(cell);

    // Each evaluated rule decides whether it matched this cell; a matched
    // rule applies its formatting and then honours `stopIfTrue` (§18.3.1.10:
    // "no rules with lower priority shall be applied over this rule, when
    // this rule evaluates to true"). The stop is per cell and does not
    // depend on whether the matched rule and the skipped ones touch the same
    // properties, so a lower-priority rule's explicit font-toggle off cannot
    // erase formatting beneath a stopping rule that sets only a colour.
    let matched = false;
    if (rule.type === 'expression') {
      const anchor = entry.sqref[0];
      if (!anchor) continue;
      // One evaluation, reused for both the match and the diagnostic (TODAY
      // / NOW stay consistent). A valid FALSE / 0 and an Excel error value
      // are normal no-matches; only unsupported / invalid are reported.
      const evaluation = evaluateFormula(rule.formula, {
        row, col,
        anchorRow: anchor.top, anchorCol: anchor.left,
        cellIndex: cfCtx.cellIndex,
        definedNames: cfCtx.definedNames,
        depth: 0,
      });
      if (evaluation.kind === 'unsupported' || evaluation.kind === 'invalid') {
        diagnostics?.record(entry, 'expression', evaluation.kind, row, col);
      }
      matched = formulaEvaluationToBool(evaluation);
      if (matched) applyDxfToResult(result, rule.dxfId != null ? dxfs[rule.dxfId] : null);
    } else if (rule.type === 'cellIs') {
      const anchor = entry.sqref[0];
      if (!anchor) continue;
      const cellIsOperands = entry.cellIsOperands ?? [];
      // §18.3.1.10: between/notBetween take two formulas, every other
      // operator one; the operator is ST_ConditionalFormattingOperator
      // (§18.18.15). An unknown operator or a missing operand is a
      // structurally invalid rule: reported, no match, no stop.
      const arity = rule.operator === 'between' || rule.operator === 'notBetween' ? 2 : 1;
      if (!CELL_IS_OPERATORS.has(rule.operator) || cellIsOperands.length < arity) {
        diagnostics?.record(entry, 'cellIs', 'invalid', row, col);
      } else {
        // Compare only with decoded operands. A cached Excel error, or an
        // operand whose type differs from the cell's (blank and logical
        // included), is a normal no-match. An operand outside the decoder's
        // literal / single-cell-reference grammar, or a relative reference
        // shifted past the admitted grid edge, is a library
        // limitation: reported, no match, no stop.
        const operands: CfOperandValue[] = [];
        let decoded = true;
        for (const operand of cellIsOperands) {
          if ('literal' in operand) {
            operands.push(operand.literal);
            continue;
          }
          const value = decodeCfOperandResult(operand.formula, {
            row, col,
            anchorRow: anchor.top, anchorCol: anchor.left,
            cellIndex: cfCtx.cellIndex,
          });
          if (value.kind === 'value') {
            operands.push(value.value);
            continue;
          }
          if (value.kind === 'unsupported') diagnostics?.record(entry, 'cellIs', 'unsupported', row, col);
          decoded = false;
          break;
        }
        const textVal = cellTextValue(cell);
        if (!decoded) {
          // no match
        } else if (numVal != null && operands.every(a => typeof a === 'number')) {
          matched = cellIsMatch(numVal, rule.operator, operands as number[]);
        } else if (textVal != null && operands.every(a => typeof a === 'string')) {
          matched = cellIsTextMatch(textVal, rule.operator, operands as string[]);
        }
      }
      if (matched) applyDxfToResult(result, rule.dxfId != null ? dxfs[rule.dxfId] : null);
    } else if (rule.type === 'other') {
      // Parser-preserved unimplemented formula inlets: only known fixed bits
      // at a reached visible cell; no match, decoration or stopIfTrue.
      for (const [bit, phase] of OTHER_FORMULA_PHASES) {
        if (((rule.unsupportedFormulaPhases ?? 0) & bit) !== 0) {
          diagnostics?.record(entry, phase, 'unsupported', row, col);
        }
      }
    } else if (rule.type === 'top10') {
      if (numVal == null || entry.top10Threshold == null) continue;
      matched = entry.top10IsTop ? numVal >= entry.top10Threshold : numVal <= entry.top10Threshold;
      if (matched) applyDxfToResult(result, rule.dxfId != null ? dxfs[rule.dxfId] : null);
    } else if (rule.type === 'aboveAverage') {
      if (numVal == null || entry.avgValue == null) continue;
      // ECMA-376 §18.3.1.10: with `stdDev=N` the threshold is mean ± N·σ
      // (population σ); otherwise it is the plain mean. `equalAverage`
      // includes cells exactly at the threshold in the highlighted set.
      const band = entry.avgStdDev != null ? entry.avgStdDev * (rule.stdDev ?? 1) : 0;
      const threshold = entry.avgIsAbove ? entry.avgValue + band : entry.avgValue - band;
      const eq = rule.equalAverage === true;
      matched = entry.avgIsAbove
        ? (eq ? numVal >= threshold : numVal > threshold)
        : (eq ? numVal <= threshold : numVal < threshold);
      if (matched) applyDxfToResult(result, rule.dxfId != null ? dxfs[rule.dxfId] : null);
    } else if (rule.type === 'iconSet') {
      // A scale rule "evaluates to true" for every numeric cell it formats
      // while its activity condition holds and its thresholds are admitted
      // (admitScaleRule).
      if (numVal == null || (!hasThresholdFailure(entry) && !entry.iconThresholds?.length)) continue;
      if (!admitScaleRule(entry, rule.activeFormula, row, col, cfCtx, diagnostics)) continue;
      const thresholds = entry.iconThresholds;
      if (!thresholds?.length) continue;
      matched = true;
      const n = thresholds.length;
      let iconIdx = 0;
      for (let i = 1; i < n; i++) {
        if (numVal >= thresholds[i]) iconIdx = i;
      }
      if (rule.reverse) iconIdx = n - 1 - iconIdx;
      // Custom iconSets (Excel 2010+ x14 extension) override per-threshold icons.
      if (rule.customIcons && rule.customIcons[iconIdx]) {
        const ci = rule.customIcons[iconIdx];
        if (ci.iconSet !== 'NoIcons') {
          result.iconSet = { name: ci.iconSet, index: ci.iconId };
        }
      } else {
        result.iconSet = { name: rule.iconSet, index: iconIdx };
      }
    } else if (rule.type === 'colorScale') {
      if (numVal == null || (!hasThresholdFailure(entry) && !entry.scaleStops)) continue;
      if (!admitScaleRule(entry, rule.activeFormula, row, col, cfCtx, diagnostics)) continue;
      const scaleStops = entry.scaleStops;
      if (!scaleStops) continue;
      matched = true;
      if (!result.fill) {
        const color = colorScaleAt(numVal, rule.stops, scaleStops);
        result.fill = { patternType: 'solid', fgColor: color, bgColor: color };
      }
    } else if (rule.type === 'dataBar') {
      if (numVal == null
        || (!hasThresholdFailure(entry) && (entry.barMin == null || entry.barMax == null))) continue;
      if (!admitScaleRule(entry, rule.activeFormula, row, col, cfCtx, diagnostics)) continue;
      const barMin = entry.barMin;
      const barMax = entry.barMax;
      if (barMin == null || barMax == null) continue;
      matched = true;
      if (!result.dataBar) {
        const range = barMax - barMin;
        const ratio = range === 0 ? 0 : Math.max(0, Math.min(1, (numVal - barMin) / range));
        result.dataBar = { color: rule.color, ratio, gradient: rule.gradient };
      }
    }
    // `other` kinds (timePeriod, duplicateValues, uniqueValues, …) are not
    // evaluated yet: an unevaluated rule never matches, so it neither
    // formats the cell nor stops the rules after it.
    //
    // The stop applies to colorScale / dataBar / iconSet too. Excel's rule
    // editor does not offer the flag for them and Office's binary storage
    // requires it to be 0 ([MS-XLSB] 2.4.23 BrtBeginCFRule `fStopTrue`,
    // [MS-XLS] 2.4.43 CF12), but ECMA-376 places no type restriction on it,
    // and Excel for Mac's PDF export of a control workbook honours a set
    // flag in SpreadsheetML: a colorScale, dataBar or iconSet rule with
    // stopIfTrue="1" kept a lower-priority bold+underline `expression` rule
    // off every numeric cell it formatted, and the same rules without the
    // flag let it apply. Non-numeric cells in a scale range were not tested.
    if (matched && rule.stopIfTrue) break;
  }
  return result;
}


// ────────────────────────────────────────────────────────────────
// Shared state for a single renderViewport call
// ────────────────────────────────────────────────────────────────

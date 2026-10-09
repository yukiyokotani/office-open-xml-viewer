import { describe, expect, it } from 'vitest';
import { compileCf, evaluateCf } from './conditional-format.js';
import {
  CfDiagnosticCollector,
  decodeCfDiagnosticsWire,
  type CfDiagnosticKind,
  type CfDiagnosticPhase,
  type CfRuleDiagnostic,
} from './cf-diagnostics.js';
import type {
  Cell, CellValue, CfRule, CfValue, ConditionalFormat, Dxf, Worksheet, WorksheetCellRange,
} from './types.js';

// Shared CF evaluator used by every XLSX consumer (main renderer, render
// workers). The Node facade has no XLSX render API, so this is the entry.

const RED: Dxf = { font: null, fill: { patternType: 'solid', fgColor: '#FF0000', bgColor: '#FF0000' }, border: null };
const BLUE: Dxf = {
  font: { bold: false, italic: false, underline: false, strike: false, size: 11, color: '#0000FF', name: null },
  fill: null, border: null,
};
const DXFS = [RED, BLUE];

const at = (row: number, col: number, value: CellValue): Cell => ({ row, col, value, styleIndex: 0 });
const num = (row: number, col: number, n: number): Cell => at(row, col, { type: 'number', number: n });
const err = (row: number, col: number, e: string): Cell => at(row, col, { type: 'error', error: e });

function sheet(cells: Cell[], conditionalFormats: ConditionalFormat[]): Worksheet {
  const rows = new Map<number, Cell[]>();
  for (const c of cells) rows.set(c.row, [...(rows.get(c.row) ?? []), c]);
  return {
    name: 'Sheet1',
    rows: [...rows].sort(([a], [b]) => a - b).map(([index, rowCells]) => ({ index, height: null, cells: rowCells })),
    colWidths: {}, rowHeights: {}, defaultColWidth: 8.43, defaultRowHeight: 15,
    mergeCells: [], freezeRows: 0, freezeCols: 0, conditionalFormats, images: [], charts: [],
  };
}

const C1: WorksheetCellRange = { top: 1, left: 3, bottom: 1, right: 3 };
const A1_A3: WorksheetCellRange = { top: 1, left: 1, bottom: 3, right: 1 };
const lowerRed: CfRule = { type: 'expression', formula: 'TRUE', dxfId: 0, priority: 9, stopIfTrue: false };

function evaluate(ws: Worksheet, row: number, col: number) {
  const collector = new CfDiagnosticCollector();
  const target = ws.rows.find((r) => r.index === row)?.cells.find((c) => c.col === col);
  const result = evaluateCf(target, row, col, compileCf(ws), DXFS, collector);
  return { result, diagnostics: collector.snapshot() };
}

const diag = (
  phase: CfDiagnosticPhase, kind: CfDiagnosticKind, ruleIndex: number,
  row = 1, col = 3, blockIndex = 0,
): CfRuleDiagnostic => ({ kind, phase, blockIndex, ruleIndex, row, col });

describe('parser-preserved unimplemented formula inlets', () => {
  const marker = (mask: number): CfRule => ({ type: 'other', kind: 'expression',
    priority: 1, stopIfTrue: true, unsupportedFormulaPhases: mask } as CfRule);
  it('new diagnostic-only blocks do not rescan a warm worksheet cell index', () => {
    const ws = sheet([], [{sqref:[C1],rules:[marker(1)]}]);
    Object.defineProperty(ws, 'rows', {get: () => {throw new Error('unexpected cached-cell rescan');}});
    const collector = new CfDiagnosticCollector();
    const compiled = compileCf(ws, new Map());
    evaluateCf(undefined, 1, 3, compiled, DXFS, collector);
    expect(collector.snapshot()).toEqual([diag('expression','unsupported',0)]);
  });
  it('reports only the fixed four provenance phases, without formatting or stopping', () => {
    const ws = sheet([num(1, 3, 5)], [{ sqref: [C1], rules: [marker(15), lowerRed] }]);
    const { result, diagnostics } = evaluate(ws, 1, 3);
    expect(result.fill?.fgColor).toBe('#FF0000');
    expect(diagnostics).toEqual(['expression', 'cellIs', 'activity', 'threshold']
      .map(phase => diag(phase as CfDiagnosticPhase, 'unsupported', 0)));
  });
  it('deduplicates repeated cells at one admitted rule rather than retaining per-cell records', () => {
    const ws = sheet([num(1, 1, 5)], [{sqref: [A1_A3], rules: [marker(1)]}]);
    const compiled = compileCf(ws); const collector = new CfDiagnosticCollector();
    for (let row = 1; row <= 3; row++) evaluateCf(undefined, row, 1, compiled, DXFS, collector);
    expect(collector.snapshot()).toEqual([diag('expression', 'unsupported', 0, 1, 1)]);
  });
  it('makes no claim outside range, after a stopping rule, or from unknown mask bits', () => {
    const ws = sheet([num(1, 3, 5)], [{sqref: [C1], rules: [marker(1)]}]);
    expect(evaluate(ws, 1, 2).diagnostics).toEqual([]);
    expect(evaluate(sheet([num(1,3,5)], [{sqref:[C1], rules:[
      {...lowerRed,priority:0,stopIfTrue:true}, marker(1)]}]), 1, 3).diagnostics).toEqual([]);
    expect(evaluate(sheet([num(1,3,5)], [{sqref:[C1],rules:[marker(16)]}]),1,3).diagnostics).toEqual([]);
  });
});

describe('expression ingress', () => {
  // Baseline probe: mathematically TRUE (1*1+2*1>0), but SUMPRODUCT is not in
  // the evaluator's function table. Before #1547 this was a silent no-match.
  const sumproductCells = [num(1, 1, 1), num(2, 1, 2), num(1, 2, 1), num(2, 2, 1), num(1, 3, 5)];
  const sumproduct: CfRule = {
    type: 'expression', formula: 'SUMPRODUCT(A1:A2,B1:B2)>0', dxfId: 1, priority: 1, stopIfTrue: true,
  };

  it('reports an unsupported SUMPRODUCT, keeps the lower rule and does not stop', () => {
    // Document order lowerRed, sumproduct: ruleIndex is captured before the
    // priority sort, so the report names ruleIndex 1.
    const ws = sheet(sumproductCells, [{ sqref: [C1], rules: [lowerRed, sumproduct] }]);
    const before = structuredClone(ws.rows);
    const { result, diagnostics } = evaluate(ws, 1, 3);
    expect(result.fontColor).toBeUndefined();
    expect(result.fill?.fgColor).toBe('#FF0000');
    expect(diagnostics).toEqual([diag('expression', 'unsupported', 1)]);
    // Cached values and formulas are never recalculated or mutated.
    expect(ws.rows).toEqual(before);
  });

  it('keeps valid FALSE / zero / Excel error results silent', () => {
    for (const formula of ['FALSE', '0', '#N/A=1', '1/0>0', 'A1>100']) {
      const rule: CfRule = { type: 'expression', formula, dxfId: 1, priority: 1, stopIfTrue: true };
      const { result, diagnostics } = evaluate(sheet(sumproductCells, [{ sqref: [C1], rules: [rule, lowerRed] }]), 1, 3);
      expect(diagnostics, formula).toEqual([]);
      expect(result.fill?.fgColor, formula).toBe('#FF0000');
    }
  });

  it('reports invalid syntax as invalid', () => {
    const rule: CfRule = { type: 'expression', formula: '1+', dxfId: 1, priority: 1, stopIfTrue: true };
    const { diagnostics } = evaluate(sheet([num(1, 3, 5)], [{ sqref: [C1], rules: [rule] }]), 1, 3);
    expect(diagnostics).toEqual([diag('expression', 'invalid', 0)]);
  });

  it('a supported stopping rule prevents evaluation (and report) of lower rules', () => {
    const stop: CfRule = { type: 'expression', formula: 'TRUE', dxfId: 1, priority: 1, stopIfTrue: true };
    const lower = { ...sumproduct, priority: 2, stopIfTrue: false };
    const { result, diagnostics } = evaluate(sheet(sumproductCells, [{ sqref: [C1], rules: [stop, lower] }]), 1, 3);
    expect(result.fontColor).toBe('#0000FF');
    expect(diagnostics).toEqual([]);
  });

  it('classifies data-dependent limitations at runtime, not statically', () => {
    // §18.17.7.278 does not establish fractional number-digits: unsupported
    // only when the cached cell actually holds a fraction.
    const rule: CfRule = { type: 'expression', formula: 'ROUND(1.25,$D$1)=1.3', dxfId: 1, priority: 1, stopIfTrue: false };
    const whole = evaluate(sheet([num(1, 3, 5), num(1, 4, 1)], [{ sqref: [C1], rules: [rule] }]), 1, 3);
    expect(whole.diagnostics).toEqual([]);
    expect(whole.result.fontColor).toBe('#0000FF');
    const fractional = evaluate(sheet([num(1, 3, 5), num(1, 4, 1.5)], [{ sqref: [C1], rules: [rule] }]), 1, 3);
    expect(fractional.diagnostics).toEqual([diag('expression', 'unsupported', 0)]);
    expect(fractional.result.fontColor).toBeUndefined();
  });
});

describe('cellIs ingress', () => {
  const cellIs = (operator: string, formulas: string[]): CfRule =>
    ({ type: 'cellIs', operator, formulas, dxfId: 1, priority: 1, stopIfTrue: true });
  const run = (rule: CfRule, extra: Cell[] = []) =>
    evaluate(sheet([num(1, 3, 5), ...extra], [{ sqref: [C1], rules: [rule, lowerRed] }]), 1, 3);

  it('reports an operand outside the literal/reference grammar', () => {
    const { result, diagnostics } = run(cellIs('greaterThan', ['A1+1']), [num(1, 1, 1)]);
    expect(diagnostics).toEqual([diag('cellIs', 'unsupported', 0)]);
    expect(result.fill?.fgColor).toBe('#FF0000');
    expect(result.fontColor).toBeUndefined();
  });

  it('keeps a cached Excel error, a blank and a type mismatch silent', () => {
    for (const extra of [[err(1, 2, '#N/A')], [], [at(1, 2, { type: 'bool', bool: false })], [at(1, 2, { type: 'text', text: 'x' })]]) {
      const { result, diagnostics } = run(cellIs('greaterThan', ['$B$1']), extra);
      expect(diagnostics).toEqual([]);
      expect(result.fill?.fgColor).toBe('#FF0000');
    }
  });

  it('literal and reference controls are supported', () => {
    expect(run(cellIs('greaterThan', ['4']))).toMatchObject({ diagnostics: [], result: { fontColor: '#0000FF' } });
    expect(run(cellIs('greaterThan', ['$B$1']), [num(1, 2, 3)])).toMatchObject({ diagnostics: [], result: { fontColor: '#0000FF' } });
  });

  it('reports a missing operand or an unknown operator as invalid', () => {
    expect(run(cellIs('between', ['1'])).diagnostics).toEqual([diag('cellIs', 'invalid', 0)]);
    expect(run(cellIs('greaterThan', [])).diagnostics).toEqual([diag('cellIs', 'invalid', 0)]);
    expect(run(cellIs('greaterish', ['1'])).diagnostics).toEqual([diag('cellIs', 'invalid', 0)]);
  });
});

describe('scale-rule activity ingress', () => {
  const bar = (activeFormula?: string, extra: Partial<Extract<CfRule, { type: 'dataBar' }>> = {}): CfRule => ({
    type: 'dataBar', color: '#638EC6', min: { kind: 'min', value: null }, max: { kind: 'max', value: null },
    priority: 1, gradient: true, activeFormula, stopIfTrue: true, ...extra,
  });
  const run = (rule: CfRule, extra: Cell[] = []) =>
    evaluate(sheet([num(1, 1, 1), num(2, 1, 5), num(3, 1, 9), ...extra], [{ sqref: [A1_A3], rules: [rule, lowerRed] }]), 2, 1);

  it('reports an unsupported activity formula and neither paints nor stops', () => {
    const { result, diagnostics } = run(bar('SUM(A1:A2)'));
    expect(diagnostics).toEqual([diag('activity', 'unsupported', 0, 2, 1)]);
    expect(result.dataBar).toBeUndefined();
    expect(result.fill?.fgColor).toBe('#FF0000');
  });

  it('absent is active, valid zero is inactive, cached error is inactive; all silent', () => {
    expect(run(bar())).toMatchObject({ diagnostics: [] });
    expect(run(bar()).result.dataBar).toBeDefined();
    expect(run(bar('0')).result.dataBar).toBeUndefined();
    expect(run(bar('0')).diagnostics).toEqual([]);
    const errored = run(bar('$B$1'), [err(1, 2, '#DIV/0!')]);
    expect(errored.diagnostics).toEqual([]);
    expect(errored.result.dataBar).toBeUndefined();
  });

  it('an effective linked activity formula outside the decoder is reported', () => {
    const { result, diagnostics } = run(bar('SUMPRODUCT(A1:A3)'));
    expect(diagnostics).toEqual([diag('activity', 'unsupported', 0, 2, 1)]);
    expect(result.dataBar).toBeUndefined();
  });
});

describe('cfvo threshold ingress', () => {
  const cells = [num(1, 1, 1), num(2, 1, 5), num(3, 1, 9)];
  const run = (rule: CfRule) => evaluate(sheet(cells, [{ sqref: [A1_A3], rules: [rule, lowerRed] }]), 2, 1);
  it('does not convert automatic or unknown threshold kinds into zero', () => {
    for (const kind of ['autoMin', 'autoMax', 'futureKind']) {
      const {result, diagnostics} = run({type:'dataBar', priority:1, stopIfTrue:true,
        color:'#638EC6', gradient:true, min:{kind,value:null}, max:{kind:'max',value:null}});
      expect(result.dataBar).toBeUndefined();
      expect(result.fill?.fgColor).toBe('#FF0000');
      expect(diagnostics).toEqual([diag('threshold','unsupported',0,2,1)]);
    }
  });
  it('a parser-marked x14 literal formula is not a standard literal threshold', () => {
    const {result, diagnostics} = run({type:'dataBar', priority:1, stopIfTrue:true,
      color:'#638EC6', gradient:true, extThresholdFormula:true,
      min:{kind:'formula',value:null},max:{kind:'num',value:'10'}});
    expect(result.dataBar).toBeUndefined();
    expect(result.fill?.fgColor).toBe('#FF0000');
    expect(diagnostics).toEqual([diag('threshold','unsupported',0,2,1)]);
  });
  const scale = (first: CfValue): CfRule => ({
    type: 'colorScale', priority: 1, stopIfTrue: true,
    stops: [{ ...first, color: '#000000' }, { kind: 'max', value: null, color: '#FFFFFF' }],
  });
  const bar = (min: CfValue): CfRule => ({
    type: 'dataBar', color: '#638EC6', min, max: { kind: 'max', value: null }, priority: 1, gradient: true, stopIfTrue: true,
  });
  const icons = (middle: CfValue): CfRule => ({
    type: 'iconSet', iconSet: '3Arrows', reverse: false, priority: 1, stopIfTrue: true,
    cfvos: [{ kind: 'percent', value: '0' }, middle, { kind: 'percent', value: '67' }],
  });

  it('formula thresholds (including the numeric-prefix 1+1) are unsupported in all three types', () => {
    for (const value of ['$A$1', '1+1', 'MAX(A1:A3)']) {
      const formula: CfValue = { kind: 'formula', value };
      const colorScale = run(scale(formula));
      expect(colorScale.diagnostics, value).toEqual([diag('threshold', 'unsupported', 0, 2, 1)]);
      expect(colorScale.result.fill?.fgColor, value).toBe('#FF0000'); // no guessed scale; lower rule paints
      const dataBar = run(bar(formula));
      expect(dataBar.diagnostics, value).toEqual([diag('threshold', 'unsupported', 0, 2, 1)]);
      expect(dataBar.result.dataBar, value).toBeUndefined();
      const iconSet = run(icons(formula));
      expect(iconSet.diagnostics, value).toEqual([diag('threshold', 'unsupported', 0, 2, 1)]);
      expect(iconSet.result.iconSet, value).toBeUndefined();
    }
    // A num threshold written as a formula is not read as its numeric prefix.
    expect(run(bar({ kind: 'num', value: '1+1' })).diagnostics).toEqual([diag('threshold', 'unsupported', 0, 2, 1)]);
  });

  it('a formula threshold without text is invalid', () => {
    expect(run(bar({ kind: 'formula', value: null })).diagnostics).toEqual([diag('threshold', 'invalid', 0, 2, 1)]);
  });

  it('min / max / num / percent / literal formula controls stay valid and paint', () => {
    expect(run(scale({ kind: 'min', value: null }))).toMatchObject({ diagnostics: [] });
    expect(run(scale({ kind: 'min', value: null })).result.fill?.fgColor).not.toBe('#FF0000');
    expect(run(bar({ kind: 'num', value: '2' })).result.dataBar).toBeDefined();
    expect(run(bar({ kind: 'num', value: '2' })).diagnostics).toEqual([]);
    expect(run(icons({ kind: 'percent', value: '33' })).result.iconSet).toBeDefined();
    expect(run(icons({ kind: 'formula', value: '5' })).result.iconSet).toBeDefined();
    expect(run(icons({ kind: 'formula', value: '5' })).diagnostics).toEqual([]);
  });

  it('a linked x14 threshold formula marker skips and reports the rule', () => {
    const rule: CfRule = { ...(bar({ kind: 'min', value: null }) as Extract<CfRule, { type: 'dataBar' }>), extThresholdFormula: true };
    const { result, diagnostics } = run(rule);
    expect(diagnostics).toEqual([diag('threshold', 'unsupported', 0, 2, 1)]);
    expect(result.dataBar).toBeUndefined();
  });

  it('an inactive rule makes no threshold claim', () => {
    const rule: CfRule = { ...(bar({ kind: 'formula', value: '$A$1' }) as Extract<CfRule, { type: 'dataBar' }>), activeFormula: '0' };
    expect(run(rule).diagnostics).toEqual([]);
  });
});

describe('collector bounds and wire', () => {
  it('repeated viewport cells do not grow the report beyond one record per rule/phase/kind', () => {
    const cells = Array.from({ length: 200 }, (_, i) => num(i + 1, 1, i));
    const ws = sheet(cells, [{
      sqref: [{ top: 1, left: 1, bottom: 200, right: 1 }],
      rules: [
        { type: 'expression', formula: 'SUMPRODUCT(A1,A1)>0', dxfId: 0, priority: 1, stopIfTrue: false },
        { type: 'cellIs', operator: 'greaterThan', formulas: ['A1+1'], dxfId: 0, priority: 2 },
      ],
    }]);
    const ctx = compileCf(ws);
    const collector = new CfDiagnosticCollector();
    for (const row of ws.rows) {
      for (const c of row.cells) evaluateCf(c, c.row, c.col, ctx, DXFS, collector);
      for (const c of row.cells) evaluateCf(c, c.row, c.col, ctx, DXFS, collector); // repeated paint
    }
    expect(collector.snapshot()).toEqual([
      diag('expression', 'unsupported', 0, 1, 1),
      diag('cellIs', 'unsupported', 1, 1, 1),
    ]);
  });

  it('keeps distinct kinds of one rule and phase', () => {
    const rule: CfRule = {
      type: 'colorScale', priority: 1,
      stops: [{ kind: 'formula', value: null, color: '#000000' }, { kind: 'num', value: '$A$1', color: '#FFFFFF' }],
    };
    const { diagnostics } = evaluate(sheet([num(1, 1, 1), num(2, 1, 5)], [{ sqref: [A1_A3], rules: [rule] }]), 2, 1);
    expect(diagnostics).toEqual([diag('threshold', 'unsupported', 0, 2, 1), diag('threshold', 'invalid', 0, 2, 1)]);
  });

  it('evaluation without a collector is unchanged', () => {
    const ws = sheet([num(1, 3, 5)], [{ sqref: [C1], rules: [
      { type: 'expression', formula: 'SUMPRODUCT(A1,A1)>0', dxfId: 1, priority: 1, stopIfTrue: true }, lowerRed,
    ] }]);
    expect(evaluateCf(ws.rows[0].cells[0], 1, 3, compileCf(ws), DXFS).fill?.fgColor).toBe('#FF0000');
  });

  it('validates the structured-clone wire batch', () => {
    const collector = new CfDiagnosticCollector();
    collector.record({ blockIndex: 1, ruleIndex: 0 }, 'cellIs', 'invalid', 4, 2);
    collector.record({ blockIndex: 0, ruleIndex: 2 }, 'expression', 'unsupported', 1, 1);
    const wire = structuredClone(collector.snapshot());
    expect(decodeCfDiagnosticsWire(wire)).toEqual(collector.snapshot());
    expect(Object.isFrozen(decodeCfDiagnosticsWire(wire))).toBe(true);
    expect(() => decodeCfDiagnosticsWire([...wire, wire[0]])).toThrow(TypeError);
    expect(() => decodeCfDiagnosticsWire([{ ...wire[0], phase: 'other' }])).toThrow(TypeError);
    expect(() => decodeCfDiagnosticsWire([{ ...wire[0], row: 0 }])).toThrow(TypeError);
    expect(() => decodeCfDiagnosticsWire(undefined)).toThrow(TypeError);
  });
});

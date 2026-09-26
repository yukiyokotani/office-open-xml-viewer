import { describe, it, expect } from 'vitest';
import { compileCf, evaluateCf } from './conditional-format.js';
import type {
  Worksheet,
  Cell,
  CellValue,
  ConditionalFormat,
  CfRule,
  Dxf,
} from './types.js';

// ────────────────────────────────────────────────────────────────
// Conditional formatting — unit tests for the pure rule-evaluation
// layer (compileCf / evaluateCf).
//
// These exercise the spec logic (ECMA-376 §18.3.1.10 `<cfRule>`,
// §18.3.1.11 `<dxf>`) directly, independent of the WASM parser and the
// Canvas renderer:
//   - cellIs       — numeric & text comparison operators
//   - colorScale   — 2-color / 3-color gradient interpolation, min==max
//   - dataBar      — bar ratio clamping, min==max
//   - top10        — rank & percent thresholds, top vs bottom
//   - aboveAverage — mean, equalAverage boundary, stdDev bands
//   - iconSet      — threshold→icon index assignment, reverse
// ────────────────────────────────────────────────────────────────

/** Build a numeric cell. */
function numCell(row: number, col: number, n: number): Cell {
  return {
    col,
    row,
    value: { type: 'number', number: n } as CellValue,
    styleIndex: 0,
  };
}

/** Build a text cell. */
function textCell(row: number, col: number, t: string): Cell {
  return {
    col,
    row,
    value: { type: 'text', text: t } as CellValue,
    styleIndex: 0,
  };
}

/**
 * Build a single-column worksheet from a list of numbers (col 0, rows 0..n-1)
 * plus a list of conditional-format blocks covering that column.
 */
function sheetFromColumn(
  values: number[],
  cfs: ConditionalFormat[],
): Worksheet {
  return {
    name: 'Sheet1',
    rows: values.map((n, i) => ({ index: i, height: null, cells: [numCell(i, 0, n)] })),
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: cfs,
    images: [],
    charts: [],
  };
}

function fullColumnSqref(rows: number): { top: number; left: number; bottom: number; right: number } {
  return { top: 0, left: 0, bottom: rows - 1, right: 0 };
}

const FILL_RED: Dxf = {
  font: null,
  fill: { patternType: 'solid', fgColor: '#FF0000', bgColor: '#FF0000' },
  border: null,
};

const FONT_BLUE: Dxf = {
  font: {
    bold: false, italic: false, underline: false, strike: false,
    size: 11, color: '#0000FF', name: null,
  },
  fill: null,
  border: null,
};

// ─── cellIs (numeric) ────────────────────────────────────────────
describe('cellIs — numeric operators (ECMA-376 §18.3.1.10)', () => {
  function evalCellIs(value: number, operator: string, args: string[]): boolean {
    const rule: CfRule = { type: 'cellIs', operator, formulas: args, dxfId: 0, priority: 1 };
    const ws = sheetFromColumn([value], [{ sqref: [fullColumnSqref(1)], rules: [rule] }]);
    const ctx = compileCf(ws);
    const res = evaluateCf(numCell(0, 0, value), 0, 0, ctx, [FILL_RED]);
    return res.fill?.fgColor === '#FF0000';
  }

  it('greaterThan matches strictly greater', () => {
    expect(evalCellIs(5, 'greaterThan', ['3'])).toBe(true);
    expect(evalCellIs(3, 'greaterThan', ['3'])).toBe(false);
  });

  it('greaterThanOrEqual includes the boundary', () => {
    expect(evalCellIs(3, 'greaterThanOrEqual', ['3'])).toBe(true);
    expect(evalCellIs(2, 'greaterThanOrEqual', ['3'])).toBe(false);
  });

  it('lessThan / lessThanOrEqual', () => {
    expect(evalCellIs(2, 'lessThan', ['3'])).toBe(true);
    expect(evalCellIs(3, 'lessThan', ['3'])).toBe(false);
    expect(evalCellIs(3, 'lessThanOrEqual', ['3'])).toBe(true);
  });

  it('equal / notEqual', () => {
    expect(evalCellIs(3, 'equal', ['3'])).toBe(true);
    expect(evalCellIs(3, 'notEqual', ['3'])).toBe(false);
    expect(evalCellIs(4, 'notEqual', ['3'])).toBe(true);
  });

  it('between is inclusive of both endpoints', () => {
    expect(evalCellIs(3, 'between', ['3', '5'])).toBe(true);
    expect(evalCellIs(5, 'between', ['3', '5'])).toBe(true);
    expect(evalCellIs(2, 'between', ['3', '5'])).toBe(false);
    expect(evalCellIs(6, 'between', ['3', '5'])).toBe(false);
  });

  it('notBetween excludes the closed interval', () => {
    expect(evalCellIs(2, 'notBetween', ['3', '5'])).toBe(true);
    expect(evalCellIs(4, 'notBetween', ['3', '5'])).toBe(false);
  });
});

// ─── cellIs (text) ───────────────────────────────────────────────
describe('cellIs — text operators (case-insensitive)', () => {
  function evalText(value: string, operator: string, args: string[]): boolean {
    const rule: CfRule = {
      type: 'cellIs',
      operator,
      // quoted string literals as Excel writes them in <formula>
      formulas: args.map(a => `"${a}"`),
      dxfId: 0,
      priority: 1,
    };
    const ws: Worksheet = {
      ...sheetFromColumn([0], [{ sqref: [fullColumnSqref(1)], rules: [rule] }]),
      rows: [{ index: 0, height: null, cells: [textCell(0, 0, value)] }],
    };
    const ctx = compileCf(ws);
    const res = evaluateCf(textCell(0, 0, value), 0, 0, ctx, [FILL_RED]);
    return res.fill?.fgColor === '#FF0000';
  }

  it('equal is case-insensitive', () => {
    expect(evalText('Apple', 'equal', ['apple'])).toBe(true);
    expect(evalText('Apple', 'equal', ['pear'])).toBe(false);
  });

  it('containsText / notContains', () => {
    expect(evalText('Pineapple', 'containsText', ['apple'])).toBe(true);
    expect(evalText('Pear', 'notContains', ['apple'])).toBe(true);
  });

  it('beginsWith / endsWith', () => {
    expect(evalText('Pineapple', 'beginsWith', ['pine'])).toBe(true);
    expect(evalText('Pineapple', 'endsWith', ['apple'])).toBe(true);
    expect(evalText('Pineapple', 'endsWith', ['pine'])).toBe(false);
  });
});

// ─── colorScale ──────────────────────────────────────────────────
describe('colorScale — gradient interpolation (ECMA-376 §18.3.1.16)', () => {
  function colorAt(value: number, values: number[], rule: CfRule): string | null | undefined {
    const ws = sheetFromColumn(values, [{ sqref: [fullColumnSqref(values.length)], rules: [rule] }]);
    const ctx = compileCf(ws);
    const res = evaluateCf(numCell(0, 0, value), 0, 0, ctx, []);
    return res.fill?.fgColor;
  }

  const twoColor: CfRule = {
    type: 'colorScale',
    priority: 1,
    stops: [
      { kind: 'min', value: null, color: '#000000' },
      { kind: 'max', value: null, color: '#FFFFFF' },
    ],
  };

  it('2-color: min stop → first color, max stop → last color', () => {
    expect(colorAt(0, [0, 10], twoColor)).toBe('#000000');
    expect(colorAt(10, [0, 10], twoColor)).toBe('#FFFFFF');
  });

  it('2-color: midpoint interpolates halfway (#000000→#FFFFFF = #808080)', () => {
    expect(colorAt(5, [0, 10], twoColor)).toBe('#808080');
  });

  it('2-color: value below min / above max clamps to the endpoints', () => {
    expect(colorAt(-5, [0, 10], twoColor)).toBe('#000000');
    expect(colorAt(99, [0, 10], twoColor)).toBe('#FFFFFF');
  });

  it('min==max (all equal) does not divide by zero — returns a stop color', () => {
    // All samples equal → scaleMin == scaleMax. interpolation t is forced to 0.
    const color = colorAt(7, [7, 7, 7], twoColor);
    expect(color).toBe('#000000');
  });

  const threeColor: CfRule = {
    type: 'colorScale',
    priority: 1,
    stops: [
      { kind: 'min', value: null, color: '#FF0000' },
      { kind: 'percentile', value: '50', color: '#FFFF00' },
      { kind: 'max', value: null, color: '#00FF00' },
    ],
  };

  it('3-color: midpoint stop hits the middle color exactly', () => {
    // values 0..10, 50th percentile of [0,5,10] = 5 → middle color #FFFF00
    expect(colorAt(5, [0, 5, 10], threeColor)).toBe('#FFFF00');
  });

  it('3-color: lower band interpolates between first and middle stops', () => {
    // value 2.5 sits halfway between min(0) and mid(5): #FF0000→#FFFF00 = #FF8000
    expect(colorAt(2.5, [0, 5, 10], threeColor)).toBe('#FF8000');
  });
});

// ─── dataBar ─────────────────────────────────────────────────────
describe('dataBar — bar ratio (ECMA-376 §18.3.1.28)', () => {
  function ratioAt(value: number, values: number[]): number | undefined {
    const rule: CfRule = {
      type: 'dataBar',
      color: '#638EC6',
      min: { kind: 'min', value: null },
      max: { kind: 'max', value: null },
      priority: 1,
      gradient: true,
    };
    const ws = sheetFromColumn(values, [{ sqref: [fullColumnSqref(values.length)], rules: [rule] }]);
    const ctx = compileCf(ws);
    const res = evaluateCf(numCell(0, 0, value), 0, 0, ctx, []);
    return res.dataBar?.ratio;
  }

  it('min → 0, max → 1, midpoint → 0.5', () => {
    expect(ratioAt(0, [0, 10])).toBe(0);
    expect(ratioAt(10, [0, 10])).toBe(1);
    expect(ratioAt(5, [0, 10])).toBe(0.5);
  });

  it('clamps below-min and above-max into [0,1]', () => {
    expect(ratioAt(-5, [-5, 0, 10])).toBe(0);
    expect(ratioAt(10, [0, 5, 10])).toBe(1);
  });

  it('min==max yields ratio 0 (no division by zero)', () => {
    expect(ratioAt(5, [5, 5, 5])).toBe(0);
  });
});

// ─── top10 ───────────────────────────────────────────────────────
describe('top10 — rank & percent thresholds (ECMA-376 §18.3.1.10)', () => {
  function matchesTop10(
    value: number,
    values: number[],
    opts: { top: boolean; percent: boolean; rank: number },
  ): boolean {
    const rule: CfRule = { type: 'top10', ...opts, dxfId: 0, priority: 1 };
    const ws = sheetFromColumn(values, [{ sqref: [fullColumnSqref(values.length)], rules: [rule] }]);
    const ctx = compileCf(ws);
    const res = evaluateCf(numCell(0, 0, value), 0, 0, ctx, [FILL_RED]);
    return res.fill?.fgColor === '#FF0000';
  }

  const data = [1, 2, 3, 4, 5];

  it('top 2 by rank highlights the two largest', () => {
    expect(matchesTop10(5, data, { top: true, percent: false, rank: 2 })).toBe(true);
    expect(matchesTop10(4, data, { top: true, percent: false, rank: 2 })).toBe(true);
    expect(matchesTop10(3, data, { top: true, percent: false, rank: 2 })).toBe(false);
  });

  it('bottom 2 by rank highlights the two smallest', () => {
    expect(matchesTop10(1, data, { top: false, percent: false, rank: 2 })).toBe(true);
    expect(matchesTop10(2, data, { top: false, percent: false, rank: 2 })).toBe(true);
    expect(matchesTop10(3, data, { top: false, percent: false, rank: 2 })).toBe(false);
  });

  it('top 40% selects roughly the top two of five', () => {
    // percent threshold uses the (1 - rank/100) percentile of the samples.
    expect(matchesTop10(5, data, { top: true, percent: true, rank: 40 })).toBe(true);
    expect(matchesTop10(1, data, { top: true, percent: true, rank: 40 })).toBe(false);
  });
});

// ─── aboveAverage ────────────────────────────────────────────────
describe('aboveAverage — mean / stdDev bands (ECMA-376 §18.3.1.10)', () => {
  function matchesAvg(
    value: number,
    values: number[],
    rule: CfRule,
  ): boolean {
    const ws = sheetFromColumn(values, [{ sqref: [fullColumnSqref(values.length)], rules: [rule] }]);
    const ctx = compileCf(ws);
    const res = evaluateCf(numCell(0, 0, value), 0, 0, ctx, [FILL_RED]);
    return res.fill?.fgColor === '#FF0000';
  }

  const data = [1, 2, 3, 4, 5]; // mean = 3

  it('aboveAverage=true highlights values strictly above the mean', () => {
    const rule: CfRule = { type: 'aboveAverage', aboveAverage: true, dxfId: 0, priority: 1 };
    expect(matchesAvg(4, data, rule)).toBe(true);
    expect(matchesAvg(3, data, rule)).toBe(false);
    expect(matchesAvg(2, data, rule)).toBe(false);
  });

  it('aboveAverage=false highlights values strictly below the mean', () => {
    const rule: CfRule = { type: 'aboveAverage', aboveAverage: false, dxfId: 0, priority: 1 };
    expect(matchesAvg(2, data, rule)).toBe(true);
    expect(matchesAvg(3, data, rule)).toBe(false);
  });

  it('equalAverage=true includes values exactly at the mean', () => {
    const rule: CfRule = {
      type: 'aboveAverage', aboveAverage: true, equalAverage: true, dxfId: 0, priority: 1,
    };
    expect(matchesAvg(3, data, rule)).toBe(true);
    expect(matchesAvg(4, data, rule)).toBe(true);
    expect(matchesAvg(2, data, rule)).toBe(false);
  });

  it('stdDev=1 above highlights values beyond mean + 1·σ', () => {
    // population σ of [1..5] = sqrt(2) ≈ 1.414 → threshold ≈ 4.414
    const rule: CfRule = {
      type: 'aboveAverage', aboveAverage: true, stdDev: 1, dxfId: 0, priority: 1,
    };
    expect(matchesAvg(5, data, rule)).toBe(true);
    expect(matchesAvg(4, data, rule)).toBe(false);
  });

  it('stdDev=1 below highlights values beyond mean - 1·σ', () => {
    const rule: CfRule = {
      type: 'aboveAverage', aboveAverage: false, stdDev: 1, dxfId: 0, priority: 1,
    };
    expect(matchesAvg(1, data, rule)).toBe(true);
    expect(matchesAvg(2, data, rule)).toBe(false);
  });
});

// ─── iconSet ─────────────────────────────────────────────────────
describe('iconSet — threshold→icon index (ECMA-376 §18.3.1.10)', () => {
  function iconAt(value: number, values: number[], reverse = false): number | undefined {
    const rule: CfRule = {
      type: 'iconSet',
      iconSet: '3TrafficLights1',
      // default 3-icon cfvos: 0% / 33% / 67%
      cfvos: [
        { kind: 'percent', value: '0' },
        { kind: 'percent', value: '33' },
        { kind: 'percent', value: '67' },
      ],
      reverse,
      priority: 1,
    };
    const ws = sheetFromColumn(values, [{ sqref: [fullColumnSqref(values.length)], rules: [rule] }]);
    const ctx = compileCf(ws);
    const res = evaluateCf(numCell(0, 0, value), 0, 0, ctx, []);
    return res.iconSet?.index;
  }

  it('assigns the highest icon for top values, lowest for bottom', () => {
    // values 0..9: 0% → idx0, 33% (≈3.3) → idx1, 67% (≈6.7) → idx2
    expect(iconAt(0, [0, 9])).toBe(0);
    expect(iconAt(9, [0, 9])).toBe(2);
  });

  it('reverse flips the icon index', () => {
    expect(iconAt(9, [0, 9], true)).toBe(0);
    expect(iconAt(0, [0, 9], true)).toBe(2);
  });
});

// ─── rule priority / first-match-wins ────────────────────────────
describe('rule priority — lower number wins per property', () => {
  it('a higher-priority (lower number) fill rule wins over a later one', () => {
    const high: CfRule = { type: 'cellIs', operator: 'greaterThan', formulas: ['0'], dxfId: 0, priority: 1 };
    const low: CfRule = { type: 'cellIs', operator: 'greaterThan', formulas: ['0'], dxfId: 1, priority: 2 };
    const ws = sheetFromColumn([5], [{ sqref: [fullColumnSqref(1)], rules: [low, high] }]);
    const ctx = compileCf(ws);
    const res = evaluateCf(numCell(0, 0, 5), 0, 0, ctx, [FILL_RED, FONT_BLUE]);
    // priority 1 (FILL_RED via dxfId 0) wins the fill slot
    expect(res.fill?.fgColor).toBe('#FF0000');
  });
});

describe('font toggles — explicit off is a set property (ECMA-376 §18.8.2)', () => {
  it('a higher-priority explicit off claims every toggle over a lower-priority on', () => {
    const font = { bold: false, italic: false, underline: false, strike: false, size: 11, color: null, name: null };
    const off: Dxf = {
      font, fill: null, border: null,
      fontToggles: { bold: false, italic: false, underline: false, strike: false },
    };
    const on: Dxf = {
      font: { ...font, bold: true, italic: true, underline: true, strike: true }, fill: null, border: null,
      fontToggles: { bold: true, italic: true, underline: true, strike: true },
    };
    const high: CfRule = { type: 'cellIs', operator: 'greaterThan', formulas: ['0'], dxfId: 0, priority: 1 };
    const low: CfRule = { type: 'cellIs', operator: 'greaterThan', formulas: ['0'], dxfId: 1, priority: 2 };
    const ws = sheetFromColumn([5], [{ sqref: [fullColumnSqref(1)], rules: [low, high] }]);
    const res = evaluateCf(numCell(0, 0, 5), 0, 0, compileCf(ws), [off, on]);
    expect(res).toMatchObject({ fontBold: false, fontItalic: false, fontUnderline: false, fontStrike: false });
  });
});

// ─── stopIfTrue ──────────────────────────────────────────────────
describe('stopIfTrue — a matching rule stops lower-priority rules (ECMA-376 §18.3.1.10)', () => {
  // Column [1, 5]; the cell under test holds 5, which each stopping rule
  // below matches (5 > 0, TRUE, top 1, above the mean 3).
  const stoppers: Array<[string, CfRule]> = [
    ['cellIs', { type: 'cellIs', operator: 'greaterThan', formulas: ['0'], dxfId: 1, priority: 1, stopIfTrue: true }],
    ['expression', { type: 'expression', formula: 'TRUE', dxfId: 1, priority: 1, stopIfTrue: true }],
    ['top10', { type: 'top10', top: true, percent: false, rank: 1, dxfId: 1, priority: 1, stopIfTrue: true }],
    ['aboveAverage', { type: 'aboveAverage', aboveAverage: true, dxfId: 1, priority: 1, stopIfTrue: true }],
  ];
  // Lower-priority rules that would otherwise all apply to 5: a dxf fill,
  // an icon set and a data bar.
  const lower: CfRule[] = [
    { type: 'expression', formula: 'TRUE', dxfId: 0, priority: 2, stopIfTrue: false },
    { type: 'iconSet', iconSet: '3Arrows', cfvos: [{ kind: 'percent', value: '0' }, { kind: 'percent', value: '33' }, { kind: 'percent', value: '67' }], reverse: false, priority: 3 },
    { type: 'dataBar', color: '#638EC6', min: { kind: 'min', value: null }, max: { kind: 'max', value: null }, priority: 4, gradient: true },
  ];
  const evalWith = (stopper: CfRule) => {
    const ws = sheetFromColumn([1, 5], [{ sqref: [fullColumnSqref(2)], rules: [...lower, stopper] }]);
    return evaluateCf(numCell(1, 0, 5), 1, 0, compileCf(ws), [FILL_RED, FONT_BLUE]);
  };

  it.each(stoppers)('%s: applies its own dxf and skips every lower-priority rule', (_, stopper) => {
    const res = evalWith(stopper);
    expect(res.fontColor).toBe('#0000FF');
    expect(res.fill).toBeUndefined();
    expect(res.iconSet).toBeUndefined();
    expect(res.dataBar).toBeUndefined();
  });

  it('does not stop when the rule does not match, or when the flag is off', () => {
    const miss = evalWith({ type: 'cellIs', operator: 'greaterThan', formulas: ['10'], dxfId: 1, priority: 1, stopIfTrue: true });
    expect(miss.fontColor).toBeUndefined();
    expect(miss.fill?.fgColor).toBe('#FF0000');
    const noFlag = evalWith({ type: 'cellIs', operator: 'greaterThan', formulas: ['0'], dxfId: 1, priority: 1 });
    expect(noFlag.fontColor).toBe('#0000FF');
    expect(noFlag.fill?.fgColor).toBe('#FF0000');
    expect(noFlag.iconSet).toBeDefined();
    expect(noFlag.dataBar).toBeDefined();
  });

  // Excel for Mac honours a set flag on scale rules in SpreadsheetML (see
  // evaluateCf): a lower bold rule stays off every numeric cell they format.
  const scales: Array<[string, CfRule]> = [
    ['colorScale', { type: 'colorScale', stops: [{ kind: 'min', value: null, color: '#FFFF00' }, { kind: 'max', value: null, color: '#00B0F0' }], priority: 1 }],
    ['dataBar', { type: 'dataBar', color: '#638EC6', min: { kind: 'min', value: null }, max: { kind: 'max', value: null }, priority: 1, gradient: true }],
    ['iconSet', { type: 'iconSet', iconSet: '3Arrows', cfvos: [{ kind: 'percent', value: '0' }, { kind: 'percent', value: '33' }, { kind: 'percent', value: '67' }], reverse: false, priority: 1 }],
  ];
  it.each(scales)('%s: stops lower-priority rules only when stopIfTrue is set', (_, scale) => {
    const bold: Dxf = {
      font: { bold: true, italic: false, underline: false, strike: false, size: 11, color: null, name: null },
      fill: null, border: null, fontToggles: { bold: true },
    };
    const evalScale = (stopIfTrue: boolean) => {
      const rules: CfRule[] = [
        { ...scale, stopIfTrue } as CfRule,
        { type: 'expression', formula: 'TRUE', dxfId: 0, priority: 2, stopIfTrue: false },
      ];
      const ws = sheetFromColumn([1, 9], [{ sqref: [fullColumnSqref(2)], rules }]);
      return evaluateCf(numCell(1, 0, 9), 1, 0, compileCf(ws), [bold]);
    };
    expect(evalScale(true).fontBold).toBeUndefined();
    expect(evalScale(false).fontBold).toBe(true);
  });

  it('a stopping colour-only rule keeps a lower-priority explicit font-toggle off from applying', () => {
    const font = { bold: false, italic: false, underline: false, strike: false, size: 11, color: null, name: null };
    const off: Dxf = { font, fill: null, border: null, fontToggles: { bold: false, italic: false } };
    const colour: Dxf = { font: { ...font, color: '#FF0000' }, fill: null, border: null, fontToggles: {} };
    const ws = sheetFromColumn([5], [{
      sqref: [fullColumnSqref(1)],
      rules: [
        { type: 'cellIs', operator: 'greaterThan', formulas: ['0'], dxfId: 0, priority: 1, stopIfTrue: true },
        { type: 'expression', formula: 'TRUE', dxfId: 1, priority: 2, stopIfTrue: false },
      ],
    }]);
    const res = evaluateCf(numCell(0, 0, 5), 0, 0, compileCf(ws), [colour, off]);
    expect(res.fontColor).toBe('#FF0000');
    expect(res.fontBold).toBeUndefined();
    expect(res.fontItalic).toBeUndefined();
  });
});

// ─── A rule stops only where its match is established ────────────
describe('stopIfTrue needs an established match', () => {
  const lowerRed: CfRule = { type: 'expression', formula: 'TRUE', dxfId: 0, priority: 2, stopIfTrue: false };
  const RED: Dxf = { font: { ...FONT_BLUE.font!, color: '#FF0000' }, fill: null, border: null };
  // The evaluated cell holds `value`; `extra` cells sit in sheet row 1
  // (so `B1` refers to `extra` column 2).
  const evalAt = (rules: CfRule[], extra: Cell[] = [], value = 5) => {
    const ws = sheetFromColumn([value], [{ sqref: [fullColumnSqref(1)], rules }]);
    ws.rows.push({ index: 1, height: null, cells: extra });
    return evaluateCf(ws.rows[0].cells[0], 0, 0, compileCf(ws), [RED]);
  };
  const cellIs = (operator: string, formula: string): CfRule =>
    ({ type: 'cellIs', operator, formulas: [formula], dxfId: null, priority: 1, stopIfTrue: true });
  const icons = (activeFormula?: string): CfRule => ({
    type: 'iconSet', iconSet: '3Arrows', reverse: false, priority: 1, stopIfTrue: true, activeFormula,
    cfvos: [{ kind: 'num', value: '0' }, { kind: 'num', value: '3' }, { kind: 'num', value: '7' }],
  });

  // Review reproducer: `0+10` used to be read as the literal 0, so 5 > 0
  // matched and the stop removed the lower rule's red text.
  it('cellIs evaluates a formula operand instead of truncating it to a literal', () => {
    expect(evalAt([cellIs('greaterThan', '0+10'), lowerRed]).fontColor).toBe('#FF0000');
    expect(evalAt([cellIs('greaterThan', '0+1'), lowerRed]).fontColor).toBeUndefined();
    // A cell reference operand resolves (Excel: 5 > $B$1 = 3 matches).
    expect(evalAt([cellIs('greaterThan', '$B$1'), lowerRed], [numCell(1, 2, 3)]).fontColor).toBeUndefined();
    // Outside the exactly evaluated grammar: no match, no stop.
    expect(evalAt([cellIs('greaterThan', 'UNSUPPORTED(1)'), lowerRed]).fontColor).toBe('#FF0000');
    expect(evalAt([cellIs('greaterThan', '2^2'), lowerRed]).fontColor).toBe('#FF0000');
    // A blank operand is not converted to 0.
    expect(evalAt([cellIs('greaterThan', '$B$1'), lowerRed]).fontColor).toBe('#FF0000');
  });

  // Round-3 review reproducers: each operand used to be mis-evaluated to a
  // value that matched, and the stop removed the lower rule's red text.
  it('functions, blank selection and overflow are unevaluable, not guessed', () => {
    // OR over a referenced "TRUE" string (Excel ignores text in references)
    // and OR short-circuiting past a type error.
    const textTrue: Cell = { row: 1, col: 2, value: { type: 'text', text: 'TRUE' }, styleIndex: 0 };
    expect(evalAt([cellIs('greaterThan', 'IF(OR(0,B1),0,10)'), lowerRed], [textTrue]).fontColor).toBe('#FF0000');
    expect(evalAt([cellIs('greaterThan', 'IF(OR(1,"invalid"),0,10)'), lowerRed]).fontColor).toBe('#FF0000');
    // IF selecting a blank cell (was replaced by TRUE).
    expect(evalAt([cellIs('lessThan', 'IF(1,B1,0)+1'), lowerRed], [], 1.5).fontColor).toBe('#FF0000');
    expect(evalAt([icons('IF(1,B1,0)'), lowerRed]).fontColor).toBe('#FF0000');
    // An overflowing product is #NUM!, not Infinity.
    const huge = numCell(1, 2, 1e200);
    expect(evalAt([cellIs('lessThan', 'B1*B1'), lowerRed], [huge]).fontColor).toBe('#FF0000');
    expect(evalAt([icons('B1*B1'), lowerRed], [huge]).fontColor).toBe('#FF0000');
  });

  // Review reproducer: an x14 icon set with activePresent="1" and xm:f 0
  // is not displayed ([MS-XLSX] 2.6.27), so it must not stop the red text.
  it('a scale rule applies and stops only while its activity formula is nonzero', () => {
    const inactive = evalAt([icons('0'), lowerRed]);
    expect(inactive.fontColor).toBe('#FF0000');
    expect(inactive.iconSet).toBeUndefined();
    const active = evalAt([icons('1'), lowerRed]);
    expect(active.fontColor).toBeUndefined();
    expect(active.iconSet).toBeDefined();
    expect(evalAt([icons(), lowerRed]).fontColor).toBeUndefined();
    // Relative references anchor at the rule range's top-left: over B2:B3
    // (B2 = 5, B3 = 1), `B2-5` is 0 at B2 and -4 at B3.
    const relative = (formula: string, row: number) => {
      const ws = sheetFromColumn([], [{ sqref: [{ top: 2, left: 2, bottom: 3, right: 2 }], rules: [icons(formula), lowerRed] }]);
      ws.rows = [{ index: 2, height: null, cells: [numCell(2, 2, 5)] }, { index: 3, height: null, cells: [numCell(3, 2, 1)] }];
      return evaluateCf(ws.rows[row - 2].cells[0], row, 2, compileCf(ws), [RED]).fontColor;
    };
    expect(relative('B2-5', 2)).toBe('#FF0000');
    expect(relative('B2-5', 3)).toBeUndefined();
    // A logical result is not read as zero / nonzero.
    expect(evalAt([icons('B1'), lowerRed], [{ row: 1, col: 2, value: { type: 'bool', bool: true }, styleIndex: 0 }]).fontColor).toBe('#FF0000');
    // Unevaluable: inactive rather than guessed.
    const unknown = evalAt([icons('UNSUPPORTED(A1)'), lowerRed]);
    expect(unknown.fontColor).toBe('#FF0000');
    expect(unknown.iconSet).toBeUndefined();
    const scale: CfRule = { type: 'colorScale', stops: [{ kind: 'min', value: null, color: '#FFFF00' }, { kind: 'max', value: null, color: '#00B0F0' }], priority: 1, stopIfTrue: true, activeFormula: '0' };
    const offScale = evalAt([scale, lowerRed]);
    expect(offScale.fill).toBeUndefined();
    expect(offScale.fontColor).toBe('#FF0000');
  });
});

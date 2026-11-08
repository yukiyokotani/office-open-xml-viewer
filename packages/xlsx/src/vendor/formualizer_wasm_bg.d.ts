/**
 * Minimal surface of the vendored Formualizer wrapper
 * (`formualizer_wasm_bg.js`, from `formualizer@0.13.0`) used by the recalc
 * pass.
 */
export type FormualizerCellValue =
  | null
  | number
  | boolean
  | string
  | Array<Array<FormualizerCellValue>>
  | { kind: 'error'; code: string };

export interface FormualizerCellAddress {
  sheet: string;
  row: number;
  column: number;
}

export interface FormualizerCellSnapshot {
  stamp: unknown;
  cell: {
    address: FormualizerCellAddress;
    formula?: string | null;
    value: FormualizerCellValue;
  };
}

/** Opaque engine workbook. Methods throw engine errors; callers degrade. */
export interface FormualizerWorkbook {
  setTemporalEgress(policy: string): void;
  setFormula(sheet: string, row: number, col: number, formula: string): void;
  setValue(sheet: string, row: number, col: number, value: number | boolean | string): void;
  evaluateAll(): unknown;
  evaluateCell(sheet: string, row: number, col: number): FormualizerCellValue;
  inspectCell(cell: FormualizerCellAddress): FormualizerCellSnapshot;
  free(): void;
}

export interface FormualizerEngine {
  Workbook: {
    fromJson(json: string): FormualizerWorkbook;
  };
}

declare const engine: FormualizerEngine;
/** Shim namespace (`__wbg_*`) consumed by the hand-rolled instantiation. */
declare function __wbg_set_wasm(exports: WebAssembly.Exports): void;

export default engine;
export { __wbg_set_wasm };

import type { Cell, CellValue, SharedString, Worksheet } from './types.js';

/**
 * Opt-in Excel formula recalculation (Formualizer-backed).
 *
 * This module is a deliberate fork of the upstream no-recalc policy
 * (packages/xlsx/CLAUDE.md §"数式の扱い": cell `<f>` formulas are NEVER
 * calculated, the renderer always shows the cached `<v>`). Everything here
 * runs ONLY when the caller opts in via `LoadOptions.recalculateFormulas`.
 * The default path never imports Formualizer (lazy `await import`, so the
 * default bundle pays nothing) and never touches `cell.value`.
 *
 * Injection point: every rendered cell goes through
 * `formatCellValueWithColor(cell, ...)` (number-format.ts), which reads ONLY
 * `cell.value`. Writing computed results back into `cell.value` (and never
 * into `cell.formula`) makes number formats, date systems, `;;;` hiding,
 * `####` overflow, CF dxf numFmt and showZeros work unchanged. Other
 * `cell.value` consumers (`XlsxWorkbook.cellText`, `resolveValidationList`,
 * markdown export) pick up recalculated values for free. In `worker` mode
 * the worker admits its own copy before the main-thread recalc runs, so the
 * caller forwards overwritten cells (`capture` → `applyRecalculatedValues`)
 * and the worker applies them at admission; the engine never runs in-worker.
 *
 * Contracts:
 * - Never throws: ANY Formualizer failure (parse failure, unsupported
 *   function, cycle, OOM, WASM init failure) degrades to the cell's cached
 *   value. Results are computed into a pending list first and applied only
 *   after a successful read-back, so a catastrophic failure leaves cells
 *   untouched.
 * - `cell.formula` (raw `<f>` text, ECMA-376 §18.3.1.40) is preserved.
 * - Resource governance is library policy, not an OOXML rule: feeding is
 *   capped by `maxCells` (default below). On budget exceed the sheet is
 *   skipped WHOLE (no half-recalculated sheets) and cached values are kept.
 *
 * Formualizer is MIT/Apache-2.0 licensed (permissive). It is a lazy optional
 * dependency: the default (opt-out) path never loads its WASM bytes.
 */

export interface RecalcContext {
  /** Workbook `sharedStrings` table for resolving `shared{si}` cells before
   *  feeding (mirrors `resolveSharedStrings`: out-of-range ⇒ empty text). */
  sharedStrings: SharedString[];
  /** Workbook date system (`<workbookPr date1904>`, ECMA-376 §18.2.28). */
  date1904: boolean;
  /** Sheet-name → worksheet resolver for cross-sheet refs (`Sheet2!A1`).
   *  Lookup is case-insensitive, matching `resolveValidationList`
   *  (workbook.ts). May be sync (already-materialized cache) or async (lazy
   *  lease load). When omitted — or when a name does not resolve — the
   *  dependent formula keeps its cached value. */
  resolveSheet?: (name: string) => Worksheet | undefined | Promise<Worksheet | undefined>;
  /** Cap on total cells fed to Formualizer (literals + formulas). Whole
   *  recalc is skipped when exceeded. Default {@link DEFAULT_RECALC_MAX_CELLS}. */
  maxCells?: number;
}

export interface RecalcStats {
  /** Formula cells overwritten with a fresh Formualizer result. */
  recalculated: number;
  /** Formula cells that kept their cached value (unsupported, cycle,
   *  non-Excel error, exception). Rendering is unaffected. */
  failed: number;
  /** Formula cells skipped without evaluation (resource cap). */
  skipped: number;
}

/**
 * Default feed cap (library resource policy, not an OOXML rule).
 * Formualizer evaluates in-process on the calling thread; an unbounded feed
 * would let a hostile sheet stall the UI or OOM the worker. Worksheets admit
 * up to 250k cells (HARD_MAX_XLSX_WORKSHEET_CELLS) but typical formula sheets
 * are far smaller — 100k fed cells keeps a recalc pause bounded while
 * covering realistic workbooks. Configurable via {@link RecalcContext.maxCells}.
 */
export const DEFAULT_RECALC_MAX_CELLS = 100_000;

/** Excel error codes round-tripped into `CellValue.error` (ECMA-376 §18.17,
 *  plus the dynamic-array `#SPILL!` / `#CALC!` errors Excel added later).
 *  Formualizer-specific errors (`#CIRC!`, `#N/IMPL!`, `#CANCELLED!`) are NOT
 *  Excel values: those cells keep their cached `<v>` and count as `failed`. */
const EXCEL_ERROR_CODES = new Set([
  '#DIV/0!',
  '#N/A',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#NULL!',
  '#SPILL!',
  '#CALC!',
]);

const MAX_WORKSHEET_ROW = 1_048_576;
const MAX_WORKSHEET_COL = 16_384;

/** A materialized sheet fed to one shared Formualizer instance. */
export interface RecalcSheetInput {
  name: string;
  ws: Worksheet;
}

/** One recalculated cell, shipped to the render worker in `worker` mode so
 *  its sheet cache matches the main-thread model. `CellValue` is plain
 *  JSON-safe data (the same payload worksheets already cross as). */
export interface RecalcPatchedCell {
  row: number;
  col: number;
  value: CellValue;
}

/**
 * Recalculate ONE worksheet, pulling cross-sheet inputs through
 * `ctx.resolveSheet`. Referenced sheets are fed read-only (their cells are
 * NOT mutated); only `ws` is written. Prefer {@link recalculateWorksheets}
 * when all sheets are already materialized (two-phase workbook recalc).
 * When `capture` is given, every overwritten cell is appended (row/col/value)
 * so the caller can forward values across the worker boundary.
 */
export async function recalculateWorksheet(
  ws: Worksheet,
  ctx: RecalcContext,
  capture?: RecalcPatchedCell[],
): Promise<RecalcStats> {
  const feed: RecalcSheetInput[] = [{ name: ws.name, ws }];
  if (ctx.resolveSheet) {
    const wanted = collectSheetReferences(ws);
    const seen = new Set(feed.map((s) => s.name.toLowerCase()));
    const resolved = await Promise.all(
      wanted.map(async (ref) => {
        if (seen.has(ref.toLowerCase())) return undefined;
        seen.add(ref.toLowerCase());
        try {
          const target = await ctx.resolveSheet!(ref);
          return target ? { name: ref, ws: target } : undefined;
        } catch {
          return undefined;
        }
      }),
    );
    for (const r of resolved) if (r) feed.push(r);
  }
  const captured = new Map<string, RecalcPatchedCell[]>();
  const perSheet = await recalculateWorksheets(feed, {
    sharedStrings: ctx.sharedStrings,
    date1904: ctx.date1904,
    maxCells: ctx.maxCells,
    mutate: new Set([ws.name.toLowerCase()]),
    capture: captured,
  });
  if (capture) capture.push(...(captured.get(ws.name.toLowerCase()) ?? []));
  return perSheet.get(ws.name.toLowerCase()) ?? { recalculated: 0, failed: 0, skipped: 0 };
}

/**
 * Recalculate a set of materialized worksheets sharing ONE Formualizer
 * instance (two-phase workbook recalc: call after every sheet is loaded, so
 * cross-sheet refs resolve). By default every fed sheet is mutated; pass
 * `mutate` to restrict writes (used by {@link recalculateWorksheet} for its
 * read-only cross-sheet inputs).
 */
export async function recalculateWorksheets(
  sheets: RecalcSheetInput[],
  opts: {
    sharedStrings: SharedString[];
    date1904: boolean;
    maxCells?: number;
    mutate?: Set<string>;
    /** Per-sheet capture of overwritten cells (keyed by lowercase sheet
     *  name) for worker-boundary forwarding. Only mutated sheets capture. */
    capture?: Map<string, RecalcPatchedCell[]>;
  },
): Promise<Map<string, RecalcStats>> {
  const stats = new Map<string, RecalcStats>();
  for (const s of sheets) stats.set(s.name.toLowerCase(), { recalculated: 0, failed: 0, skipped: 0 });
  const statFor = (name: string): RecalcStats => {
    const key = name.toLowerCase();
    let st = stats.get(key);
    if (!st) {
      st = { recalculated: 0, failed: 0, skipped: 0 };
      stats.set(key, st);
    }
    return st;
  };

  if (sheets.length === 0) return stats;

  // Shared-formula anchors are sheet-local (ECMA-376 §18.3.1.40): resolve
  // each sheet's `si` groups once so followers (no `<f>` text) evaluate the
  // anchor with shifted relative refs instead of their cached `<v>`.
  const anchorsBySheet = new Map<string, SharedAnchors>();
  const targetsBySheet = new Map<string, FormulaTarget[]>();
  for (const { name, ws } of sheets) {
    const key = name.toLowerCase();
    const anchors = buildSharedAnchors(ws);
    anchorsBySheet.set(key, anchors);
    targetsBySheet.set(key, collectFormulaTargets(ws, anchors));
  }

  const maxCells = opts.maxCells ?? DEFAULT_RECALC_MAX_CELLS;
  let feedCount = 0;
  let formulaCount = 0;
  {
    const targeted = new Set<Cell>();
    for (const targets of targetsBySheet.values()) {
      for (const t of targets) {
        targeted.add(t.cell);
        formulaCount++;
      }
    }
    for (const { ws } of sheets) {
      for (const row of ws.rows) {
        for (const cell of row.cells) {
          // Every fed cell counts once: literals by their cached value,
          // formula targets (including empty-cached followers) by evaluation.
          if (cell.value.type !== 'empty' || targeted.has(cell)) feedCount++;
        }
      }
    }
  }
  if (feedCount > maxCells) {
    for (const { name } of sheets) {
      statFor(name).skipped += targetsBySheet.get(name.toLowerCase())?.length ?? 0;
    }
    return stats;
  }
  if (formulaCount === 0) return stats;

  // Lazy init: the default (opt-out) path never loads Formualizer WASM.
  // The module (and its WASM bytes) stay out of the default bundle chunk.
  let formModule: typeof import('formualizer');
  try {
    formModule = await loadFormualizerModule();
  } catch {
    for (const { name } of sheets) {
      statFor(name).failed += targetsBySheet.get(name.toLowerCase())?.length ?? 0;
    }
    return stats;
  }
  const { Workbook } = formModule;

  // Admitted engine names: Formualizer rejects empty sheet names, so map
  // `''` to the stable placeholder below. Collisions (two sheets sanitizing
  // to one engine name) keep cached values for the colliding sheets.
  const admittedByKey = new Map<string, string>();
  const admittedTaken = new Set<string>();
  for (const { name } of sheets) {
    const key = name.toLowerCase();
    const admitted = sanitizeSheetName(name);
    const takenKey = admitted.toLowerCase();
    if (admittedTaken.has(takenKey)) {
      // Collision: leave unadmitted (read-back skips it → cached, failed).
      continue;
    }
    admittedTaken.add(takenKey);
    admittedByKey.set(key, admitted);
  }

  let wb: ReturnType<typeof Workbook.fromJson> | undefined;
  try {
    try {
      // Date system is workbook-level (ECMA-376 §18.2.28): the empty-Workbook
      // constructor defaults to 1900, so build via the JSON adapter which
      // propagates `date_system_1904` to the engine (verified: DATE(2023,1,1)
      // reads 44927 in 1900, 43465 in 1904).
      const manifest: { sheets: Record<string, object> } = { sheets: {} };
      for (const admitted of admittedByKey.values()) {
        manifest.sheets[admitted] = opts.date1904 ? { date_system_1904: true } : {};
      }
      wb = Workbook.fromJson(JSON.stringify(manifest));
    } catch {
      for (const { name } of sheets) {
        statFor(name).failed += targetsBySheet.get(name.toLowerCase())?.length ?? 0;
      }
      return stats;
    }

    try {
      wb.setTemporalEgress('serial');
    } catch {
      // Serial egress keeps date results as Excel serials (the format layer
      // renders them). If this ever fails, date formulas fall back to cached
      // below instead of writing Date objects into `cell.value`.
    }

    // Feed phase: literals as context, formulas as `=...`. Shared-string
    // cells resolve BEFORE feeding (number-format.ts treats `shared` as
    // unreachable at render time; recalc must see the same text). Error
    // literals feed as `=<code>` so the engine holds a real error (which
    // propagates, e.g. `=A1+1` over `#N/A` stays `#N/A`).
    // `formulaByCell` carries the per-cell evaluation text (followers hold
    // the expanded anchor); `dataTable` targets (null) are never fed.
    for (const { name, ws } of sheets) {
      const key = name.toLowerCase();
      const admitted = admittedByKey.get(key);
      if (admitted === undefined) {
        statFor(name).failed += targetsBySheet.get(key)?.length ?? 0;
        continue;
      }
      const formulaByCell = new Map<Cell, string | null>();
      for (const t of targetsBySheet.get(key) ?? []) formulaByCell.set(t.cell, t.formula);
      for (const row of ws.rows) {
        for (const cell of row.cells) {
          const formula = formulaByCell.get(cell);
          if (formula === undefined && cell.value.type === 'empty') continue;
          // `dataTable` (null): never fed — the cached value is kept and the
          // target counts `failed` at read-back.
          if (formula === null) continue;
          try {
            if (formula !== undefined) {
              wb.setFormula(admitted, cell.row, cell.col, '=' + formula);
            } else {
              const literal = feedLiteral(cell, opts.sharedStrings);
              if (literal.kind === 'skip') continue;
              if (literal.kind === 'error') {
                wb.setFormula(admitted, cell.row, cell.col, '=' + literal.code);
              } else {
                wb.setValue(admitted, cell.row, cell.col, literal.value);
              }
            }
          } catch {
            // One poisoned cell (e.g. unparseable literal) must not abort
            // the sheet: its formula dependents fall back to cached below.
          }
        }
      }
    }

    try {
      await wb.evaluateAll();
    } catch {
      // Catastrophic evaluation failure: keep every cached value. Count the
      // unwritten remainder below via the generic tail handler.
      for (const { name } of sheets) {
        const key = name.toLowerCase();
        if (opts.mutate && !opts.mutate.has(key)) continue;
        statFor(name).failed += targetsBySheet.get(key)?.length ?? 0;
      }
      return stats;
    }

    // Read-back phase (no mutation yet): collect pending writes so a
    // catastrophic read failure leaves cells untouched.
    const pending: Array<{ cell: Cell; value: CellValue; sheetName: string }> = [];
    const failures: Array<{ sheetName: string }> = [];
    const captureBucket = (sheetName: string): RecalcPatchedCell[] | undefined => {
      if (!opts.capture) return undefined;
      const key = sheetName.toLowerCase();
      let bucket = opts.capture.get(key);
      if (!bucket) {
        bucket = [];
        opts.capture.set(key, bucket);
      }
      return bucket;
    };
    for (const { name } of sheets) {
      const key = name.toLowerCase();
      if (opts.mutate && !opts.mutate.has(key)) continue;
      const admitted = admittedByKey.get(key);
      if (admitted === undefined) continue;
      for (const t of targetsBySheet.get(key) ?? []) {
        if (t.formula === null) {
          failures.push({ sheetName: name });
          continue;
        }
        let value: CellValue | null = null;
        try {
          const snapshot = wb.inspectCell({ sheet: admitted, row: t.cell.row, column: t.cell.col });
          value = formValueToCellValue(snapshot.cell.value);
        } catch {
          value = null;
        }
        if (value === null) failures.push({ sheetName: name });
        else pending.push({ cell: t.cell, value, sheetName: name });
      }
    }

    for (const { cell, value, sheetName } of pending) {
      cell.value = value;
      statFor(sheetName).recalculated++;
      // Pending holds only mutated sheets' cells (read-back skips the
      // `mutate` allow-list), so capture is inherently scoped correctly.
      captureBucket(sheetName)?.push({ row: cell.row, col: cell.col, value });
    }
    for (const { sheetName } of failures) statFor(sheetName).failed++;
  } catch {
    // Catastrophic engine failure after partial reads: cells already
    // rewritten keep their recomputed values (valid results); every other
    // formula cell keeps its cached value. Count the unwritten remainder.
    for (const { name } of sheets) {
      const key = name.toLowerCase();
      if (opts.mutate && !opts.mutate.has(key)) continue;
      const st = statFor(name);
      const accounted = st.recalculated + st.failed;
      const total = targetsBySheet.get(key)?.length ?? 0;
      st.failed += Math.max(0, total - accounted);
    }
  } finally {
    try {
      (wb as unknown as { free?: () => void })?.free?.();
    } catch {
      // Engine teardown must never break rendering.
    }
    wb = undefined;
  }
  return stats;
}

/** Memoized Formualizer WASM init: the default path never imports it, and
 *  concurrent recalcs share one init. Reset on failure so a later recalc can
 *  retry instead of pinning a rejected promise. */
let formInit: Promise<typeof import('formualizer')> | null = null;
function loadFormualizerModule(): Promise<typeof import('formualizer')> {
  if (!formInit) {
    formInit = (async () => {
      const mod = await import('formualizer');
      await mod.default();
      return mod;
    })();
    formInit.catch(() => {
      formInit = null;
    });
  }
  return formInit;
}

/** One shared-formula anchor: the cell whose `<f t="shared">` carries text
 *  (ECMA-376 §18.3.1.40). Followers reference it by `formulaSi`. */
interface SharedAnchor {
  formula: string;
  row: number;
  col: number;
}

/** Anchor cells keyed by `<f si>` group index (sheet-local). */
type SharedAnchors = Map<number, SharedAnchor>;

/** One cell to evaluate. `formula` is the text to feed (`=`-prefixed by the
 *  caller); `null` marks a `dataTable` cell — kept cached, counted `failed`. */
interface FormulaTarget {
  cell: Cell;
  formula: string | null;
}

/** Index a sheet's shared-formula anchors by group. First anchor wins on a
 *  malformed duplicate `si`; orphans keep cached values downstream. */
function buildSharedAnchors(ws: Worksheet): SharedAnchors {
  const anchors: SharedAnchors = new Map();
  for (const row of ws.rows) {
    for (const cell of row.cells) {
      if (
        cell.formula !== undefined &&
        cell.formulaType === 'shared' &&
        cell.formulaSi !== undefined &&
        !anchors.has(cell.formulaSi)
      ) {
        anchors.set(cell.formulaSi, { formula: cell.formula, row: cell.row, col: cell.col });
      }
    }
  }
  return anchors;
}

/**
 * Collect every cell to evaluate:
 * - normal/array cells: their own `<f>` text (`array` repeats per cell, so
 *   no shifting is needed);
 * - `shared` followers (no text): the anchor expanded with shifted refs;
 * - `dataTable`: `null` — Formualizer has no TABLE() support and would
 *   report `#NAME?`, clobbering a valid cached value, so the cache is kept.
 */
function collectFormulaTargets(ws: Worksheet, anchors: SharedAnchors): FormulaTarget[] {
  const out: FormulaTarget[] = [];
  for (const row of ws.rows) {
    for (const cell of row.cells) {
      if (cell.formula !== undefined) {
        out.push({ cell, formula: cell.formulaType === 'dataTable' ? null : cell.formula });
      } else if (cell.formulaType === 'shared' && cell.formulaSi !== undefined) {
        const anchor = anchors.get(cell.formulaSi);
        if (anchor) {
          out.push({
            cell,
            formula: expandSharedFormula(anchor.formula, anchor.row, anchor.col, cell.row, cell.col),
          });
        }
      }
    }
  }
  return out;
}

/** Map one literal cell to a Formualizer feed. Error cells are returned as
 *  `{ kind: 'error' }` so the caller feeds `=<code>` (a real engine error,
 *  which propagates, e.g. `=A1+1` over `#N/A` stays `#N/A`). Shared-string
 *  cells resolve BEFORE feeding (number-format.ts treats `shared` as
 *  unreachable at render time; recalc must see the same text).
 *  `{ kind: 'skip' }` means leave the cell blank (empty literals). */
function feedLiteral(
  cell: Cell,
  sharedStrings: SharedString[],
):
  | { kind: 'skip' }
  | { kind: 'error'; code: string }
  | { kind: 'value'; value: string | number | boolean } {
  const v = cell.value;
  switch (v.type) {
    case 'empty':
      return { kind: 'skip' };
    case 'number':
      return { kind: 'value', value: v.number };
    case 'bool':
      return { kind: 'value', value: v.bool };
    case 'error':
      return { kind: 'error', code: v.error };
    case 'text':
      return { kind: 'value', value: v.text };
    case 'shared': {
      const ss = sharedStrings[v.si];
      return { kind: 'value', value: ss ? ss.text : '' };
    }
  }
}

/**
 * Map a Formualizer `inspectCell` value to a `CellValue` (via `evaluateAll`
 * + `inspectCell`, so errors arrive tagged as `{ kind: 'error', code }` and
 * never masquerade as text). Date serials stay numbers — the format layer
 * (`ws.date1904` + numFmt) renders them. Returns null when the cached `<v>`
 * must be kept: EMPTY/`null`/pending results from a formula are NOT blanked
 * (a cached value is more informative than null), arrays have no `CellValue`
 * shape, and Formualizer-specific errors (`#CIRC!`, `#N/IMPL!`,
 * `#CANCELLED!`) have no Excel representation.
 */
function formValueToCellValue(value: unknown): CellValue | null {
  if (value === null || value === undefined) return null;
  if (typeof value === 'number') {
    return Number.isFinite(value) ? { type: 'number', number: value } : null;
  }
  // Excel cells hold at most 32,767 characters (observed Office limit). A
  // longer recomputed string would blow the worksheet JSON budget the
  // admission accounting just asserted, so the cached `<v>` is kept instead
  // — counted `failed`, never truncated. (Formualizer already caps REPT-style
  // growth with `#VALUE!`, which round-trips as an Excel error below.)
  if (typeof value === 'string') {
    return value.length <= 32_767 ? { type: 'text', text: value } : null;
  }
  if (typeof value === 'boolean') return { type: 'bool', bool: value };
  if (Array.isArray(value)) return null;
  if (typeof value === 'object') {
    const rec = value as { kind?: unknown; code?: unknown };
    if (rec.kind === 'error' && typeof rec.code === 'string') {
      return EXCEL_ERROR_CODES.has(rec.code) ? { type: 'error', error: rec.code } : null;
    }
    // `{ kind: 'pending' }` and any other engine-internal shape: keep cached.
    return null;
  }
  return null;
}

/** Formualizer rejects empty sheet names; fall back to a stable placeholder
 *  (collisions surface as unadmitted sheets and keep cached values). */
function sanitizeSheetName(name: string): string {
  return name === '' ? 'Sheet' : name;
}

/**
 * Whether any formula on `ws` references sheet `name` (`Sheet!A1`,
 * case-insensitive). Used by the workbook to refresh cached dependents when
 * a newly admitted sheet changes their inputs. String literals never count.
 */
export function worksheetReferencesSheet(ws: Worksheet, name: string): boolean {
  const want = name.toLowerCase();
  for (const row of ws.rows) {
    for (const cell of row.cells) {
      if (cell.formula === undefined) continue;
      for (const ref of scanSheetReferences(cell.formula)) {
        if (ref.toLowerCase() === want) return true;
      }
    }
  }
  return false;
}

/**
 * Collect sheet names referenced as `Sheet!A1` qualifiers in a sheet's
 * formulas, so single-sheet recalc can feed cross-sheet inputs. String
 * literals (`"Sheet!A1"`) are skipped; quoted names (`'My Sheet'!A1`) are
 * unescaped (`''` → `'`). Unqualified refs need no input. Malformed
 * qualifiers fall back to cached values downstream — never an exception.
 */
function collectSheetReferences(ws: Worksheet): string[] {
  const names = new Set<string>();
  for (const row of ws.rows) {
    for (const cell of row.cells) {
      const f = cell.formula;
      if (f === undefined) continue;
      for (const name of scanSheetReferences(f)) names.add(name);
    }
  }
  return [...names];
}

function scanSheetReferences(formula: string): string[] {
  const out: string[] = [];
  let i = 0;
  const n = formula.length;
  while (i < n) {
    const c = formula[i];
    if (c === '"') {
      i++;
      while (i < n) {
        if (formula[i] === '"') {
          if (formula[i + 1] === '"') i += 2;
          else {
            i++;
            break;
          }
        } else i++;
      }
      continue;
    }
    if (c === "'") {
      let j = i + 1;
      let name = '';
      let closed = false;
      while (j < n) {
        if (formula[j] === "'") {
          if (formula[j + 1] === "'") {
            name += "'";
            j += 2;
          } else {
            closed = true;
            j++;
            break;
          }
        } else {
          name += formula[j];
          j++;
        }
      }
      if (closed && formula[j] === '!') {
        out.push(name);
        i = j + 1;
        continue;
      }
      i = j;
      continue;
    }
    if (isNameStart(c)) {
      let j = i + 1;
      while (j < n && isNameChar(formula[j])) j++;
      if (formula[j] === '!' && j > i) {
        const candidate = formula.slice(i, j);
        if (!isCellRef(candidate)) out.push(candidate);
        i = j + 1;
        continue;
      }
      i = j;
      continue;
    }
    i++;
  }
  return out;
}

function isNameStart(c: string): boolean {
  return (
    (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c === '_' || c === '\\'
  );
}

function isNameChar(c: string): boolean {
  return isNameStart(c) || (c >= '0' && c <= '9') || c === '.';
}

/** True for A1-style cell spellings (which cannot be sheet names in a
 *  qualifier position — `A1!B2` is invalid; the qualifier is a name). */
function isCellRef(s: string): boolean {
  let i = 0;
  if (s[i] === '$') i++;
  const colStart = i;
  while (i < s.length && isColLetter(s[i])) i++;
  if (i === colStart || i - colStart > 3) return false;
  if (s[i] === '$') i++;
  const rowStart = i;
  while (i < s.length && s[i] >= '0' && s[i] <= '9') i++;
  if (i === rowStart || i !== s.length) return false;
  const rowNum = parseInt(s.slice(rowStart, i), 10);
  return rowNum >= 1 && rowNum <= MAX_WORKSHEET_ROW;
}

function isColLetter(c: string): boolean {
  return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
}

/**
 * Expand a shared/array anchor formula to a follower cell (ECMA-376
 * §18.3.1.40 `t="shared"|"array"`, `ref`, `si`).
 *
 * Relative A1 refs shift by (follower − anchor); `$` pins a coordinate.
 * Sheet qualifiers (`Sheet!A1`, `'My Sheet'!A1`) do NOT pin — filling across
 * rows still shifts the cell. String literals are never rewritten. Range
 * ends shift independently. Tokens that are not in-grid cell refs
 * (names, functions, structured refs, R1C1) pass through untouched.
 *
 * NOTE: the parser does not yet retain `<f t/ref/si>` attributes, so every
 * formula today behaves as `t="normal"` and this is exercised only once the
 * parser extension lands. Kept here (not in formula.ts, which is the
 * CF-only evaluator and must stay untouched) so the expansion rule lives
 * beside its consumer.
 */
export function expandSharedFormula(
  anchorFormula: string,
  anchorRow: number,
  anchorCol: number,
  targetRow: number,
  targetCol: number,
): string {
  const dRow = targetRow - anchorRow;
  const dCol = targetCol - anchorCol;
  if (dRow === 0 && dCol === 0) return anchorFormula;
  let out = '';
  let i = 0;
  const n = anchorFormula.length;
  while (i < n) {
    const c = anchorFormula[i];
    if (c === '"') {
      const start = i;
      i++;
      while (i < n) {
        if (anchorFormula[i] === '"') {
          if (anchorFormula[i + 1] === '"') i += 2;
          else {
            i++;
            break;
          }
        } else i++;
      }
      out += anchorFormula.slice(start, i);
      continue;
    }
    const SheetRef = tryReadSheetQualifier(anchorFormula, i);
    const refStart = i + (SheetRef ? SheetRef.length : 0);
    const cellRef = tryReadCellRef(anchorFormula, refStart);
    if (cellRef) {
      // A function call (`LOG10(`) is a name, not a cell — but a sheet
      // qualifier already disambiguates (`Sheet!A1` is always a ref).
      const after = refStart + cellRef.length;
      if (!SheetRef && anchorFormula[after] === '(') {
        out += cellRef.text;
        i = after;
        continue;
      }
      const shifted = shiftCellRef(cellRef, dRow, dCol);
      out += (SheetRef ?? '') + shifted;
      i = after;
      continue;
    }
    out += c;
    i++;
  }
  return out;
}

interface ParsedCellRef {
  text: string;
  length: number;
  colAbs: boolean;
  col: number;
  rowAbs: boolean;
  row: number;
}

/** Read `'Quoted'!` or `Name!` at pos; null when absent. A bare cell ref
 *  (`A1!...`) is never a qualifier. */
function tryReadSheetQualifier(s: string, pos: number): string | null {
  const n = s.length;
  if (s[pos] === "'") {
    let j = pos + 1;
    while (j < n) {
      if (s[j] === "'") {
        if (s[j + 1] === "'") j += 2;
        else {
          j++;
          break;
        }
      } else j++;
    }
    if (j < n && s[j] === '!') return s.slice(pos, j + 1);
    return null;
  }
  if (isNameStart(s[pos] ?? '')) {
    let j = pos + 1;
    while (j < n && isNameChar(s[j])) j++;
    if (j < n && s[j] === '!' && j > pos) {
      const candidate = s.slice(pos, j);
      if (!isCellRef(candidate)) return candidate + '!';
    }
  }
  return null;
}

/** Read `$?[A-Z]+\$?[0-9]+` at pos (1-based col/row); null when the spelling
 *  is not an in-grid cell (mirrors formula.ts `tryParseCellRef` bounds). */
function tryReadCellRef(s: string, pos: number): ParsedCellRef | null {
  const n = s.length;
  let i = pos;
  let colAbs = false;
  let rowAbs = false;
  if (s[i] === '$') {
    colAbs = true;
    i++;
  }
  const colStart = i;
  while (i < n && isColLetter(s[i])) i++;
  if (i === colStart || i - colStart > 3) return null;
  const colLetters = s.slice(colStart, i).toUpperCase();
  if (s[i] === '$') {
    rowAbs = true;
    i++;
  }
  const rowStart = i;
  while (i < n && s[i] >= '0' && s[i] <= '9') i++;
  if (i === rowStart) return null;
  const text = s.slice(pos, i);
  // A trailing identifier char means this is a name prefix, not a ref.
  if (i < n && (isNameChar(s[i]) || s[i] === '!')) {
    // `!` after a cell spelling is invalid qualifier syntax — not a ref.
    if (s[i] === '!') return null;
    if (isNameChar(s[i])) return null;
  }
  const rowNum = parseInt(s.slice(rowStart, i), 10);
  let col = 0;
  for (let k = 0; k < colLetters.length; k++) col = col * 26 + (colLetters.charCodeAt(k) - 64);
  if (rowNum < 1 || rowNum > MAX_WORKSHEET_ROW || col < 1 || col > MAX_WORKSHEET_COL) return null;
  return { text, length: i - pos, colAbs, col, rowAbs, row: rowNum };
}

function shiftCellRef(ref: ParsedCellRef, dRow: number, dCol: number): string {
  const col = ref.colAbs ? ref.col : ref.col + dCol;
  const row = ref.rowAbs ? ref.row : ref.row + dRow;
  // Shifting off-grid (deleted rows/cols) is an invalid ref: Excel reports
  // `#REF!`. Emit the `#REF!` literal directly — it feeds back as an error
  // without depending on engine parsing.
  if (col < 1 || col > MAX_WORKSHEET_COL || row < 1 || row > MAX_WORKSHEET_ROW) return '#REF!';
  return `${ref.colAbs ? '$' : ''}${columnLetters(col)}${ref.rowAbs ? '$' : ''}${row}`;
}

function columnLetters(col: number): string {
  let letters = '';
  let c = col;
  while (c > 0) {
    const rem = (c - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    c = Math.floor((c - 1) / 26);
  }
  return letters;
}

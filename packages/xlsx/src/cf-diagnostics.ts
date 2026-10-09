import type { ViewportRange } from './types.js';

// ────────────────────────────────────────────────────────────────
// Conditional-formatting evaluation boundary (#1547)
//
// Library policy, not an Office behaviour. The CF evaluator implements a
// partial formula grammar over cached values. When a rule that the rendered
// frame reached cannot be evaluated, it neither formats nor stops lower
// rules, and that fact is recorded here instead of staying silent:
//
//   unsupported  the authored text is outside the implemented grammar,
//                function table, operand decoder or resource admission, or
//                needs a conversion this library has not established.
//   invalid      the rule is structurally incomplete or malformed (missing
//                operand, operator outside ST_ConditionalFormattingOperator,
//                `formula` threshold without text, malformed expression).
//
// A record names the rule by document position and the first cell at which
// this frame observed the condition. It never copies formula text, cell text
// or values. Records describe only cells the render evaluated: an empty list
// is not a certification that the sheet's conditional formatting is fully
// supported. Some limitations are data dependent (range admission, ROUND
// digits from a cell, ...), so no static per-rule classification is claimed.
// ────────────────────────────────────────────────────────────────

export type CfDiagnosticKind = 'unsupported' | 'invalid';
export type CfDiagnosticPhase = 'expression' | 'cellIs' | 'activity' | 'threshold';

export interface CfRuleDiagnostic {
  readonly kind: CfDiagnosticKind;
  /** `expression` rule formula; `cellIs` operator/operands; the
   *  colorScale/dataBar/iconSet activity formula; their cfvo thresholds. */
  readonly phase: CfDiagnosticPhase;
  /** Index into `Worksheet.conditionalFormats`. */
  readonly blockIndex: number;
  /** Index into that block's `rules` (document order, not priority order). */
  readonly ruleIndex: number;
  /** 1-based cell at which the frame first observed this condition. */
  readonly row: number;
  readonly col: number;
}

export interface XlsxConditionalFormattingReport {
  readonly sheetIndex: number;
  /** Requested scrollable viewport; frozen panes painted with it are included. */
  readonly viewport: Readonly<ViewportRange>;
  readonly diagnostics: readonly CfRuleDiagnostic[];
}

/** Rule identity used by the collector (a compiled CF rule satisfies it). */
export interface CfDiagnosticRuleKey {
  readonly blockIndex: number;
  readonly ruleIndex: number;
}

const PHASE_ORDER: Readonly<Record<CfDiagnosticPhase, number>> = {
  expression: 0, cellIs: 1, activity: 2, threshold: 3,
};
const KIND_ORDER: Readonly<Record<CfDiagnosticKind, number>> = { unsupported: 0, invalid: 1 };

/**
 * Transient collector for ONE render invocation. Never attach it to a cached
 * CfContext or worksheet: concurrent renders would swap each other's report.
 * Storage is bounded by the admitted rule count: one 8-bit mask per rule
 * reached and at most 2 kinds x 4 phases fixed-size records per rule; there
 * is no per-cell list, so repeated viewport cells cannot grow the payload.
 */
export class CfDiagnosticCollector {
  private readonly seen = new Map<CfDiagnosticRuleKey, number>();
  private readonly records: CfRuleDiagnostic[] = [];
  private paintedFrame = false;

  record(
    rule: CfDiagnosticRuleKey,
    phase: CfDiagnosticPhase,
    kind: CfDiagnosticKind,
    row: number,
    col: number,
  ): void {
    const bit = 1 << (PHASE_ORDER[phase] * 2 + KIND_ORDER[kind]);
    const mask = this.seen.get(rule) ?? 0;
    if (mask & bit) return;
    this.seen.set(rule, mask | bit);
    this.records.push({
      kind, phase, blockIndex: rule.blockIndex, ruleIndex: rule.ruleIndex, row, col,
    });
  }

  /** Called by the render orchestrator once the frame passed its commit guard
   *  and is about to paint. A frame that never painted has no report. */
  markPainted(): void {
    this.paintedFrame = true;
  }

  get painted(): boolean {
    return this.paintedFrame;
  }

  /** Deterministic, deeply frozen plain records (structured-clone safe). */
  snapshot(): readonly CfRuleDiagnostic[] {
    return canonicalDiagnostics(this.records);
  }
}

function compareDiagnostics(a: CfRuleDiagnostic, b: CfRuleDiagnostic): number {
  return a.blockIndex - b.blockIndex
    || a.ruleIndex - b.ruleIndex
    || PHASE_ORDER[a.phase] - PHASE_ORDER[b.phase]
    || KIND_ORDER[a.kind] - KIND_ORDER[b.kind];
}

/** Sort by rule/phase/kind; the stable sort keeps the first observed cell. */
function canonicalDiagnostics(records: readonly CfRuleDiagnostic[]): readonly CfRuleDiagnostic[] {
  const sorted = [...records].sort(compareDiagnostics);
  const out: CfRuleDiagnostic[] = [];
  for (const record of sorted) {
    const last = out[out.length - 1];
    if (last && compareDiagnostics(last, record) === 0) continue;
    out.push(Object.freeze({
      kind: record.kind,
      phase: record.phase,
      blockIndex: record.blockIndex,
      ruleIndex: record.ruleIndex,
      row: record.row,
      col: record.col,
    }));
  }
  return Object.freeze(out);
}

function isKind(value: unknown): value is CfDiagnosticKind {
  return value === 'unsupported' || value === 'invalid';
}

function isPhase(value: unknown): value is CfDiagnosticPhase {
  return typeof value === 'string' && Object.hasOwn(PHASE_ORDER, value);
}

function isIndex(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;
}

function isCoordinate(value: unknown): value is number {
  return isIndex(value) && value >= 1;
}

/** Validate a worker batch. Rejects malformed or duplicated records instead
 *  of silently truncating or repairing them. */
export function decodeCfDiagnosticsWire(value: unknown): readonly CfRuleDiagnostic[] {
  if (!Array.isArray(value)) {
    throw new TypeError('XLSX render worker returned invalid conditional-formatting diagnostics');
  }
  const records: CfRuleDiagnostic[] = [];
  for (const item of value) {
    if (!item || typeof item !== 'object') {
      throw new TypeError('XLSX render worker returned an invalid conditional-formatting diagnostic');
    }
    const { kind, phase, blockIndex, ruleIndex, row, col } = item as Record<string, unknown>;
    if (!isKind(kind) || !isPhase(phase) || !isIndex(blockIndex) || !isIndex(ruleIndex)
        || !isCoordinate(row) || !isCoordinate(col)) {
      throw new TypeError('XLSX render worker returned an invalid conditional-formatting diagnostic');
    }
    records.push({ kind, phase, blockIndex, ruleIndex, row, col });
  }
  const canonical = canonicalDiagnostics(records);
  if (canonical.length !== records.length) {
    throw new TypeError('XLSX render worker returned duplicated conditional-formatting diagnostics');
  }
  return canonical;
}

export function createCfReport(
  sheetIndex: number,
  viewport: ViewportRange,
  diagnostics: readonly CfRuleDiagnostic[],
): XlsxConditionalFormattingReport {
  return Object.freeze({
    sheetIndex,
    viewport: Object.freeze({
      row: viewport.row, col: viewport.col, rows: viewport.rows, cols: viewport.cols,
    }),
    diagnostics: canonicalDiagnostics(diagnostics),
  });
}

/** Internal per-invocation report delivery (Viewer only). Symbol-keyed so it
 *  never appears in public options and is removed before any worker post. */
export type CfReportSink = (report: XlsxConditionalFormattingReport) => void;
export const XLSX_CF_REPORT_SINK: unique symbol = Symbol('xlsx-cf-report-sink');
type WithCfReportSink = { [XLSX_CF_REPORT_SINK]?: CfReportSink };

export function withCfReportSink<T extends object>(options: T, sink: CfReportSink): T {
  return { ...options, [XLSX_CF_REPORT_SINK]: sink } as T;
}

export function takeCfReportSink<T extends object>(
  options: T,
): { options: T; sink: CfReportSink | undefined } {
  const sink = (options as WithCfReportSink)[XLSX_CF_REPORT_SINK];
  if (sink === undefined) return { options, sink: undefined };
  const rest = { ...options } as T & WithCfReportSink;
  delete rest[XLSX_CF_REPORT_SINK];
  return { options: rest, sink };
}

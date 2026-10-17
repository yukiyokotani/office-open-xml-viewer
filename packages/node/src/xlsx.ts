import type { ParsedWorkbook, Row, Worksheet } from '@silurus/ooxml-xlsx';
import type { OoxmlResourceUsageSnapshot } from '@silurus/ooxml-core';
import {
  decodeOoxmlResourceUsage,
  parseResourceLimitError,
  type OoxmlResourceMetricsSession,
  type PullSessionCommand,
  type PullSessionResponse,
} from '@silurus/ooxml-core/worker';
// The Node facade consumes the format-owned coordinator through its explicit
// internal entry point rather than reconstructing XLSX orchestration here.
import {
  DEFAULT_XLSX_WORKSHEET_POLICY,
  normalizeXlsxWorksheetPolicy,
  type NormalizedXlsxWorksheetPolicy,
  acquireXlsxNodeSession,
  addWorksheetCacheUsage,
  addWorksheetUsage,
  assertWorksheetCacheUsage,
  assertWorksheetJsonBytes,
  assertWorksheetModelUsage,
  bindWorksheetPolicy,
  completeWorksheetUsage,
  measureRows,
  XlsxWorksheetPullClient,
  WorksheetPullWorker,
  type WorksheetCacheUsage,
  type WorksheetModelUsage,
  type XlsxNodeArchive,
} from '@silurus/ooxml-xlsx/internal/session';
import { InProcessPullTransport } from '@silurus/ooxml-core/internal/in-process-pull-transport';
import type { OoxmlNodeSessionOptions } from './session-options.ts';
import { createLazyWasmModule, resolveWasm } from './wasm-loader.ts';
import { usingOwnedSession } from '@silurus/ooxml-core/internal/owned-session';

const getXlsxWasmModule = createLazyWasmModule(() => resolveWasm(
    import.meta.url,
    'xlsx_parser_bg.wasm',
    '@silurus/ooxml-xlsx/wasm-binary',
  ));

export type DeepReadonly<T> =
  T extends (...args: never[]) => unknown ? T
    : T extends readonly (infer TValue)[] ? readonly DeepReadonly<TValue>[]
      : T extends object ? { readonly [TKey in keyof T]: DeepReadonly<T[TKey]> }
        : T;

export type ReadonlyParsedWorkbook = DeepReadonly<ParsedWorkbook>;

export interface MaterializedXlsxWorkbook {
  readonly workbookIndex: ParsedWorkbook;
  /** Caller-owned worksheets in workbook sheet-index order. */
  readonly worksheets: readonly Worksheet[];
}

/** Options for the bounded Node workbook session. */
export type OpenXlsxWorkbookOptions = OoxmlNodeSessionOptions & {
  /** Canvas for a selected model source's Normal-font measurement. */
  readonly factory?: import('./render.ts').NodeCanvasFactory;
};

export type XlsxWorksheetRowChunk =
  | {
      readonly kind: 'rows';
      readonly rows: Row[];
      readonly sequence: number;
      readonly wireBytes: number;
      readonly usage?: OoxmlResourceUsageSnapshot;
    }
  | {
      /** The terminal worksheet contains layout/tail data and always has zero rows. */
      readonly kind: 'finished';
      readonly worksheet: Worksheet;
      readonly sequence: number;
      readonly wireBytes: number;
      readonly usage?: OoxmlResourceUsageSnapshot;
    };

/**
 * An explicitly-owned Node workbook session. The workbook index and shared
 * strings are parsed once; worksheet bodies are pulled as bounded row batches
 * from the retained archive. Only one worksheet stream may be active at a time
 * because the native archive owns one worksheet cursor.
 */
export interface XlsxWorkbookSession {
  readonly workbookIndex: ReadonlyParsedWorkbook;
  readonly sheetCount: number;
  readonly sheetNames: ReadonlyArray<string>;
  readonly resourceUsage: OoxmlResourceUsageSnapshot | undefined;
  worksheetRows(sheetIndex: number): AsyncGenerator<XlsxWorksheetRowChunk, void, void>;
  /** Read one bounded embedded image after the row stream terminal ACK. */
  getImage(path: string, mimeType: string): Promise<Blob>;
  /** Bounded raw OOXML part for trusted relationship/content inventory. */
  readPart(path: string): Promise<Uint8Array>;
  close(): Promise<void>;
}

/**
 * Open a bounded XLSX workbook session. Unlike the materializing compatibility
 * helpers, this retains one archive, parses the workbook index once, and lets
 * callers consume worksheet rows sequentially without reopening the package.
 */
export async function openXlsxWorkbook(
  buffer: ArrayBuffer | Uint8Array,
  options: OpenXlsxWorkbookOptions = {},
): Promise<XlsxWorkbookSession> {
  const policy = normalizeXlsxWorksheetPolicy(options);
  // Rebind the local parameter to an immutable policy capture so later caller
  // mutation cannot affect this session, while keeping the canonical
  // `options.modelSources !== undefined` branch shape recognized by the
  // source-boundary gate.
  options = { ...options, xlsxWorksheetLimits: policy.worksheet };
  if (__OOXML_MODEL_SOURCES__ && options.modelSources !== undefined) {
    const { openXlsxSource } = await import('./xlsx-model-source.ts');
    return openXlsxSource(buffer, options, getXlsxWasmModule);
  }
  const bytes = toUint8(buffer);
  const acquired = await acquireXlsxNodeSession(bytes, getXlsxWasmModule(), options);
  return new XlsxWorkbookSessionImpl(
    acquired.closeArchive,
    acquired.archive as XlsxNodeArchive,
    acquired.workbookIndex,
    acquired.metrics,
    acquired.usage,
    options.signal,
    acquired.worksheetPolicy,
  );
}

type ActiveWorksheetOperation = {
  cleanupPromise?: Promise<void>;
};

/** Realm-local, non-public per-session normalized worksheet policy binding. */
const sessionWorksheetPolicies = new WeakMap<object, NormalizedXlsxWorksheetPolicy>();

function sessionWorksheetPolicy(session: object): NormalizedXlsxWorksheetPolicy {
  return sessionWorksheetPolicies.get(session) ?? DEFAULT_XLSX_WORKSHEET_POLICY;
}

export class XlsxWorkbookSessionImpl implements XlsxWorkbookSession {
  readonly workbookIndex: ReadonlyParsedWorkbook;
  readonly sheetCount: number;
  readonly sheetNames: ReadonlyArray<string>;

  protected pull: Pick<WorksheetPullWorker, 'dispatchSafely' | 'open' | 'reserveOpen' | 'reset'>;
  private readonly transport: InProcessPullTransport<PullSessionResponse<ArrayBuffer, number>>;
  private readonly worksheetPullClient: XlsxWorksheetPullClient;
  private active: ActiveWorksheetOperation | undefined;
  private closed = false;
  private closePromise: Promise<void> | undefined;
  private lastUsage: OoxmlResourceUsageSnapshot | undefined;
  private completedWorksheets = 0;
  private rowBatches = 0;
  private emittedRows = 0;

  constructor(
    private readonly closeArchive: () => void,
    private readonly archive: XlsxNodeArchive,
    workbook: ParsedWorkbook,
    private readonly metrics: OoxmlResourceMetricsSession,
    usage: OoxmlResourceUsageSnapshot | undefined,
    private readonly signal?: AbortSignal,
    policy: NormalizedXlsxWorksheetPolicy = DEFAULT_XLSX_WORKSHEET_POLICY,
  ) {
    sessionWorksheetPolicies.set(this, policy);
    this.workbookIndex = freezeRecursively(workbook);
    this.sheetNames = Object.freeze(this.workbookIndex.workbook.sheets.map((sheet) => sheet.name));
    this.sheetCount = this.sheetNames.length;
    this.lastUsage = usage;
    this.pull = new WorksheetPullWorker(() => this.archive);
    this.transport = new InProcessPullTransport(
      (command, respond) => this.pull.dispatchSafely(
        command as PullSessionCommand<number>,
        respond,
      ),
      () => undefined,
    );
    this.worksheetPullClient = new XlsxWorksheetPullClient({
      transport: this.transport,
      sharedStrings: workbook.sharedStrings,
      open: async (sheetIndex, sheetName, identity) => {
        this.pull.reserveOpen(identity);
        await this.pull.open(sheetIndex, sheetName, identity);
      },
      onUsage: (nextUsage) => {
        this.lastUsage = nextUsage;
        this.metrics.observeUsage(nextUsage);
      },
    });
  }

  get resourceUsage(): OoxmlResourceUsageSnapshot | undefined {
    if (this.closed) return this.lastUsage;
    try {
      this.lastUsage = decodeUsage(this.archive.resource_usage());
    } catch {
      // Keep the last valid diagnostic after a trapped or closing archive.
    }
    return this.lastUsage;
  }

  async getImage(path: string, mimeType: string): Promise<Blob> {
    const bytes = await this.readPart(path);
    if (bytes.byteLength === 0) throw new Error('XLSX image extraction returned no bytes');
    return new Blob([bytes as BlobPart], { type: mimeType });
  }

  async readPart(path: string): Promise<Uint8Array> {
    if (this.closed) throw new Error('XLSX workbook session is closed');
    if (this.active) throw new Error('finish the XLSX worksheet row stream before reading images');
    if (this.signal?.aborted) {
      await this.close();
      const error = new Error('XLSX workbook session was aborted');
      error.name = 'AbortError';
      throw error;
    }
    try {
      // Reuse the retained, quota-enforced archive; no second ZIP or raw cache.
      const bytes = this.archive.extract_image(path);
      this.lastUsage = decodeUsage(this.archive.resource_usage());
      this.metrics.observeUsage(this.lastUsage);
      return bytes;
    } catch (error) {
      const normalized = parseResourceLimitError(error) ?? error;
      this.metrics.fail(normalized);
      await this.close().catch(() => undefined);
      throw normalized;
    }
  }

  /**
   * Stream one worksheet as bounded complete-row batches. The final yielded
   * value is a row-free worksheet containing layout and ancillary model data.
   * Row batches remain provisional until the consumer advances past that
   * terminal value, which acknowledges the native operation.
   */
  async *worksheetRows(
    sheetIndex: number,
  ): AsyncGenerator<XlsxWorksheetRowChunk, void, void> {
    if (this.closed) throw new Error('XLSX workbook session is closed');
    if (this.active) throw new Error('another XLSX worksheet row stream is already active');
    if (!Number.isSafeInteger(sheetIndex) || sheetIndex < 0) {
      throw new RangeError('sheetIndex must be a non-negative safe integer');
    }
    const sheetName = this.requireSheetName(sheetIndex);
    const operation: ActiveWorksheetOperation = {};
    this.active = operation;
    let operationError: unknown;
    try {
      for await (const unit of this.worksheetPullClient.stream(
        sheetIndex,
        sheetName,
        this.signal,
      )) {
        if (this.closed) throw new Error('XLSX workbook session is closed');
        // Browser viewers may paint from this provisional shell. The Node row
        // stream keeps its established rows-then-terminal contract.
        if (unit.kind === 'preview') continue;
        if (unit.kind === 'rows') {
          this.rowBatches += 1;
          this.emittedRows += unit.rows.length;
          yield {
            kind: 'rows',
            rows: unit.rows,
            sequence: unit.sequence,
            wireBytes: unit.wireBytes,
            usage: unit.usage,
          };
          continue;
        }
        bindWorksheetPolicy(unit.worksheet, sessionWorksheetPolicy(this));
        yield {
          kind: 'finished',
          worksheet: unit.worksheet,
          sequence: unit.sequence,
          wireBytes: unit.wireBytes,
          usage: unit.usage,
        };
      }
      this.completedWorksheets += 1;
      this.metrics.checkpoint('worksheet stream complete', this.lastUsage);
    } catch (error) {
      operationError = parseResourceLimitError(error) ?? error;
      this.metrics.fail(operationError);
      await this.close().catch(() => undefined);
      throw operationError;
    } finally {
      try {
        await this.cleanupOperation(operation, operationError === undefined ? 'closed' : 'request-error');
      } catch (cleanupError) {
        if (operationError === undefined) {
          const normalized = parseResourceLimitError(cleanupError) ?? cleanupError;
          this.metrics.fail(normalized);
          await this.close().catch(() => undefined);
          throw normalized;
        }
      }
    }
  }

  close(): Promise<void> {
    if (this.closePromise) return this.closePromise;
    this.closed = true;
    this.closePromise = this.release();
    return this.closePromise;
  }

  private cleanupOperation(
    operation: ActiveWorksheetOperation,
    reason: 'closed' | 'request-error',
  ): Promise<void> {
    if (operation.cleanupPromise) return operation.cleanupPromise;
    operation.cleanupPromise = (async () => {
      let cleanupError: unknown;
      try {
        await this.worksheetPullClient.cancelAll(reason);
      } catch (error) {
        cleanupError = parseResourceLimitError(error) ?? error;
      }
      try {
        await this.pull.reset();
      } catch (error) {
        cleanupError ??= parseResourceLimitError(error) ?? error;
      }
      if (this.active === operation) this.active = undefined;
      if (cleanupError !== undefined) throw cleanupError;
    })();
    return operation.cleanupPromise;
  }

  private async release(): Promise<void> {
    let cleanupError: unknown;
    if (this.active) {
      try {
        await this.cleanupOperation(this.active, 'closed');
      } catch (error) {
        cleanupError = parseResourceLimitError(error) ?? error;
      }
    }
    this.transport.terminate();
    try {
      this.closeArchive();
    } catch (error) {
      cleanupError ??= parseResourceLimitError(error) ?? error;
    }
    if (cleanupError !== undefined) {
      this.metrics.fail(cleanupError);
      throw cleanupError;
    }
    this.metrics.checkpoint('workbook session closed', this.lastUsage);
    this.metrics.succeed({
      worksheets: this.completedWorksheets,
      'row-batches': this.rowBatches,
      rows: this.emittedRows,
    });
  }

  private requireSheetName(sheetIndex: number): string {
    const sheet = this.workbookIndex.workbook.sheets[sheetIndex];
    if (!sheet) throw new RangeError(`Sheet index ${sheetIndex} out of range`);
    return sheet.name;
  }
}

/** Materialize only the workbook metadata/index through the canonical owned
 * archive. No worksheet cursor is opened. */
export async function materializeXlsxWorkbookIndex(
  buffer: ArrayBuffer | Uint8Array,
  options: OpenXlsxWorkbookOptions = {},
): Promise<ParsedWorkbook> {
  return usingOwnedSession(
    () => openXlsxWorkbook(buffer, options),
    async (session) => structuredClone(session.workbookIndex) as ParsedWorkbook,
  );
}

/** Materialize one caller-owned worksheet by sheet index. */
export async function materializeXlsxWorksheet(
  buffer: ArrayBuffer | Uint8Array,
  sheetIndex: number,
  options: OpenXlsxWorkbookOptions = {},
): Promise<Worksheet> {
  return usingOwnedSession(
    () => openXlsxWorkbook(buffer, options),
    async (session) => (await materializeWorksheetFromSession(session, sheetIndex)).worksheet,
  );
}

/** Materialize the complete workbook without reopening the package per sheet. */
export async function materializeXlsxWorkbook(
  buffer: ArrayBuffer | Uint8Array,
  options: OpenXlsxWorkbookOptions = {},
): Promise<MaterializedXlsxWorkbook> {
  return usingOwnedSession(
    () => openXlsxWorkbook(buffer, options),
    async (session) => {
      const policy = sessionWorksheetPolicy(session);
      const worksheets: Worksheet[] = [];
      let retainedUsage: WorksheetCacheUsage = {
        rows: 0, cells: 0, ownedUtf8Bytes: 0, jsonBytes: 0,
      };
      for (let sheetIndex = 0; sheetIndex < session.sheetCount; sheetIndex += 1) {
        const materialized = await materializeWorksheetFromSession(session, sheetIndex);
        const nextUsage = addWorksheetCacheUsage(retainedUsage, materialized.usage, {}, policy);
        assertWorksheetCacheUsage(
          nextUsage,
          'materialize-workbook',
          undefined,
          materialized.resourceUsage,
          policy,
        );
        retainedUsage = nextUsage;
        worksheets.push(materialized.worksheet);
      }
      return {
        workbookIndex: structuredClone(session.workbookIndex) as ParsedWorkbook,
        worksheets: Object.freeze(worksheets),
      };
    },
  );
}

async function materializeWorksheetFromSession(
  session: XlsxWorkbookSession,
  sheetIndex: number,
): Promise<{
  worksheet: Worksheet;
  usage: WorksheetModelUsage & { jsonBytes: number };
  resourceUsage?: OoxmlResourceUsageSnapshot;
}> {
  const policy = sessionWorksheetPolicy(session);
  const rows: Row[] = [];
  let modelUsage: WorksheetModelUsage = { rows: 0, cells: 0, ownedUtf8Bytes: 0 };
  let retainedUsage: (WorksheetModelUsage & { jsonBytes: number }) | undefined;
  let terminal: Worksheet | undefined;
  let resourceUsage: OoxmlResourceUsageSnapshot | undefined;
  for await (const chunk of session.worksheetRows(sheetIndex)) {
    resourceUsage = chunk.usage ?? resourceUsage;
    if (chunk.kind === 'rows') {
      const nextUsage = addWorksheetUsage(modelUsage, measureRows(chunk.rows, policy), policy);
      assertWorksheetModelUsage(
        nextUsage,
        'materialize-worksheet',
        undefined,
        resourceUsage,
        policy,
      );
      rows.push(...chunk.rows);
      modelUsage = nextUsage;
    } else {
      terminal = chunk.worksheet;
      bindWorksheetPolicy(terminal, policy);
      terminal.rows = terminal.parseError ? [] : rows;
      const measured = completeWorksheetUsage(
        terminal,
        terminal.parseError
          ? { rows: 0, cells: 0, ownedUtf8Bytes: 0 }
          : modelUsage,
        policy,
      );
      assertWorksheetModelUsage(
        measured,
        'materialize-worksheet',
        undefined,
        resourceUsage,
        policy,
      );
      assertWorksheetJsonBytes(
        measured.jsonBytes,
        'materialize-worksheet',
        undefined,
        resourceUsage,
        policy,
      );
      modelUsage = measured;
      retainedUsage = measured;
    }
  }
  if (!terminal || !retainedUsage) {
    throw new Error(`XLSX worksheet ${sheetIndex} did not produce a terminal model`);
  }
  return {
    // The coordinator decoded this terminal graph and the retained row units
    // exclusively for this materializer. Transfer ownership directly instead
    // of doubling the complete worksheet graph at the session boundary.
    worksheet: terminal,
    usage: retainedUsage,
    resourceUsage,
  };
}

function freezeRecursively<T>(value: T, seen = new WeakSet<object>()): DeepReadonly<T> {
  if (value === null || typeof value !== 'object') return value as DeepReadonly<T>;
  const object = value as object;
  if (seen.has(object)) return value as DeepReadonly<T>;
  seen.add(object);
  for (const child of Object.values(object)) freezeRecursively(child, seen);
  return Object.freeze(value) as DeepReadonly<T>;
}

function decodeUsage(bytes: Uint8Array): OoxmlResourceUsageSnapshot | undefined {
  try {
    return decodeOoxmlResourceUsage(bytes);
  } catch (error) {
    if (String(error).includes('worksheet cursor usage is unavailable')) return undefined;
    throw error;
  }
}

function toUint8(buffer: ArrayBuffer | Uint8Array): Uint8Array {
  return buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer as ArrayBuffer);
}
declare const __OOXML_MODEL_SOURCES__: boolean;

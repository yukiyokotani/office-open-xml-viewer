import { OoxmlResourceLimitError, type OoxmlResourceUsageSnapshot } from '@silurus/ooxml-core';
import {
  decodeOoxmlResourceUsage,
  exactTransferableArrayBuffer,
  PULL_SESSION_PROTOCOL,
  PullSessionHost,
  PullSessionHostCoordinator,
  serializeWorkerError,
  type PullSessionCommand,
  type PullSessionIdentity,
  type PullSessionResponse,
} from '@silurus/ooxml-core/worker';
import type { Row, Worksheet } from './types.js';
import { decodeWorksheetPullChunk } from './worksheet-pull-codec.js';
import {
  addWorksheetUsage,
  assertWorksheetJsonBytes,
  assertWorksheetModelUsage,
  completeWorksheetUsage,
  measureRows,
  type WorksheetCacheUsage,
  type WorksheetModelUsage,
} from './worksheet-resource-limits.js';

export const XLSX_WORKSHEET_PULL_BYTES = 64 * 1024 * 1024;
export const XLSX_WORKSHEET_PULL_ROWS = 128;

export interface WorksheetCursorArchive {
  open_sheet_cursor(sheetIndex: number, name: string): void;
  pull_sheet_cursor(rowCredit: number): Uint8Array;
  sheet_cursor_pull_finished(): boolean;
  sheet_cursor_resource_usage(): Uint8Array;
  acknowledge_sheet_cursor_terminal(): void;
  cancel_sheet_cursor(): void;
  close_sheet_cursor(): void;
}

export type { WorksheetWireChunk } from './worksheet-pull-codec.js';

/** Format-owned XLSX driver composed with core's shared pull state machine. */
export class WorksheetPullWorker {
  readonly coordinator = new PullSessionHostCoordinator();
  private readonly sessions = new Map<number, {
    host: PullSessionHost<ArrayBuffer, number>;
    identity: PullSessionIdentity<number>;
  }>();
  private operationTail: Promise<void> = Promise.resolve();
  private readonly pendingOpens = new Map<number, {
    identity: PullSessionIdentity<number>;
    canceled: boolean;
  }>();
  private resourceFailure: OoxmlResourceLimitError | undefined;

  constructor(
    private readonly archive: () => WorksheetCursorArchive | null | undefined,
    private readonly acceptWorksheet?: (
      sheetIndex: number,
      worksheet: Worksheet,
      modelUsage: WorksheetCacheUsage,
      usage?: OoxmlResourceUsageSnapshot,
    ) => void | (() => void) | { rollback?: () => void; commit?: () => void },
    private readonly executeArchive: <T>(operation: (archive: WorksheetCursorArchive) => T) => T =
      (operation) => operation(this.requireArchive()),
    private readonly prepareRows?: (rows: Row[]) => void,
  ) {}

  /** Register synchronously before a worker handler's first await. */
  reserveOpen(identity: PullSessionIdentity<number>): void {
    this.pendingOpens.set(identity.sessionId, { identity, canceled: false });
  }

  abandonOpen(sessionId: number): void {
    this.pendingOpens.delete(sessionId);
  }

  get pendingOpenCount(): number {
    return this.pendingOpens.size;
  }

  async open(
    sheetIndex: number,
    name: string,
    identity: PullSessionIdentity<number>,
  ): Promise<void> {
    if (this.resourceFailure) throw this.resourceFailure;
    const pending = this.pendingOpens.get(identity.sessionId);
    if (
      !pending ||
      pending.identity.operationId !== identity.operationId ||
      pending.identity.generation !== identity.generation
    ) {
      throw new Error('worksheet pull session open reservation is stale or missing');
    }
    let completeOperation!: () => void;
    const completion = new Promise<void>((resolve) => {
      completeOperation = resolve;
    });
    const started = this.operationTail.then(() => this.coordinator.enqueue(async () => {
      if (pending.canceled) throw new Error('worksheet pull session open was canceled');
      this.executeArchive((archive) => archive.open_sheet_cursor(sheetIndex, name));
      const rows: Row[] = [];
      let modelUsage: WorksheetModelUsage = { rows: 0, cells: 0, ownedUtf8Bytes: 0 };
      let terminal: Worksheet | undefined;
      let terminalPending = false;
      const session = new PullSessionHost<ArrayBuffer, number>({
        ...identity,
        maxByteCredit: XLSX_WORKSHEET_PULL_BYTES,
        coordinator: this.coordinator,
        driver: {
          pull: () => {
            const bytes = this.executeArchive((archive) =>
              archive.pull_sheet_cursor(XLSX_WORKSHEET_PULL_ROWS));
            const done = this.executeArchive((archive) => archive.sheet_cursor_pull_finished());
            if (this.acceptWorksheet) {
              const decoded = decodeWorksheetPullChunk(bytes, done, undefined, this.prepareRows);
              try {
                if (decoded.kind === 'rows') {
                  const next = addWorksheetUsage(modelUsage, measureRows(decoded.rows));
                  assertWorksheetModelUsage(
                    next,
                    'get-worksheet-worker',
                    undefined,
                    this.readResourceUsage(),
                  );
                  rows.push(...decoded.rows);
                  modelUsage = next;
                } else terminal = decoded.worksheet;
              } catch (error) {
                if (error instanceof OoxmlResourceLimitError) this.resourceFailure ??= error;
                throw error;
              }
            }
            terminalPending = done;
            const payload = exactTransferableArrayBuffer(bytes);
            return { payload, byteLength: payload.byteLength, done, transfer: [payload] };
          },
          measureChunk: ({ payload }) => payload.byteLength,
          acknowledge: () => {
            if (!terminalPending) return;
            // Main has accepted the terminal transfer before sending this ACK.
            // Render-worker mode must also accept its local retained model
            // before Rust commits, so a cache/assembly failure remains cancelable.
            let rollback: (() => void) | undefined;
            let commit: (() => void) | undefined;
            try {
              if (this.acceptWorksheet) {
                if (!terminal) throw new Error('worksheet terminal payload is missing');
                terminal.rows = terminal.parseError ? [] : rows;
                const retainedModelUsage = terminal.parseError
                  ? { rows: 0, cells: 0, ownedUtf8Bytes: 0 }
                  : modelUsage;
                const measured = completeWorksheetUsage(terminal, retainedModelUsage);
                const resourceUsage = this.readResourceUsage();
                assertWorksheetModelUsage(
                  measured,
                  'get-worksheet-worker',
                  undefined,
                  resourceUsage,
                );
                assertWorksheetJsonBytes(
                  measured.jsonBytes,
                  'get-worksheet-worker',
                  undefined,
                  resourceUsage,
                );
                const accepted = this.acceptWorksheet(
                  sheetIndex,
                  terminal,
                  measured,
                  resourceUsage,
                );
                if (typeof accepted === 'function') rollback = accepted;
                else if (accepted) ({ rollback, commit } = accepted);
              }
              this.executeArchive((archive) => archive.acknowledge_sheet_cursor_terminal());
              commit?.();
            } catch (error) {
              rollback?.();
              if (error instanceof OoxmlResourceLimitError) this.resourceFailure ??= error;
              throw error;
            }
            terminalPending = false;
            this.sessions.delete(identity.sessionId);
            completeOperation();
          },
          cancel: () => {
            try {
              if (this.archive()) {
                this.executeArchive((archive) => archive.cancel_sheet_cursor());
              }
            } finally {
              this.sessions.delete(identity.sessionId);
              completeOperation();
            }
          },
          close: () => {
            try {
              if (this.archive()) {
                this.executeArchive((archive) => archive.close_sheet_cursor());
              }
            } finally {
              this.sessions.delete(identity.sessionId);
              completeOperation();
            }
          },
          resourceUsage: () =>
            this.readResourceUsage(),
        },
      });
      this.sessions.set(identity.sessionId, { host: session, identity });
      this.pendingOpens.delete(identity.sessionId);
    }));
    this.operationTail = started.then(() => completion, () => undefined);
    try {
      await started;
    } catch (error) {
      this.pendingOpens.delete(identity.sessionId);
      completeOperation();
      throw error;
    }
  }

  /**
   * Deliver an opened response, closing the just-opened shared lifecycle when
   * neither that response nor its plain fallback error can reach main.
   */
  async postOpenedSafely(
    identity: PullSessionIdentity<number>,
    postOpened: () => void,
    postError: (error: unknown) => void,
  ): Promise<void> {
    try {
      postOpened();
    } catch (error) {
      await this.closeIdentity(identity);
      try {
        postError(error);
      } catch {
        // The response channel is unavailable. Lifecycle cleanup already
        // released the cursor and package-operation FIFO.
      }
    }
  }

  dispatch(
    command: PullSessionCommand<number>,
    post: (response: PullSessionResponse<ArrayBuffer, number>, transfer?: Transferable[]) => void,
  ): Promise<void> {
    const session = this.sessions.get(command.sessionId);
    if (session) return session.host.dispatch(command, post);
    const pending = this.pendingOpens.get(command.sessionId);
    if (pending && (command.kind === 'cancel' || command.kind === 'close')) {
      const matches = pending.identity.operationId === command.operationId &&
        pending.identity.generation === command.generation;
      if (matches) pending.canceled = true;
      post(matches
        ? {
            protocol: PULL_SESSION_PROTOCOL,
            kind: 'accepted',
            sessionId: command.sessionId,
            operationId: command.operationId,
            generation: command.generation,
            requestId: command.requestId,
            command: command.kind,
          }
        : {
            protocol: PULL_SESSION_PROTOCOL,
            kind: 'error',
            sessionId: command.sessionId,
            operationId: command.operationId,
            generation: command.generation,
            requestId: command.requestId,
            error: {
              message: 'stale lifecycle targets another pending worksheet operation',
              errorName: 'PullSessionProtocolError',
              code: 'ooxml-stale-lifecycle',
            },
          });
      return Promise.resolve();
    }
    if (command.kind === 'cancel' || command.kind === 'close') {
      post({
        protocol: PULL_SESSION_PROTOCOL,
        kind: 'accepted',
        sessionId: command.sessionId,
        operationId: command.operationId,
        generation: command.generation,
        requestId: command.requestId,
        command: command.kind,
      });
      return Promise.resolve();
    }
    post({
      protocol: PULL_SESSION_PROTOCOL,
      kind: 'error',
      sessionId: command.sessionId,
      operationId: command.operationId,
      generation: command.generation,
      requestId: command.requestId,
      error: serializeWorkerError(new Error('worksheet pull session is not open')),
    });
    return Promise.resolve();
  }

  /** PullSessionHost already rolls back ownership when posting throws. */
  async dispatchSafely(
    command: PullSessionCommand<number>,
    post: (response: PullSessionResponse<ArrayBuffer, number>, transfer?: Transferable[]) => void,
  ): Promise<void> {
    try {
      await this.dispatch(command, post);
    } catch (error) {
      // The data-bearing post may have failed structured clone after the host
      // rolled ownership back. A plain error envelope can still settle main.
      try {
        post({
          protocol: PULL_SESSION_PROTOCOL,
          kind: 'error',
          sessionId: command.sessionId,
          operationId: command.operationId,
          generation: command.generation,
          requestId: command.requestId,
          error: serializeWorkerError(error),
        });
      } catch {
        // If even the plain envelope cannot be posted, host cleanup is already
        // complete and the worker's error/messageerror containment must settle main.
      }
    }
  }

  run<T>(operation: () => T | Promise<T>): Promise<T> {
    const result = this.operationTail.then(() =>
      this.coordinator.enqueue(async () => {
        if (this.resourceFailure) throw this.resourceFailure;
        return operation();
      }));
    const latched = result.catch((error: unknown) => {
      if (error instanceof OoxmlResourceLimitError) this.resourceFailure ??= error;
      throw error;
    });
    this.operationTail = latched.then(() => undefined, () => undefined);
    return latched;
  }

  /** Close every shared host before replacing the package archive generation. */
  async reset(): Promise<void> {
    for (const pending of this.pendingOpens.values()) pending.canceled = true;
    let requestId = 1;
    for (const { host, identity } of [...this.sessions.values()]) {
      await host.dispatch({
        protocol: PULL_SESSION_PROTOCOL,
        kind: 'close',
        ...identity,
        requestId: requestId++,
      }, () => undefined);
    }
    this.sessions.clear();
    await this.operationTail;
    this.pendingOpens.clear();
    this.resourceFailure = undefined;
  }

  private requireArchive(): WorksheetCursorArchive {
    const archive = this.archive();
    if (!archive) throw new Error('Workbook not loaded');
    return archive;
  }

  private async closeIdentity(identity: PullSessionIdentity<number>): Promise<void> {
    const session = this.sessions.get(identity.sessionId);
    if (session) {
      await session.host.dispatch({
        protocol: PULL_SESSION_PROTOCOL,
        kind: 'close',
        ...identity,
        requestId: 1,
      }, () => undefined);
      return;
    }
    const pending = this.pendingOpens.get(identity.sessionId);
    if (
      pending &&
      pending.identity.operationId === identity.operationId &&
      pending.identity.generation === identity.generation
    ) {
      pending.canceled = true;
    }
  }

  // Every opened worksheet cursor owns a PackageOperation ledger (the archive
  // admits only OPC packages), so any checkpoint failure is a real error.
  private readResourceUsage(): OoxmlResourceUsageSnapshot {
    return decodeOoxmlResourceUsage(
      this.executeArchive((archive) => archive.sheet_cursor_resource_usage()),
    );
  }
}

export function isWorksheetPullCommand(value: unknown): value is PullSessionCommand<number> {
  return !!value && typeof value === 'object' &&
    (value as { protocol?: unknown }).protocol === PULL_SESSION_PROTOCOL;
}

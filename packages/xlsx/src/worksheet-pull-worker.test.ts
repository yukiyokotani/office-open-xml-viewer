import { describe, expect, it, vi } from 'vitest';
import { PULL_SESSION_PROTOCOL, type PullSessionCommand, type PullSessionResponse } from '@silurus/ooxml-core/worker';
import { WorksheetPullWorker } from './worksheet-pull-worker.js';
import { OoxmlResourceLimitError } from '@silurus/ooxml-core';
import type { Worksheet } from './types.js';

const identity = { sessionId: 4, operationId: 9, generation: 2 } as const;
const usageBytes = new TextEncoder().encode(JSON.stringify({
  archiveEntryCount: 1,
  declaredInflatedBytes: 2,
  distinctInflatedBytes: 3,
  operationInflatedBytes: 4,
}));

async function openWorker(
  worker: WorksheetPullWorker,
  value: { sessionId: number; operationId: number; generation: number } = identity,
): Promise<void> {
  worker.reserveOpen(value);
  await worker.open(0, 'Sheet1', value);
}

function command(
  requestId: number,
  body:
    | { kind: 'pull'; sequence: number; byteCredit: number }
    | { kind: 'ack'; sequence: number }
    | { kind: 'cancel'; reason: 'request-error' },
): PullSessionCommand<number> {
  return { protocol: PULL_SESSION_PROTOCOL, requestId, ...identity, ...body };
}

describe('WorksheetPullWorker', () => {
  it('latches an ordinary worker-side renderer violation for sibling operations', async () => {
    const fatal = new OoxmlResourceLimitError('renderer index limit', {
      stage: 'layout',
      violation: {
        format: 'xlsx', operation: 'render-viewport', resource: 'renderer-index', metric: 'entries',
        limit: 250_000, observed: 250_001, configurable: false,
        usage: { archiveEntryCount: 1, declaredInflatedBytes: 2, distinctInflatedBytes: 3, operationInflatedBytes: 4 },
      },
    });
    const worker = new WorksheetPullWorker(() => null);
    await expect(worker.run(() => { throw fatal; })).rejects.toBe(fatal);
    const sibling = vi.fn();
    await expect(worker.run(sibling)).rejects.toBe(fatal);
    expect(sibling).not.toHaveBeenCalled();
  });

  it('measures exact transfer bytes and commits Rust only after terminal ACK', async () => {
    const terminal: Worksheet = {
      name: 'Sheet1', rows: [], colWidths: {}, rowHeights: {}, defaultColWidth: 8.43,
      defaultRowHeight: 15, mergeCells: [], freezeRows: 0, freezeCols: 0,
      conditionalFormats: [], images: [], charts: [],
    };
    const payloads = [
      new TextEncoder().encode(JSON.stringify({ kind: 'rows', rows: [{ index: 1, height: null, cells: [] }] })),
      new TextEncoder().encode(JSON.stringify({ kind: 'finished', worksheet: terminal })),
    ];
    let pullIndex = 0;
    const acknowledge = vi.fn();
    const archive = {
      open_sheet_cursor: vi.fn(),
      pull_sheet_cursor: vi.fn(() => payloads[pullIndex++]),
      sheet_cursor_pull_finished: vi.fn(() => pullIndex === payloads.length),
      sheet_cursor_resource_usage: vi.fn(() => new TextEncoder().encode(JSON.stringify({
        archiveEntryCount: 1, declaredInflatedBytes: 2, distinctInflatedBytes: 3, operationInflatedBytes: 4,
      }))),
      acknowledge_sheet_cursor_terminal: acknowledge,
      cancel_sheet_cursor: vi.fn(),
      close_sheet_cursor: vi.fn(),
    };
    const accepted = vi.fn();
    const worker = new WorksheetPullWorker(() => archive, accepted);
    const replies: PullSessionResponse<ArrayBuffer, number>[] = [];
    const post = (response: PullSessionResponse<ArrayBuffer, number>) => replies.push(response);

    await openWorker(worker);
    await worker.dispatch(command(1, { kind: 'pull', sequence: 0, byteCredit: 64 * 1024 * 1024 }), post);
    const rows = replies.at(-1);
    expect(rows).toMatchObject({ kind: 'chunk', done: false, byteLength: payloads[0].byteLength });
    await worker.dispatch(command(2, { kind: 'ack', sequence: 0 }), post);
    await worker.dispatch(command(3, { kind: 'pull', sequence: 1, byteCredit: 64 * 1024 * 1024 }), post);
    const finished = replies.at(-1);
    expect(finished).toMatchObject({ kind: 'chunk', done: true, byteLength: payloads[1].byteLength });
    expect(acknowledge).not.toHaveBeenCalled();
    expect(accepted).not.toHaveBeenCalled();
    const ordinaryOperation = vi.fn();
    const queued = worker.run(ordinaryOperation);
    await Promise.resolve();
    expect(ordinaryOperation).not.toHaveBeenCalled();

    await worker.dispatch(command(4, { kind: 'ack', sequence: 1 }), post);
    await queued;
    expect(acknowledge).toHaveBeenCalledOnce();
    expect(accepted).toHaveBeenCalledWith(
      0,
      expect.objectContaining({ rows: [expect.objectContaining({ index: 1 })] }),
      expect.objectContaining({ rows: 1, cells: 0, jsonBytes: expect.any(Number) }),
      expect.objectContaining({ operationInflatedBytes: 4 }),
    );
    expect(ordinaryOperation).toHaveBeenCalledOnce();
  });

  it('drops provisional rows and their usage when the terminal model has a parse error', async () => {
    const payloads = [
      new TextEncoder().encode(JSON.stringify({
        kind: 'rows',
        rows: [{ index: 1, height: null, cells: [] }],
      })),
      new TextEncoder().encode(JSON.stringify({
        kind: 'finished',
        worksheet: {
          name: 'Sheet1', rows: [], colWidths: {}, rowHeights: {}, defaultColWidth: 8.43,
          defaultRowHeight: 15, mergeCells: [], freezeRows: 0, freezeCols: 0,
          conditionalFormats: [], images: [], charts: [], parseError: 'malformed tail',
        },
      })),
    ];
    let pullIndex = 0;
    const archive = {
      open_sheet_cursor: vi.fn(),
      pull_sheet_cursor: vi.fn(() => payloads[pullIndex++]),
      sheet_cursor_pull_finished: vi.fn(() => pullIndex === payloads.length),
      sheet_cursor_resource_usage: vi.fn(() => usageBytes),
      acknowledge_sheet_cursor_terminal: vi.fn(),
      cancel_sheet_cursor: vi.fn(),
      close_sheet_cursor: vi.fn(),
    };
    const accepted = vi.fn();
    const worker = new WorksheetPullWorker(() => archive, accepted);
    const post = () => undefined;

    await openWorker(worker);
    await worker.dispatch(command(1, { kind: 'pull', sequence: 0, byteCredit: 64 * 1024 * 1024 }), post);
    await worker.dispatch(command(2, { kind: 'ack', sequence: 0 }), post);
    await worker.dispatch(command(3, { kind: 'pull', sequence: 1, byteCredit: 64 * 1024 * 1024 }), post);
    await worker.dispatch(command(4, { kind: 'ack', sequence: 1 }), post);

    expect(accepted).toHaveBeenCalledWith(
      0,
      expect.objectContaining({ rows: [], parseError: 'malformed tail' }),
      expect.objectContaining({ rows: 0, cells: 0, ownedUtf8Bytes: 0 }),
      expect.any(Object),
    );
  });

  it('does not decode or retain row payloads in the slim transfer-only worker', async () => {
    const malformedRows = new TextEncoder().encode('not decoded by the slim worker');
    const archive = {
      open_sheet_cursor: vi.fn(),
      pull_sheet_cursor: vi.fn(() => malformedRows),
      sheet_cursor_pull_finished: vi.fn(() => false),
      sheet_cursor_resource_usage: vi.fn(() => usageBytes),
      acknowledge_sheet_cursor_terminal: vi.fn(),
      cancel_sheet_cursor: vi.fn(),
      close_sheet_cursor: vi.fn(),
    };
    const worker = new WorksheetPullWorker(() => archive);
    const replies: PullSessionResponse<ArrayBuffer, number>[] = [];
    await openWorker(worker);
    await worker.dispatch(
      command(1, { kind: 'pull', sequence: 0, byteCredit: 64 * 1024 * 1024 }),
      (response) => replies.push(response),
    );

    expect(replies[0]).toMatchObject({ kind: 'chunk', byteLength: malformedRows.byteLength });
    await worker.dispatch(command(2, { kind: 'cancel', reason: 'request-error' }), () => undefined);
    expect(archive.cancel_sheet_cursor).toHaveBeenCalledOnce();
  });

  it('keeps the Rust terminal provisional when render-worker acceptance fails', async () => {
    const terminal = new TextEncoder().encode(JSON.stringify({
      kind: 'finished',
      worksheet: {
        name: 'Sheet1', rows: [], colWidths: {}, rowHeights: {}, defaultColWidth: 8.43,
        defaultRowHeight: 15, mergeCells: [], freezeRows: 0, freezeCols: 0,
        conditionalFormats: [], images: [], charts: [],
      },
    }));
    const archive = {
      open_sheet_cursor: vi.fn(),
      pull_sheet_cursor: vi.fn(() => terminal),
      sheet_cursor_pull_finished: vi.fn(() => true),
      sheet_cursor_resource_usage: vi.fn(() => usageBytes),
      acknowledge_sheet_cursor_terminal: vi.fn(),
      cancel_sheet_cursor: vi.fn(),
      close_sheet_cursor: vi.fn(),
    };
    const worker = new WorksheetPullWorker(() => archive, () => {
      throw new Error('cache rejected terminal');
    });
    const replies: PullSessionResponse<ArrayBuffer, number>[] = [];
    await openWorker(worker);
    await worker.dispatch(
      command(1, { kind: 'pull', sequence: 0, byteCredit: 64 * 1024 * 1024 }),
      (response) => replies.push(response),
    );
    await worker.dispatch(command(2, { kind: 'ack', sequence: 0 }), (response) => replies.push(response));

    expect(replies.at(-1)).toMatchObject({ kind: 'error' });
    expect(archive.acknowledge_sheet_cursor_terminal).not.toHaveBeenCalled();
    await worker.dispatch(command(3, { kind: 'cancel', reason: 'request-error' }), () => undefined);
    expect(archive.cancel_sheet_cursor).toHaveBeenCalledOnce();
  });

  it('rejects a pull whose usage checkpoint fails', async () => {
    const terminal = new TextEncoder().encode(JSON.stringify({ kind: 'finished', worksheet: {
      name: 'Sheet1', rows: [], colWidths: {}, rowHeights: {}, defaultColWidth: 8.43,
      defaultRowHeight: 15, mergeCells: [], freezeRows: 0, freezeCols: 0,
      conditionalFormats: [], images: [], charts: [],
    } }));
    const makeArchive = (usageError: Error) => ({
      open_sheet_cursor: vi.fn(),
      pull_sheet_cursor: vi.fn(() => terminal),
      sheet_cursor_pull_finished: vi.fn(() => true),
      sheet_cursor_resource_usage: vi.fn(() => { throw usageError; }),
      acknowledge_sheet_cursor_terminal: vi.fn(),
      cancel_sheet_cursor: vi.fn(),
      close_sheet_cursor: vi.fn(),
    });

    const violation = makeArchive(new Error('OOXML_RESOURCE_LIMIT: usage checkpoint failed'));
    const rejected = new WorksheetPullWorker(() => violation);
    const rejectedReplies: PullSessionResponse<ArrayBuffer, number>[] = [];
    await openWorker(rejected, { ...identity, sessionId: 5 });
    await rejected.dispatch(
      { ...command(2, { kind: 'pull', sequence: 0, byteCredit: 64 * 1024 * 1024 }), sessionId: 5 },
      (response) => rejectedReplies.push(response),
    );
    expect(rejectedReplies[0]).toMatchObject({ kind: 'error' });
  });

  it('closes the shared lifecycle before a reparse generation proceeds', async () => {
    const archive = {
      open_sheet_cursor: vi.fn(),
      pull_sheet_cursor: vi.fn(),
      sheet_cursor_pull_finished: vi.fn(),
      sheet_cursor_resource_usage: vi.fn(),
      acknowledge_sheet_cursor_terminal: vi.fn(),
      cancel_sheet_cursor: vi.fn(),
      close_sheet_cursor: vi.fn(),
    };
    const worker = new WorksheetPullWorker(() => archive);
    await openWorker(worker);
    await worker.reset();
    await worker.run(() => undefined);

    expect(archive.close_sheet_cursor).toHaveBeenCalledOnce();
    await openWorker(worker, { ...identity, sessionId: 6, generation: 3 });
    expect(archive.open_sheet_cursor).toHaveBeenCalledTimes(2);
    await worker.reset();
  });

  it('converges cleanup without an unhandled rejection when posting throws', async () => {
    const rows = new TextEncoder().encode(JSON.stringify({ kind: 'rows', rows: [] }));
    const archive = {
      open_sheet_cursor: vi.fn(),
      pull_sheet_cursor: vi.fn(() => rows),
      sheet_cursor_pull_finished: vi.fn(() => false),
      sheet_cursor_resource_usage: vi.fn(() => usageBytes),
      acknowledge_sheet_cursor_terminal: vi.fn(),
      cancel_sheet_cursor: vi.fn(),
      close_sheet_cursor: vi.fn(),
    };
    const worker = new WorksheetPullWorker(() => archive);
    const posted: PullSessionResponse<ArrayBuffer, number>[] = [];
    await openWorker(worker);
    await expect(worker.dispatchSafely(
      command(1, { kind: 'pull', sequence: 0, byteCredit: 64 * 1024 * 1024 }),
      (response) => {
        if (response.kind === 'chunk') throw new Error('structured clone failed');
        posted.push(response);
      },
    )).resolves.toBeUndefined();
    expect(posted[0]).toMatchObject({ kind: 'error', requestId: 1 });
    expect(archive.cancel_sheet_cursor).toHaveBeenCalledOnce();
    const ordinary = vi.fn();
    await worker.run(ordinary);
    expect(ordinary).toHaveBeenCalledOnce();
  });

  it('rolls back a prepared render cache when Rust terminal ACK fails', async () => {
    const terminal = new TextEncoder().encode(JSON.stringify({ kind: 'finished', worksheet: {
      name: 'Sheet1', rows: [], colWidths: {}, rowHeights: {}, defaultColWidth: 8.43,
      defaultRowHeight: 15, mergeCells: [], freezeRows: 0, freezeCols: 0,
      conditionalFormats: [], images: [], charts: [],
    } }));
    const archive = {
      open_sheet_cursor: vi.fn(),
      pull_sheet_cursor: vi.fn(() => terminal),
      sheet_cursor_pull_finished: vi.fn(() => true),
      sheet_cursor_resource_usage: vi.fn(() => usageBytes),
      acknowledge_sheet_cursor_terminal: vi.fn(() => { throw new Error('ack trap'); }),
      cancel_sheet_cursor: vi.fn(),
      close_sheet_cursor: vi.fn(),
    };
    let cached: Worksheet | undefined;
    const worker = new WorksheetPullWorker(() => archive, (_index, worksheet) => {
      cached = worksheet;
      return () => { cached = undefined; };
    });
    const replies: PullSessionResponse<ArrayBuffer, number>[] = [];
    await openWorker(worker);
    await worker.dispatch(
      command(1, { kind: 'pull', sequence: 0, byteCredit: 64 * 1024 * 1024 }),
      (response) => replies.push(response),
    );
    await worker.dispatch(command(2, { kind: 'ack', sequence: 0 }), (response) => replies.push(response));
    expect(replies.at(-1)).toMatchObject({ kind: 'error' });
    expect(cached).toBeUndefined();
    await worker.dispatch(command(3, { kind: 'cancel', reason: 'request-error' }), () => undefined);
    await expect(worker.run(() => 'after rollback')).resolves.toBe('after rollback');
  });

  it('honors cancel received while an open is delayed before registration', async () => {
    const archive = {
      open_sheet_cursor: vi.fn(),
      pull_sheet_cursor: vi.fn(),
      sheet_cursor_pull_finished: vi.fn(),
      sheet_cursor_resource_usage: vi.fn(),
      acknowledge_sheet_cursor_terminal: vi.fn(),
      cancel_sheet_cursor: vi.fn(),
      close_sheet_cursor: vi.fn(),
    };
    const worker = new WorksheetPullWorker(() => archive);
    let release!: () => void;
    const blocker = worker.run(() => new Promise<void>((resolve) => { release = resolve; }));
    await Promise.resolve();
    worker.reserveOpen(identity);
    const opening = worker.open(0, 'Sheet1', identity);
    const lifecycle: PullSessionResponse<ArrayBuffer, number>[] = [];
    await worker.dispatch(
      command(1, { kind: 'cancel', reason: 'request-error' }),
      (response) => lifecycle.push(response),
    );
    expect(lifecycle[0]).toMatchObject({ kind: 'accepted', command: 'cancel' });
    release();
    await blocker;
    await expect(opening).rejects.toThrow('open was canceled');
    expect(archive.open_sheet_cursor).not.toHaveBeenCalled();
    const ordinary = vi.fn();
    await worker.run(ordinary);
    expect(ordinary).toHaveBeenCalledOnce();
  });

  it('cleans up an opened cursor when opened and fallback error posts both throw', async () => {
    const archive = {
      open_sheet_cursor: vi.fn(),
      pull_sheet_cursor: vi.fn(),
      sheet_cursor_pull_finished: vi.fn(),
      sheet_cursor_resource_usage: vi.fn(),
      acknowledge_sheet_cursor_terminal: vi.fn(),
      cancel_sheet_cursor: vi.fn(),
      close_sheet_cursor: vi.fn(),
    };
    const worker = new WorksheetPullWorker(() => archive);
    await openWorker(worker);
    await expect(worker.postOpenedSafely(
      identity,
      () => { throw new Error('opened post failed'); },
      () => { throw new Error('fallback post failed'); },
    )).resolves.toBeUndefined();
    expect(archive.close_sheet_cursor).toHaveBeenCalledOnce();
    await expect(worker.run(() => 'ordinary')).resolves.toBe('ordinary');
  });

  it('rejects a late old-generation open after reset without retaining tombstones', async () => {
    const archive = {
      open_sheet_cursor: vi.fn(),
      pull_sheet_cursor: vi.fn(),
      sheet_cursor_pull_finished: vi.fn(),
      sheet_cursor_resource_usage: vi.fn(),
      acknowledge_sheet_cursor_terminal: vi.fn(),
      cancel_sheet_cursor: vi.fn(),
      close_sheet_cursor: vi.fn(),
    };
    const worker = new WorksheetPullWorker(() => archive);
    worker.reserveOpen(identity);
    await worker.reset();
    expect(worker.pendingOpenCount).toBe(0);
    await expect(worker.open(0, 'Sheet1', identity)).rejects.toThrow('reservation is stale or missing');
    expect(archive.open_sheet_cursor).not.toHaveBeenCalled();
    await expect(worker.run(() => 'after stale open')).resolves.toBe('after stale open');

    const fresh = { ...identity, sessionId: 7, generation: 3 };
    await openWorker(worker, fresh);
    expect(worker.pendingOpenCount).toBe(0);
    expect(archive.open_sheet_cursor).toHaveBeenCalledOnce();
    await worker.reset();
    expect(worker.pendingOpenCount).toBe(0);
  });
});

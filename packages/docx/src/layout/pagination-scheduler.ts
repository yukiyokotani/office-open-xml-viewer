/**
 * Driving a suspendable pagination across event-loop turns.
 *
 * `paginateBodySteps` and friends are generators that suspend between body
 * entries (see `PaginationSteps` in `body-paginator.ts`). Synchronously drained
 * they behave exactly as pagination always has. Drained through a scheduler,
 * the same computation releases the thread periodically, which is what lets a
 * large document lay out without freezing the main thread — and what lets a
 * consumer observe pages as they are committed.
 *
 * The scheduler is injected rather than hard-coded so tests can drive an
 * adversarial policy (yield on every single entry) and a deterministic clock,
 * and so a worker can use a different budget from the main thread.
 */

/** Wall-clock source, injectable so tests are deterministic. */
export type PaginationClock = () => number;

export interface PaginationSchedulerOptions {
  /**
   * Milliseconds of uninterrupted work before releasing the thread. The default
   * is one 60fps frame's worth of budget: long enough that the per-slice
   * overhead stays negligible against a multi-second layout, short enough that
   * input handled between slices still feels responsive.
   */
  readonly sliceMs?: number;
  /** Defaults to `performance.now`, falling back to `Date.now`. */
  readonly now?: PaginationClock;
  /** Releases the thread. Defaults to a macrotask so rendering can interleave. */
  readonly yieldToHost?: () => Promise<void>;
  /** Aborts the drain at the next suspension point. */
  readonly signal?: AbortSignal;
  /**
   * Called at each suspension point with the number of pages committed so far.
   * Monotonic within a pass, but NOT across passes: a convergence pass restarts
   * pagination from page zero, so this can go down. Consumers that surface it
   * must treat a decrease as "this pass is re-deriving what you already saw".
   */
  readonly onProgress?: (committedPages: number) => void;
}

const DEFAULT_SLICE_MS = 16;

function defaultClock(): number {
  const performanceNow = (globalThis as { performance?: { now?: () => number } })
    .performance?.now;
  return performanceNow ? performanceNow.call(globalThis.performance) : Date.now();
}

/**
 * Release the thread so the host can paint and handle input.
 *
 * `MessageChannel` is used rather than `setTimeout(0)` because timers are
 * clamped (4ms after a few nested levels) and would dominate the slice budget;
 * a message task is dispatched on the next turn without clamping. Workers have
 * `MessageChannel` too, so both threads take this path.
 */
function defaultYieldToHost(): Promise<void> {
  const channelConstructor = (globalThis as { MessageChannel?: typeof MessageChannel })
    .MessageChannel;
  if (!channelConstructor) return new Promise((resolve) => { setTimeout(resolve, 0); });
  return new Promise((resolve) => {
    const channel = new channelConstructor();
    channel.port1.onmessage = () => {
      channel.port1.close();
      channel.port2.close();
      resolve();
    };
    channel.port2.postMessage(null);
  });
}

/** Thrown when a drain is cancelled through its {@link AbortSignal}. */
export class PaginationAbortError extends Error {
  constructor() {
    super('Pagination was aborted');
    this.name = 'PaginationAbortError';
  }
}

/**
 * Drive a pagination generator to completion, releasing the thread whenever the
 * slice budget is spent.
 *
 * The generator's own suspension points decide WHERE it is safe to stop; this
 * decides WHETHER to stop there. Because the two are separate, the layout
 * produced here is identical to the synchronous one for any scheduling policy —
 * including yielding at every opportunity, which the equivalence tests use.
 */
export async function drainPaginationAsync<T>(
  steps: Generator<number, T, void>,
  options: PaginationSchedulerOptions = {},
): Promise<T> {
  const sliceMs = options.sliceMs ?? DEFAULT_SLICE_MS;
  const now = options.now ?? defaultClock;
  const yieldToHost = options.yieldToHost ?? defaultYieldToHost;
  const { signal, onProgress } = options;
  let sliceStart = now();
  let step = steps.next();
  while (!step.done) {
    onProgress?.(step.value);
    if (signal?.aborted) {
      // Let the generator run its `finally` blocks before abandoning it.
      steps.return(undefined as never);
      throw new PaginationAbortError();
    }
    if (now() - sliceStart >= sliceMs) {
      await yieldToHost();
      sliceStart = now();
    }
    step = steps.next();
  }
  return step.value;
}

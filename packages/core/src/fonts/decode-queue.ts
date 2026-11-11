/**
 * A bounded queue of shared, holder-tracked jobs (the metric-table decodes
 * of `source-session.ts`). Library resource policy, not OOXML behaviour.
 *
 * At most `limit` jobs run at a time. One job per key is shared by every
 * holder that asks for it. A queued job whose last holder leaves (released,
 * or its wait aborted) is dropped before it runs, so a released document's
 * work never runs ahead of, or delays, a live one's. A running job cannot be
 * interrupted; it finishes in its slot and its result stays with the key.
 */
import { createAbortError } from '../internal/abort.js';

interface Job {
  readonly key: object;
  readonly work: () => Promise<void>;
  readonly holders: Set<object>;
  started: boolean;
  readonly done: Promise<void>;
  settle: () => void;
}

export class DecodeQueue {
  private active = 0;
  private readonly queue: Job[] = [];
  private readonly jobs = new Map<object, Job>();

  constructor(private readonly limit: number) {}

  /**
   * Run `work` for `key` (once, shared with every other holder of `key`) and
   * wait for it on behalf of `holder`. Rejects with an `AbortError` as soon
   * as `signal` aborts; the holder then leaves the job.
   */
  run(key: object, holder: object, work: () => Promise<void>, signal?: AbortSignal): Promise<void> {
    if (signal?.aborted) return Promise.reject(createAbortError('Decode wait aborted'));
    let job = this.jobs.get(key);
    if (!job) {
      let settle!: () => void;
      const done = new Promise<void>((resolve) => { settle = resolve; });
      job = { key, work, holders: new Set(), started: false, done, settle };
      this.jobs.set(key, job);
      this.queue.push(job);
    }
    job.holders.add(holder);
    this.pump();
    if (!signal) return job.done;
    const joined = job;
    return new Promise<void>((resolve, reject) => {
      const onAbort = () => {
        this.leave(joined, holder);
        reject(createAbortError('Decode wait aborted'));
      };
      signal.addEventListener('abort', onAbort, { once: true });
      joined.done.then(() => { signal.removeEventListener('abort', onAbort); resolve(); });
    });
  }

  /** `holder` leaves every job it waits on. */
  release(holder: object): void {
    for (const job of [...this.jobs.values()]) this.leave(job, holder);
  }

  /** Jobs queued or running (tests, leak checks). */
  get size(): number { return this.jobs.size; }

  private leave(job: Job, holder: object): void {
    job.holders.delete(holder);
    if (job.started || job.holders.size > 0) return;
    const at = this.queue.indexOf(job);
    if (at >= 0) this.queue.splice(at, 1);
    if (this.jobs.get(job.key) === job) this.jobs.delete(job.key);
    job.settle();
  }

  private pump(): void {
    while (this.active < this.limit && this.queue.length > 0) {
      const job = this.queue.shift()!;
      job.started = true;
      this.active++;
      void job.work().catch(() => undefined).then(() => {
        this.active--;
        if (this.jobs.get(job.key) === job) this.jobs.delete(job.key);
        job.settle();
        this.pump();
      });
    }
  }
}

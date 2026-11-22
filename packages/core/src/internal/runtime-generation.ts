type RuntimeState = 'uninitialized' | 'ready' | 'poisoned';

/**
 * Environment-neutral realm initialization, generation, and poison fan-out.
 *
 * Each format's wasm-bindgen glue keeps one module instance per JavaScript
 * realm (DOCX, XLSX and PPTX each have their own generation), so sessions
 * never reinitialize WASM themselves. A trap poisons the whole generation:
 * every live session registered through `onPoison` fails, and the next
 * `ensureReady()` performs the single reinitialization. Sessions from the old
 * generation stay failed; only sessions opened afterwards use the new one.
 */
export class RuntimeGeneration<TFailure extends Error> {
  private state: RuntimeState = 'uninitialized';
  private generationValue = 0;
  private readiness: Promise<void> | undefined;
  private readonly poisonListeners = new Set<(error: TFailure) => void>();

  constructor(
    private readonly initialize: () => Promise<unknown> | unknown,
    private readonly reinitialize: () => Promise<unknown> | unknown,
    private readonly normalizeFailure: (error: unknown) => TFailure | null,
  ) {}

  get generation(): number { return this.generationValue; }
  get poisoned(): boolean { return this.state === 'poisoned'; }

  onPoison(listener: (error: TFailure) => void): () => void {
    this.poisonListeners.add(listener);
    return () => this.poisonListeners.delete(listener);
  }

  async ensureReady(): Promise<void> {
    if (this.state === 'ready') return;
    if (!this.readiness) {
      const operation = this.state === 'uninitialized' ? this.initialize : this.reinitialize;
      this.readiness = Promise.resolve().then(operation).then(() => {
        this.generationValue += 1;
        this.state = 'ready';
        this.readiness = undefined;
      }, (error: unknown) => {
        this.readiness = undefined;
        throw error;
      });
    }
    await this.readiness;
  }

  run<TResult>(operation: () => TResult): TResult {
    try {
      return operation();
    } catch (error) {
      const failure = this.normalizeFailure(error);
      if (!failure) throw error;
      this.poison(failure);
      throw failure;
    }
  }

  /**
   * Run an operation only while the currently observed generation is ready.
   * JavaScript run-to-completion makes the readiness check, generation capture,
   * and synchronous operation one acquisition step relative to sibling promise
   * continuations. A trap thrown by `operation` still propagates through `run`
   * and is never reported as a retryable stale-generation result.
   */
  tryRunReady<TResult>(operation: () => TResult):
    | { readonly current: false }
    | { readonly current: true; readonly generation: number; readonly value: TResult } {
    if (this.state !== 'ready') return { current: false };
    const generation = this.generationValue;
    const value = this.run(operation);
    if (this.state !== 'ready' || generation !== this.generationValue) {
      return { current: false };
    }
    return { current: true, generation, value };
  }

  poison(error: TFailure): void {
    this.state = 'poisoned';
    this.readiness = undefined;
    for (const listener of this.poisonListeners) listener(error);
  }

  assertCurrent(generation: number): void {
    if (this.state !== 'ready' || generation !== this.generationValue) {
      throw new Error('WASM archive session belongs to a discarded runtime generation');
    }
  }
}

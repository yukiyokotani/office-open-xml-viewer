/** Isolates application observer failures from the authoritative layout result. */
export class ProgressiveLayoutObserverNotifier {
  /** One callback may intentionally serve more than one lifecycle slot. A
   * failure disables only that registration, not the same function registered
   * under another observer name. */
  private readonly failed = new WeakMap<Function, Set<string>>();

  notify<Args extends readonly unknown[]>(
    name: string,
    callback: ((...args: Args) => unknown) | undefined,
    ...args: Args
  ): void {
    if (!callback || this.failed.get(callback)?.has(name)) return;
    try {
      const result = callback(...args);
      if (result && typeof (result as PromiseLike<unknown>).then === 'function') {
        Promise.resolve(result).catch((error) => this.reportOnce(name, callback, error));
      }
    } catch (error) {
      this.reportOnce(name, callback, error);
    }
  }

  private reportOnce(name: string, callback: Function, cause: unknown): void {
    const registrations = this.failed.get(callback) ?? new Set<string>();
    if (registrations.has(name)) return;
    registrations.add(name);
    if (!this.failed.has(callback)) this.failed.set(callback, registrations);
    const error = cause instanceof Error ? cause : new Error(String(cause));
    console.error(`[ooxml] ${name} callback failed and was disabled:`, error);
  }
}

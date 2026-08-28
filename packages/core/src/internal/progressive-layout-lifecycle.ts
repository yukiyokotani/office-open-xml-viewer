/** Internal success/failure state shared by progressive document engines. */
type ProgressiveLayoutState =
  | { readonly status: 'complete' }
  | { readonly status: 'pending' }
  | { readonly status: 'failed'; readonly error: Error };

/**
 * Keeps "all content is ready" separate from "background work has stopped".
 * A failed progressive load is settled, but it is never complete.
 */
export class ProgressiveLayoutLifecycle {
  private state: ProgressiveLayoutState = Object.freeze({ status: 'complete' });

  get complete(): boolean {
    return this.state.status === 'complete';
  }

  get settled(): boolean {
    return this.state.status !== 'pending';
  }

  begin(): void {
    this.state = Object.freeze({ status: 'pending' });
  }

  succeed(): void {
    this.state = Object.freeze({ status: 'complete' });
  }

  fail(cause: unknown): Error {
    const error = cause instanceof Error ? cause : new Error(String(cause));
    this.state = Object.freeze({ status: 'failed', error });
    return error;
  }

  throwIfFailed(): void {
    if (this.state.status === 'failed') throw this.state.error;
  }
}

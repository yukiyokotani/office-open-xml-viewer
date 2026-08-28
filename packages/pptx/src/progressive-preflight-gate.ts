/** Worker-side acknowledgement gate for progressive PPTX preflight.
 *
 * The worker registers a checkpoint before publishing a slide prefix. The host
 * acknowledges it only after crossing its own task boundary, which gives load
 * continuations time to enqueue an opening-slide render. Messages sent by one
 * host to one Worker are ordered, so that render request is handled before the
 * acknowledgement lets the next slide's preflight begin.
 */
export class ProgressivePreflightGate {
  private pending: {
    readonly parseId: number;
    readonly availableSlides: number;
    readonly resolve: () => void;
  } | null = null;

  wait(parseId: number, availableSlides: number): Promise<void> {
    this.reset();
    return new Promise<void>((resolve) => {
      this.pending = { parseId, availableSlides, resolve };
    });
  }

  continue(parseId: number, availableSlides: number): boolean {
    const pending = this.pending;
    if (!pending || pending.parseId !== parseId || pending.availableSlides !== availableSlides) {
      return false;
    }
    this.pending = null;
    pending.resolve();
    return true;
  }

  reset(): void {
    const pending = this.pending;
    this.pending = null;
    pending?.resolve();
  }
}

export interface AutoResizeOptions {
  /**
   * Skip rendering while `document.hidden` is true and fire once with the latest
   * observed size when the tab becomes visible again. Default: true.
   */
  pauseWhenHidden?: boolean;
}

/**
 * Observe an element's size and invoke a render callback, coalescing bursts to
 * one call per animation frame and serializing overlapping async renders.
 *
 * Framework-agnostic: call from any mount/setup hook and invoke the returned
 * disposer in the corresponding teardown hook.
 *
 * @example
 * const detach = autoResize(
 *   (width) => pres.renderSlide(canvas, 0, { width }),
 *   canvas,
 * );
 * // later
 * detach();
 */
export function autoResize(
  render: (width: number, height: number) => void | Promise<void>,
  element: Element,
  opts: AutoResizeOptions = {},
): () => void {
  const pauseWhenHidden = opts.pauseWhenHidden ?? true;

  let rafId: number | null = null;
  let pendingWidth = 0;
  let pendingHeight = 0;
  let renderInFlight: Promise<void> | null = null;
  let rerunAfter = false;
  let disposed = false;

  const schedule = (): void => {
    if (disposed) return;
    if (pauseWhenHidden && typeof document !== 'undefined' && document.hidden) return;
    if (renderInFlight) {
      rerunAfter = true;
      return;
    }
    if (rafId !== null) return;
    rafId = requestAnimationFrame(runFrame);
  };

  const runFrame = async (): Promise<void> => {
    rafId = null;
    if (disposed) return;
    const w = pendingWidth;
    const h = pendingHeight;
    try {
      const result = render(w, h);
      renderInFlight = result instanceof Promise ? result : Promise.resolve();
      await renderInFlight;
    } catch (err) {
      console.error('[autoResize] render failed:', err);
    } finally {
      renderInFlight = null;
      if (rerunAfter && !disposed) {
        rerunAfter = false;
        schedule();
      }
    }
  };

  const ro = new ResizeObserver((entries) => {
    for (const entry of entries) {
      const rect = entry.contentRect;
      pendingWidth = rect.width;
      pendingHeight = rect.height;
    }
    schedule();
  });
  ro.observe(element);

  const onVisibility = (): void => {
    if (typeof document !== 'undefined' && !document.hidden) schedule();
  };
  if (pauseWhenHidden && typeof document !== 'undefined') {
    document.addEventListener('visibilitychange', onVisibility);
  }

  return () => {
    disposed = true;
    ro.disconnect();
    if (rafId !== null) {
      cancelAnimationFrame(rafId);
      rafId = null;
    }
    if (pauseWhenHidden && typeof document !== 'undefined') {
      document.removeEventListener('visibilitychange', onVisibility);
    }
  };
}

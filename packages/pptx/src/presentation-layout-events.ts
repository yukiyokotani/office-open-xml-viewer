/**
 * Internal publication channel for a presentation whose paintable slide prefix
 * grows after `load()` resolves.
 *
 * The final slide count and dimensions are known from the PPTX bootstrap, so
 * publications report availability separately from the stable total. Viewers
 * subscribe to the presentation they currently own; this keeps reload and
 * borrowed-presentation lifetimes aligned with `TerminalResourceOwner`.
 */

export interface PptxLayoutPublication {
  readonly availableSlides: number;
  readonly slideCount: number;
  readonly exact: boolean;
  readonly complete: boolean;
  readonly error?: unknown;
}

type Listener = (publication: PptxLayoutPublication) => void;
interface Subscription {
  readonly notify: Listener;
  readonly report: (error: unknown) => void;
}

const listeners = new WeakMap<object, Set<Subscription>>();

export function publishPptxLayout(
  presentation: object,
  publication: PptxLayoutPublication,
): void {
  const immutable = Object.freeze({ ...publication });
  for (const subscription of [...(listeners.get(presentation) ?? [])]) {
    try {
      subscription.notify(immutable);
    } catch (error) {
      try { subscription.report(error); } catch {}
    }
  }
}

export function subscribePptxLayout(
  presentation: object,
  current: () => PptxLayoutPublication,
  listener: Listener,
  report: (error: unknown) => void,
): () => void {
  let registered = listeners.get(presentation);
  if (!registered) {
    registered = new Set();
    listeners.set(presentation, registered);
  }
  const subscription = Object.freeze({ notify: listener, report });
  registered.add(subscription);
  try {
    listener(Object.freeze({ ...current() }));
  } catch (error) {
    try { report(error); } catch {}
  }
  return () => {
    registered?.delete(subscription);
    if (registered?.size === 0) listeners.delete(presentation);
  };
}

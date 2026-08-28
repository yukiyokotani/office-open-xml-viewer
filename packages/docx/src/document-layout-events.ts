/**
 * Internal publication channel for a document whose retained layout changes.
 *
 * Load callbacks remain part of the public acquisition API, but viewers must
 * follow the document they currently own, not the closure that happened to
 * create it. Keeping that subscription here gives both DOCX viewers one source
 * of truth and makes reload authority follow `TerminalResourceOwner` exactly.
 */

export interface DocxLayoutPublication {
  readonly pageCount: number;
  readonly exact: boolean;
  readonly complete: boolean;
  readonly error?: unknown;
}

type Listener = (publication: DocxLayoutPublication) => void;
interface Subscription {
  readonly notify: Listener;
  readonly report: (error: unknown) => void;
}

const listeners = new WeakMap<object, Set<Subscription>>();

export function publishDocxLayout(
  document: object,
  publication: DocxLayoutPublication,
): void {
  const immutable = Object.freeze({ ...publication });
  for (const subscription of [...(listeners.get(document) ?? [])]) {
    try {
      subscription.notify(immutable);
    } catch (error) {
      try { subscription.report(error); } catch {}
    }
  }
}

export function subscribeDocxLayout(
  document: object,
  current: () => DocxLayoutPublication,
  listener: Listener,
  report: (error: unknown) => void,
): () => void {
  let registered = listeners.get(document);
  if (!registered) {
    registered = new Set();
    listeners.set(document, registered);
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
    if (registered?.size === 0) listeners.delete(document);
  };
}

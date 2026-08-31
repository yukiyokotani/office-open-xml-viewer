import type { DocxDocument } from './document.js';
import { normalizeLayoutOptions } from './layout/options.js';
import { documentLayoutRuntimeOf } from './layout/runtime-state.js';

export interface ActiveDocxLayoutView {
  readonly showTrackedChanges: boolean;
  readonly currentDate: number;
}

export interface DocxLayoutViewPublication {
  readonly view: Readonly<ActiveDocxLayoutView>;
  /** Monotonic per-document installation order. Reentrant listeners may see a
   * newer publication before an outer dispatch resumes, so consumers reject
   * any generation they have already passed. */
  readonly generation: number;
  /** Viewer whose awaited method owns the repaint for this installation. */
  readonly requester?: object;
}

type Listener = (publication: DocxLayoutViewPublication) => void;
interface Subscription {
  readonly notify: Listener;
  readonly report: (error: unknown) => void;
}

const listeners = new WeakMap<object, Set<Subscription>>();
const generations = new WeakMap<object, number>();
/** Non-public request metadata. The symbol is stripped from normalized layout
 * options and never crosses the worker boundary. */
export const docxLayoutViewRequester = Symbol('docxLayoutViewRequester');

/** Internal bridge used when a Viewer borrows an already-loaded document.
 *
 * The selected view belongs to DocxDocument's retained-layout runtime. Keeping
 * this accessor outside the public class surface lets both Viewer factories
 * inherit that authority without exposing layout bookkeeping as public API.
 */
export function activeDocxLayoutViewOf(document: DocxDocument): Readonly<ActiveDocxLayoutView> {
  const runtime = documentLayoutRuntimeOf(document);
  const active = runtime.activeLayoutOptions;
  return {
    showTrackedChanges: active?.showTrackedChanges === true,
    currentDate: active?.currentDateMs ?? runtime.defaultCurrentDateMs,
  };
}

/** Select a view on behalf of one Viewer while retaining the document as the
 * sole authority. The boolean result is false when a newer concurrent worker
 * selection won, so the stale caller must not install its requested local state. */
export async function selectDocxLayoutView(
  document: DocxDocument,
  view: Readonly<{ showTrackedChanges?: boolean; currentDate?: Date | number }>,
  requester: object,
): Promise<boolean> {
  const runtime = documentLayoutRuntimeOf(document);
  const normalized = normalizeLayoutOptions(
    view.currentDate,
    runtime.defaultCurrentDateMs,
    view.showTrackedChanges === true,
  );
  const requested = Object.freeze({
    showTrackedChanges: normalized.showTrackedChanges === true,
    currentDate: normalized.currentDateMs,
  });
  const internalView = {
    showTrackedChanges: requested.showTrackedChanges,
    currentDate: requested.currentDate,
    [docxLayoutViewRequester]: requester,
  };
  await document.setLayoutView(internalView);
  const active = activeDocxLayoutViewOf(document);
  return active.currentDate === requested.currentDate
    && active.showTrackedChanges === requested.showTrackedChanges;
}

/** Publish an installed document-global view to every borrowing Viewer. Called
 * only after geometry and worker metadata have atomically adopted that view. */
export function publishDocxLayoutView(
  document: DocxDocument,
  requester?: object,
): void {
  const view = Object.freeze({ ...activeDocxLayoutViewOf(document) });
  const generation = (generations.get(document) ?? 0) + 1;
  generations.set(document, generation);
  const publication = Object.freeze({ view, generation, requester });
  for (const subscription of [...(listeners.get(document) ?? [])]) {
    try {
      subscription.notify(publication);
    } catch (error) {
      try { subscription.report(error); } catch {}
    }
  }
}

export function subscribeDocxLayoutView(
  document: DocxDocument,
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
  return () => {
    registered?.delete(subscription);
    if (registered?.size === 0) listeners.delete(document);
  };
}

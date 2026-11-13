// Message guards and reply copies of the `fontResolver` worker bridge
// (resolver-bridge.ts), a module of their own so the main-realm host
// (resolver-bridge-host.ts) recognises and answers bridge messages without
// the worker client and its resolution bookkeeping.
import type {
  BridgedFamilyAnswer,
  MainFontAcquireReply,
  MainFontResolveReply,
  WorkerFontResolveMessage,
} from './resolver-bridge.js';
import type { FontFaceDeclaration } from './source.js';

const kindOf = (value: unknown): unknown =>
  value !== null && typeof value === 'object' ? (value as { kind?: unknown }).kind : undefined;

/** Whether `value` is a worker → main resolver-bridge message (main side). */
export function isWorkerFontResolveMessage(value: unknown): value is WorkerFontResolveMessage {
  const kind = kindOf(value);
  return kind === 'ooxmlResolveFonts' || kind === 'ooxmlAcquireFonts' || kind === 'ooxmlCancelResolveFonts';
}

/** Whether `value` is a main → worker resolver-bridge reply (worker side). */
export function isMainFontResolveReply(value: unknown): value is MainFontResolveReply | MainFontAcquireReply {
  const kind = kindOf(value);
  return kind === 'ooxmlFontsResolved' || kind === 'ooxmlFontsAcquired';
}

/** The reply's answers with one copy per distinct buffer: the originals are
 * kept by main's cache (and possibly the application), so none is detached. */
export function copiedForTransfer(
  answers: readonly BridgedFamilyAnswer[],
): { answers: readonly BridgedFamilyAnswer[]; transfer: ArrayBuffer[] } {
  const copies = new Map<ArrayBuffer, ArrayBuffer>();
  const wire = (declaration: FontFaceDeclaration): FontFaceDeclaration => {
    const { data } = declaration;
    if (!data) return declaration;
    let copy = copies.get(data);
    if (!copy) {
      copy = data.slice(0);
      copies.set(data, copy);
    }
    return { ...declaration, data: copy };
  };
  const out = answers.map((answer) => (answer.kind === 'claimed'
    ? { ...answer, declarations: answer.declarations.map(wire) } : answer));
  return { answers: out, transfer: [...copies.values()] };
}

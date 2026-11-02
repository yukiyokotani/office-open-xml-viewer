import type { LayoutServices } from './types.js';

type NoteKind = 'footnote' | 'endnote';

/** Internal control transfer, raised only at an established note-capacity
 * failure. Successful pagination never enables recovery. This is content
 * retention policy for #2061, not a claim about Word's oversized pagination
 * (ECMA-376 §17.11.1/.4 describes continuation, not indivisible overflow).
 * Only the failed live page/queue is replayed. Its group's IDs gain
 * permission to use table cuts/indivisible overflow. Ordinary notes keep their existing
 * acquisition, whole-table handling, balance and unequal-column paths.
 */
export class NoteOverflowRecoveryRequired extends Error {
  constructor(readonly kind: NoteKind, readonly ids: readonly string[], readonly failure: Error) {
    super(failure.message);
  }
}

export class NoteOverflowPolicy {
  private readonly permitted = new Set<string>();
  allows(kind: NoteKind, ids: readonly string[]): boolean {
    return ids.some(id => this.permitted.has(`${kind}:${id}`));
  }
  recover(error: unknown): boolean {
    if (!(error instanceof NoteOverflowRecoveryRequired)) return false;
    const added = error.ids.filter(id => !this.permitted.has(`${error.kind}:${id}`));
    if (added.length === 0) throw error.failure;
    added.forEach(id => this.permitted.add(`${error.kind}:${id}`));
    return true;
  }
}

const policies = new WeakMap<LayoutServices, NoteOverflowPolicy>();
export function attachNoteOverflowPolicy(services: LayoutServices, policy: NoteOverflowPolicy): void {
  policies.set(services, policy);
}
export function noteOverflowPolicyOf(services: LayoutServices): NoteOverflowPolicy | undefined {
  return policies.get(services);
}
export function noteOverflowAllowed(services: LayoutServices, kind: NoteKind, ids: readonly string[]): boolean {
  return policies.get(services)?.allows(kind, ids) === true;
}

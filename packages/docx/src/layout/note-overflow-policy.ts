import type { LayoutServices } from './types.js';

type NoteKind = 'footnote' | 'endnote';

export const NOTE_REFRESH_PREFIX_WORK_MULTIPLIER = 4;

/** Resource policy, not an Office fit rule. A destination's geometrically
 * growing paragraph prefixes may spend at most four times its final prefix
 * (see note-paragraph-window.ts). Allocate two acquisition roles (template /
 * destination refresh), two uses (admission / retained placement) and two
 * convergence states (discovery / confirmation): 2 * 2 * 2 * 4 = 32.
 * These eight equivalent source sweeps are finite operational allowances,
 * not a claim that every input needs only two field passes: further passes,
 * reserve trials and replay compete for the same pool. Exhaustion preserves
 * the original rejection. Each note has its own allowance; unrelated source
 * buys no work. */
export const NOTE_RECOVERY_WORK_MULTIPLIER = 8 * NOTE_REFRESH_PREFIX_WORK_MULTIPLIER;

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
  private readonly permitted = new Map<string, Error>();
  private readonly work = new Map<string, number>();
  private originalFailure: Error | undefined;
  allows(kind: NoteKind, ids: readonly string[]): boolean {
    return ids.some(id => this.permitted.has(`${kind}:${id}`));
  }
  /** Recovery is optional: inability to continue must expose the execution's
   * first pre-recovery failure, including its class and admission geometry.
   * A later rejected group was never reachable in v0.89. */
  failure(kind: NoteKind, ids: readonly string[]): Error | undefined {
    for (const id of ids) {
      const failure = this.permitted.get(`${kind}:${id}`);
      if (failure) return this.originalFailure ?? failure;
    }
    return undefined;
  }
  fail(kind: NoteKind, ids: readonly string[], error: unknown): never {
    throw this.failure(kind, ids) ?? error;
  }
  /** Debit before acquisition, in the global acquisition ledger's units.
   * This ledger survives page replay, new concrete sessions and all field
   * passes. A multi-note rejected group cannot let one note borrow another's
   * allowance. Never refund failed trials or reset an existing permission. */
  charge(kind: NoteKind, id: string, sourceUnits: number, units: number): void {
    const key = `${kind}:${id}`;
    const failure = this.permitted.get(key);
    if (!failure) return;
    const spent = this.work.get(key) ?? 0;
    // Even if a later oversized group exhausts its allowance, v0.89 would
    // already have stopped at the execution's first rejected group.
    if (units > NOTE_RECOVERY_WORK_MULTIPLIER * (sourceUnits + 1) - spent) throw this.originalFailure ?? failure;
    this.work.set(key, spent + units);
  }
  recover(error: unknown): boolean {
    if (!(error instanceof NoteOverflowRecoveryRequired)) return false;
    this.originalFailure ??= error.failure;
    const added = error.ids.filter(id => !this.permitted.has(`${error.kind}:${id}`));
    if (added.length === 0) this.fail(error.kind, error.ids, error.failure);
    added.forEach(id => this.permitted.set(`${error.kind}:${id}`, error.failure));
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

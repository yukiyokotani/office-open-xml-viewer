import type { LayoutServices, SourceRef } from './types.js';
import { isDeepFrozenPlainDataRoot } from './plain-data.js';
import type { TextLayoutService } from './text.js';

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
  private readonly sourceSizes = new Map<string, number>();
  private readonly retainedSizes = new WeakMap<object, number>();
  private readonly sourceTraversalSizes = new WeakMap<object, number>();
  private readonly textViews = new WeakMap<TextLayoutService, Map<string, TextLayoutService>>();
  private originalFailure: Error | undefined;
  /** True only after this execution reached an established rejection gate. */
  get recovering(): boolean { return this.originalFailure !== undefined; }
  /** The transaction includes all convergence, composition and final validation.
   * Once the gate captured v0.89's rejection, no recovery-only exception may
   * replace that exact object, even outside acquisition/partition call sites. */
  errorForAttempt(error: unknown): unknown { return this.originalFailure ?? error; }
  /** Keep the service identity stable across destinations so paragraph caches
   * can reuse unchanged shaping. The wrapper changes accounting only. */
  textServiceFor(source: SourceRef, text: TextLayoutService): TextLayoutService {
    let views = this.textViews.get(text);
    if (!views) { views = new Map(); this.textViews.set(text, views); }
    const key = `${source.story}:${source.storyInstance}`;
    let view = views.get(key);
    if (!view) {
      view = Object.freeze({ ...text, shape: (request: Parameters<TextLayoutService['shape']>[0]) => {
        this.chargeWork(source, Math.max(1, request.text.length));
        return text.shape(request);
      } });
      views.set(key, view);
    }
    return view;
  }
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
  /** Debit before work; the global acquisition ledger supplies source size.
   * This ledger survives page replay, new concrete sessions and all field
   * passes. A multi-note rejected group cannot let one note borrow another's
   * allowance. Never refund failed trials or reset an existing permission. */
  charge(kind: NoteKind, id: string, sourceUnits: number, units: number): void {
    const key = `${kind}:${id}`;
    const failure = this.permitted.get(key);
    if (!failure) return;
    sourceUnits = this.sourceSizes.get(key) ?? sourceUnits;
    this.sourceSizes.set(key, sourceUnits);
    const spent = this.work.get(key) ?? 0;
    // Even if a later oversized group exhausts its allowance, v0.89 would
    // already have stopped at the execution's first rejected group.
    if (units > NOTE_RECOVERY_WORK_MULTIPLIER * (sourceUnits + 1) - spent) {
      throw this.originalFailure ?? failure;
    }
    this.work.set(key, spent + units);
  }
  registerSource(source: SourceRef, sourceUnits: number): void {
    const key = `${source.story}:${source.storyInstance}`;
    if (this.permitted.has(key) && !this.sourceSizes.has(key)) this.sourceSizes.set(key, sourceUnits);
  }
  chargeWork(source: SourceRef, units: number): void {
    if (source.story !== 'footnote' && source.story !== 'endnote') return;
    const sourceUnits = this.sourceSizes.get(`${source.story}:${source.storyInstance}`);
    if (sourceUnits !== undefined) this.charge(source.story, source.storyInstance, sourceUnits, units);
  }
  /** Charge actual retained-geometry visits as well as source acquisition.
   * Frozen subtrees have weak, scalar size memos, but visiting/copying them
   * again still pays their complete weight. Debit while sizing an unseen
   * subtree, before the materializer can allocate or traverse it. This covers
   * repeated headers, split rows and nested tables without per-candidate gates.
   * Retained strings are immutable references, not reshaped source text: one
   * field/array-entry visit costs one unit regardless of an ID's length. Text
   * shaping pays UTF-16 source units separately at acquisition. */
  chargeSourceTraversal(source: SourceRef, input: unknown): void {
    this.chargeTraversal(source, input, true);
  }
  chargeRetained(source: SourceRef, input: unknown): void {
    this.chargeTraversal(source, input, false);
  }
  private chargeTraversal(source: SourceRef, input: unknown, sourceTraversal: boolean): void {
    if (source.story !== 'footnote' && source.story !== 'endnote') return;
    const kind = source.story, id = source.storyInstance;
    const sourceUnits = this.sourceSizes.get(`${kind}:${id}`);
    if (sourceUnits === undefined) return;
    const sizes = sourceTraversal ? this.sourceTraversalSizes : this.retainedSizes;
    const visit = (value: unknown, sealed = false): number => {
      if (!value || typeof value !== 'object') {
        const size = 1;
        this.charge(kind, id, sourceUnits, size);
        return size;
      }
      const cached = sizes.get(value);
      if (cached !== undefined) { this.charge(kind, id, sourceUnits, cached); return cached; }
      this.charge(kind, id, sourceUnits, 1);
      let size = 1;
      sealed ||= isDeepFrozenPlainDataRoot(value);
      for (const key in value) {
        if (!Object.hasOwn(value, key)) continue;
        if (sourceTraversal) { this.charge(kind, id, sourceUnits, 1); size++; }
        size += visit((value as Record<string, unknown>)[key], sealed);
      }
      if (sealed) sizes.set(value, size);
      return size;
    };
    visit(input);
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
// Nested text-box stories have their own SourceRef but still spend the parent
// note's work. Service views carry that owner; immutable layouts carry no hooks.
const workOwners = new WeakMap<LayoutServices, SourceRef>();
export function attachNoteRecoveryOwner(services: LayoutServices, source: SourceRef): void {
  workOwners.set(services, source);
}
export function noteRecoveryOwnerOf(services: LayoutServices): SourceRef | undefined {
  return workOwners.get(services);
}
export function attachNoteOverflowPolicy(services: LayoutServices, policy: NoteOverflowPolicy): void {
  policies.set(services, policy);
}
export function noteOverflowPolicyOf(services: LayoutServices): NoteOverflowPolicy | undefined {
  return policies.get(services);
}
export function noteOverflowAllowed(services: LayoutServices, kind: NoteKind, ids: readonly string[]): boolean {
  return policies.get(services)?.allows(kind, ids) === true;
}

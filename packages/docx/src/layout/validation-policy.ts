/**
 * Retained-layout validation policy.
 *
 * The layout engine defends its "retained layout is structured-clone-safe plain
 * data" contract in two layers:
 *
 * - FATAL checks, which always run and are NOT governed by this policy: the
 *   full `assertDocumentLayout` invariant suite on every finished layout
 *   (non-finite or negative geometry, invalid ownership, broken layout
 *   invariants — always fatal, never downgraded to a diagnostic, with no
 *   test/production split; see `assertAndDeepFreezeDocumentLayout`), plus
 *   the backstops fused into the unconditional clone/freeze walks in
 *   `plain-data.ts` and `invariants.ts` (non-cloneable values, foreign
 *   prototypes, non-finite numbers).
 * - STRICT INPUT diagnostics, which this policy gates: the separate
 *   `assertPlainData` pre-pass reports exact property paths and rejects cycles,
 *   accessors, hidden properties, and other non-plain input. Snapshot validation
 *   also uses native structuredClone to reject proxies. These extra input checks
 *   are intended for diagnosing engine bugs; the internal clone/freeze helpers
 *   receive engine-owned data, not arbitrary untrusted objects. Without the
 *   pre-pass, those helpers do not promise to reject every non-plain input.
 *
 * `snapshotPlainData` runs on every block accepted during pagination, re-paid
 * on every convergence pass, so the duplicate pre-pass walk is a real cost on
 * big documents. The policy is:
 *
 * - OFF in every production realm. Tests may mock this internal policy to
 *   exercise the additional reports; runtime environment variables never alter
 *   the layout contract or add a second traversal.
 *
 * Freezing is NOT part of this policy and always runs: retained-layout
 * immutability is load-bearing (the identity WeakSets here and in
 * `invariants.ts`, the `DeepReadonly` contract the variant store hands out, and
 * the assumption that a cached layout cannot be mutated by a consumer).
 */

/** Production omits the extra strict-input diagnostic pre-pass. */
export function documentLayoutValidationEnabled(): boolean {
  return false;
}

/**
 * Retained-layout validation policy.
 *
 * The layout engine defends its "retained layout is structured-clone-safe plain
 * data" contract in two layers:
 *
 * - FATAL checks, which always run and are NOT governed by this policy: the
 *   full `assertDocumentLayout` invariant suite on every finished layout
 *   (non-finite or negative geometry, invalid ownership, broken layout
 *   invariants — fatal per the layout engine's error contract in
 *   docs/docx-layout-engine-redesign.md, with no test/production split), plus
 *   the backstops fused into the unconditional clone/freeze walks in
 *   `plain-data.ts` and `invariants.ts` (non-cloneable values, foreign
 *   prototypes, non-finite numbers).
 * - PATH-PRECISE diagnostics, which this policy gates: the separate
 *   `assertPlainData` pre-pass that walks the graph a second time purely to
 *   report the exact property path of a violation. It is redundant with the
 *   fused backstops — the same violation is still detected and thrown without
 *   it, just with a terser message — so it is a development-time nicety, not a
 *   safety layer.
 *
 * `snapshotPlainData` runs on every block accepted during pagination, re-paid
 * on every convergence pass, so the duplicate pre-pass walk is a real cost on
 * big documents. The policy is:
 *
 * - ON by default under a test runner, so every suite (and therefore CI) keeps
 *   the path-precise reports.
 * - OFF otherwise. Internal tests and benchmarks can change the policy
 *   explicitly when they need to compare validated and production-equivalent
 *   execution.
 *
 * Freezing is NOT part of this policy and always runs: retained-layout
 * immutability is load-bearing (the identity WeakSets here and in
 * `invariants.ts`, the `DeepReadonly` contract the variant store hands out, and
 * the assumption that a cached layout cannot be mutated by a consumer).
 */

/** Whether a test runner is driving this process. Vitest sets `VITEST`; the
 *  `NODE_ENV` check covers other runners and Node's own test tooling. Guarded
 *  for realms without `process` (browser bundles, workers). */
function detectTestEnvironment(): boolean {
  const env = (globalThis as { process?: { env?: Record<string, string | undefined> } })
    .process?.env;
  if (!env) return false;
  return env.VITEST !== undefined || env.NODE_ENV === 'test';
}

let enabled = detectTestEnvironment();

/** Whether the recursive retained-layout contract checks should run. */
export function documentLayoutValidationEnabled(): boolean {
  return enabled;
}

/**
 * Turn the retained-layout contract checks on or off for this realm.
 *
 * This is an internal test and benchmark control, not a public runtime option.
 * Keeping the mutable policy inside the layout module avoids exposing realm-wide
 * process state as part of the document API.
 */
export function setDocumentLayoutValidation(next: boolean): void {
  enabled = next;
}

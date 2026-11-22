/** Inclusive geometry comparisons must tolerate the single-ULP drift produced
 * when the same authored boundary is reconstructed through different retained
 * subtraction/addition paths. This is floating-point precision, not a layout
 * tolerance: a geometrically distinct point value remains outside. Page
 * containment (invariants.ts) and body admission (body-paginator.ts) share
 * this predicate, so admission accepts exactly the drift containment accepts. */
export function atMostWithinFloatingPrecision(left: number, right: number): boolean {
  if (left <= right) return true;
  const precision = Number.EPSILON * Math.max(1, Math.abs(left), Math.abs(right));
  return left - right <= precision;
}

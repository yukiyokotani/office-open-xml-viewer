/**
 * @deprecated The family-keyed Canvas/Word advance allowance was removed.
 * Width must be measured from the selected text route. Retained as a no-op
 * because this helper was exported by released core packages.
 */
export function fontAdvanceBiasEm(_family: string | null | undefined): number {
  return 0;
}

/** ECMA-376 §17.13.5.14/§17.13.5.22: final view omits deleted and
 * moved-away occurrences. All acquisition and retained-coordinate paths use
 * the same visibility decision; markup view retains their original ownership. */
export function revisionIsOmitted(kind: string | undefined, showTrackedChanges: boolean | undefined): boolean {
  return showTrackedChanges !== true && (kind === 'deletion' || kind === 'moveFrom');
}

/** Whether a run is omitted (emits nothing) in the given revision view: the
 * single predicate for every run type (text, picture, chart, shape, field,
 * math, break, tab, symbol, note reference, anchor host ...). The segment
 * builder and the anchor-only classification both call it, so what layout
 * emits and what classification counts cannot disagree. */
export function runIsOmitted(
  run: Readonly<{ revision?: Readonly<{ kind?: string }> | null }>,
  showTrackedChanges: boolean | undefined,
): boolean {
  return revisionIsOmitted(run.revision?.kind, showTrackedChanges);
}

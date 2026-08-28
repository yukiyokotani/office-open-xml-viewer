import type { DocumentLayout } from './types.js';
import { LayoutInvariantError } from './diagnostics.js';

export interface LayoutIteration {
  readonly fingerprint: string;
  readonly pageCount: number;
  readonly layout?: DocumentLayout;
}

export class ExactConvergenceError extends LayoutInvariantError {
  readonly reason: 'cycle' | 'limit';
  readonly states: readonly string[];
  readonly passes: number;

  constructor(
    reason: 'cycle' | 'limit',
    states: readonly string[],
    passes: number,
  ) {
    super(
      'NON_CONVERGENCE',
      reason === 'cycle'
        ? `repeated exact-state cycle at ${states.at(-1) ?? '<missing>'}`
        : `hard exact-state pass limit ${passes} reached`,
    );
    this.name = 'ExactConvergenceError';
    this.reason = reason;
    this.states = Object.freeze([...states]);
    this.passes = passes;
  }
}

export interface ExactConvergenceOptions<T> {
  /** State observed before the first pass. It does not consume the pass budget. */
  readonly seedState?: string;
  /** Deterministic pass. `pass` is one-based and counts against `limit`. */
  readonly step: (previous: T | null, pass: number) => T;
  readonly stateOf: (value: T) => string;
  /** Maximum number of `step` calls, including the confirming fixed-point pass. */
  readonly limit: number;
}

/**
 * Converge any deterministic exact-state transition.
 *
 * Adjacent equality is a fixed point; a non-adjacent repeated state is a cycle.
 * The limit is an explicit resource guard, not a claim about the state-space
 * cardinality. Exhaustion fails closed and never returns the last candidate.
 */
export function convergeExactState<T>(
  options: ExactConvergenceOptions<T>,
): Readonly<{ value: T; passes: number }> {
  // One algorithm, driven to completion. The generator form below is the
  // implementation; this wrapper just refuses to suspend.
  const steps = convergeExactStateSteps<T, never>({
    ...options,
    step: function* generatorStep(previous, pass) {
      return options.step(previous, pass);
    },
  });
  let step = steps.next();
  while (!step.done) step = steps.next();
  return step.value;
}

/**
 * {@link convergeExactState} whose pass is itself suspendable.
 *
 * `step` is a generator, and its yields are forwarded to this generator's
 * consumer, so a caller driving convergence can spread each pass across
 * event-loop turns without the convergence policy — adjacent equality is a
 * fixed point, a non-adjacent repeat is a cycle, exhaustion fails closed —
 * being restated anywhere.
 */
export function* convergeExactStateSteps<T, Y>(
  options: Omit<ExactConvergenceOptions<T>, 'step'> & {
    readonly step: (previous: T | null, pass: number) => Generator<Y, T, void>;
  },
): Generator<Y, Readonly<{ value: T; passes: number }>, void> {
  const { seedState, step, stateOf, limit } = options;
  const minimumLimit = seedState === undefined ? 2 : 1;
  if (!Number.isInteger(limit) || limit < minimumLimit) {
    throw new RangeError(
      `Exact convergence limit must be an integer >= ${minimumLimit}`,
    );
  }
  const states: string[] = seedState === undefined ? [] : [seedState];
  const seen = new Set(states);
  let previous: T | null = null;
  for (let pass = 1; pass <= limit; pass += 1) {
    // Annotated because `yield*` in a generator whose own return type mentions
    // T cannot be inferred without circularity (TS7022).
    const value: T = yield* step(previous, pass);
    const state = stateOf(value);
    const priorState = states.at(-1);
    states.push(state);
    if (priorState === state) {
      return Object.freeze({ value, passes: pass });
    }
    if (seen.has(state)) {
      throw new ExactConvergenceError('cycle', states, pass);
    }
    seen.add(state);
    if (pass === limit) {
      throw new ExactConvergenceError('limit', states, pass);
    }
    previous = value;
  }
  throw new ExactConvergenceError('limit', states, limit);
}

export function convergeLayout<T extends LayoutIteration>(
  seed: T,
  step: (iteration: T) => T,
  limit: number,
): T {
  const steps = convergeLayoutSteps<T, never>(
    seed,
    function* generatorStep(iteration) { return step(iteration); },
    limit,
  );
  let next = steps.next();
  while (!next.done) next = steps.next();
  return next.value;
}

/** {@link convergeLayout} whose step is suspendable; yields are forwarded. */
export function* convergeLayoutSteps<T extends LayoutIteration, Y>(
  seed: T,
  step: (iteration: T) => Generator<Y, T, void>,
  limit: number,
): Generator<Y, T, void> {
  if (!Number.isInteger(limit) || limit < 1) {
    throw new LayoutInvariantError('NON_CONVERGENCE', 'limit must be a positive integer');
  }
  try {
    return (yield* convergeExactStateSteps<T, Y>({
      seedState: seed.fingerprint,
      step: function* layoutStep(previous) { return yield* step(previous ?? seed); },
      stateOf: (iteration) => iteration.fingerprint,
      limit,
    })).value;
  } catch (error) {
    if (error instanceof ExactConvergenceError) {
      throw new LayoutInvariantError(
        'NON_CONVERGENCE',
        error.reason === 'cycle'
          ? `repeated geometry fingerprint cycle at ${error.states.at(-1) ?? '<missing>'}`
          : `hard iteration limit ${limit} reached`,
      );
    }
    throw error;
  }
}

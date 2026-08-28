import { bench, describe } from 'vitest';
import { layoutDocument } from '../document-layout.js';
import {
  installStubCanvas,
  syntheticDocxModel,
  type SyntheticDocumentShape,
} from '../testing/synthetic-document.js';
import { setDocumentLayoutValidation } from './validation-policy.js';

/**
 * Layout-cost benchmark. Run with `pnpm bench:layout` (vitest picks up
 * `*.bench.ts` only under `vitest bench`, so this never slows `pnpm test`).
 *
 * The point of the shape matrix is attribution, not a single headline number:
 * `paginateBody` wraps its block loop in several whole-document fixed-point
 * solvers, and each shape turns on a different one (see
 * `testing/synthetic-document.ts`). A change that helps the seed pass shows up
 * on `plain`; one that helps convergence shows up as a much bigger delta on
 * `fields`.
 *
 * Measurement runs against the deterministic stub canvas, so the numbers are
 * pure layout cost: no font loading, no real text shaping, no WASM parse. That
 * makes them comparable across machines and across runs, and it means a
 * regression here is a regression in the engine rather than in the environment.
 *
 * Each case is benchmarked twice — with the retained-layout contract checks on
 * and off (see `validation-policy.ts`) — because "on" is what CI pays and "off"
 * is what a shipped viewer pays.
 */

const CASES: readonly (readonly [SyntheticDocumentShape, number])[] = [
  ['plain', 200],
  ['header-footer', 200],
  ['fields', 200],
  ['tables', 60],
  ['long-paragraphs', 6],
];

// Layout is measured in seconds per iteration, so the default sampling budget
// would run for many minutes. A handful of iterations is enough to separate the
// effects being measured here.
const OPTIONS = { iterations: 3, warmupIterations: 1, time: 0, warmupTime: 0 } as const;

installStubCanvas();

for (const [shape, paragraphs] of CASES) {
  describe(`${shape} (n=${paragraphs})`, () => {
    bench('validation off (production)', () => {
      setDocumentLayoutValidation(false);
      layoutDocument(syntheticDocxModel(shape, { paragraphs }));
    }, OPTIONS);

    bench('validation on (test/CI)', () => {
      setDocumentLayoutValidation(true);
      layoutDocument(syntheticDocxModel(shape, { paragraphs }));
    }, OPTIONS);
  });
}

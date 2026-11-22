# Testing

## Verification

Prefer focused verification tied to the change. Common commands:

```bash
pnpm test
pnpm typecheck
pnpm build:wasm
pnpm storybook
pnpm build-storybook
```

If a check fails because ignored/generated WASM is stale, rebuild WASM before
assuming the source change is broken. Rebuild generated WASM (`pnpm
build:wasm` or the per-package `wasm` script) before parser-backed
integration tests, Storybook, or VRT whenever Rust source changed or
`packages/*/src/wasm/` may be stale: stale WASM both hides real parser
defects and fabricates failures.

## Test value and consolidation

For every test, answer: **What realistic defect would go undetected if this
test were deleted?** Compare that protection with execution cost, flakiness,
fixture complexity, and the maintenance burden of legitimate refactoring.

- Do not retain or add tests merely as a precaution, to preserve a test count,
  or to maintain a coverage percentage. TDD means a focused regression for the
  actual defect; it does not require keeping every exploratory case.
- Delete implementation copies, tests that only verify their own mocks or
  fixtures, and assertions that merely restate type-system or dependency
  guarantees. Exercise production behavior with an independent expectation.
- Consolidate overlapping tests under the narrowest meaningful production
  boundary. Keep a distinct case only when it catches a distinct plausible
  failure: for example precedence, a real boundary, ownership, or a race.
- Avoid asserting private fields, helper call order, source text, or exact
  internal object shape when an observable contract can express the guarantee.
  Do not replace behavior tests with snapshots of implementation details.
- Reduce exploratory parameter matrices to representative cases and meaningful
  boundaries.
- Shared-helper coverage does not prove each format is wired correctly. Keep
  focused DOCX/XLSX/PPTX integration cases where their production paths
  differ; do not repeat the helper's entire input matrix in every format.
- In a test audit, identify the exact cases to remove or merge, the remaining
  test that supplies any claimed replacement, and the realistic protection
  lost. Do not report a static review as measured runtime savings.

# @silurus/ooxml-xlsx — for XLSX sessions

Read the root `AGENTS.md` (canonical) first. Follow it for workflow, branches,
VRT, and reference-image rules. This file records only XLSX-specific facts.

## Edit Scope

- Mainly `packages/xlsx/**` (including the Rust code in `parser/` and `model/`).
- Change `packages/core` / `packages/ooxml-common` or other formats only as part
  of a coherent fix that follows Cross-Package Integration in the root
  `AGENTS.md`.

## Formula Handling (policy)

- Cell formulas (`<f>`) are never calculated. The display always uses the cached
  value (`<v>`). `TODAY()` / `NOW()` are no exception (if only one cell showed
  the current time, cells derived from it would keep their cached values and the
  display would become inconsistent). A formula cell without a cached value is
  treated as a cell with no (empty) value regardless of the function, and its
  number format is applied to that value (some number formats still produce
  text).
- Conditional-formatting formulas (`expression`, etc.) are evaluated by a
  partial evaluator (`src/formula.ts`) that takes cached values as input.
  `TODAY()` / `NOW()` inside this evaluator refer to the current time. Full
  evaluation is tracked in #1547.

## Directory Layout (main entries)

```
packages/xlsx/
├── parser/src/          Rust parser (lib.rs, styles.rs, drawing.rs, chart.rs,
│                        pivot.rs, worksheet_*.rs, …)
├── model/src/           Rust model crate
├── src/
│   ├── index.ts         public API
│   ├── types.ts         1:1 with the Rust output
│   ├── workbook.ts      XlsxWorkbook
│   ├── renderer.ts      Canvas 2D renderer
│   ├── worker.ts / render-worker.ts  parse worker / render worker
│   └── wasm/            wasm-pack output (generated, git-ignored)
├── public/demo/         tracked demo sample (sample-1.xlsx)
├── public/private/xlsx/ local-only private corpus (git-ignored)
└── tests/visual/        Playwright specs (visual.spec.ts, worker-equivalence.spec.ts, …)
    └── references/demo/<name>/sheet-N.png + scores.json
                         adopted renderer output (not an Excel export). Never update automatically.
```

## WASM Build

```bash
pnpm --filter @silurus/ooxml-xlsx wasm
```

This runs `wasm-pack build --target web --out-dir ../src/wasm` in `parser/` and
then applies `scripts/append-wasm-reinit.mjs`.

## Tests

- Unit tests: `pnpm test` (vitest) at the repository root.
- VRT: `pnpm --filter @silurus/ooxml-xlsx vrt` / `vrt:private` / `vrt:fidelity`.
  See the root `AGENTS.md` for the baseline environment variables.
- Storybook: `pnpm storybook` at the repository root.
- Do not commit samples: never commit the private corpus or generated output.

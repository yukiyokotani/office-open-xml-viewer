# Repository scripts

Run every command from the repository root. Most scripts are reached through a
`package.json` script or a CI workflow; the name in brackets is the entry point
to prefer. A `*.test.mjs` file next to a script is its `node --test` suite.

The scripts stay in one directory because generated files, public API
baselines and package source comments name these paths.

## Generators of committed data

Rerun one when its input changes, then commit the regenerated output with it.

| Script | Generates | Rerun |
| --- | --- | --- |
| `generate-ooxml-resource-policy.mjs` | `packages/core/src/worker/resource-policy.generated.ts` and `packages/ooxml-common/src/resource-policy.generated.rs` from `packages/ooxml-common/resource-policy.json` | `pnpm generate:resource-policy` (CI runs `pnpm check:resource-policy`) |
| `generate-spreadsheet-style-presets.mjs` | `packages/ooxml-common/src/spreadsheet_style_presets.generated.rs` from ECMA-376 Annex G `presetTableStyles.xml` | `node scripts/generate-spreadsheet-style-presets.mjs <presetTableStyles.xml>` |
| `generate-region-map-asset.mjs` | `packages/core/src/chart/region-map-natural-earth.generated.ts` from Natural Earth GeoJSON | `node scripts/generate-region-map-asset.mjs <geojson> [output]` |
| `generate-font-reference-metrics.py` | `packages/core/src/fonts/reference-font-metrics-data.json` and `reference-font-metrics-provenance.json` from the fonts installed with Office and macOS | `uv run --no-project --with fonttools==4.67.0 python scripts/generate-font-reference-metrics.py` (test: `pnpm test:python`) |
| `generate-font-demand-inert-ranges.py` | `packages/core/src/fonts/inert-scalar-ranges.ts` from Unicode 17 `UnicodeData.txt` | `python3 scripts/generate-font-demand-inert-ranges.py --unicode-data <UnicodeData.txt>` (`--check` to verify) |
| `generate-pptx-east-asian-defaults.py` | `packages/pptx/src/east-asian-default-coverage.ts` from PowerPoint's bundled East Asian fonts | `uv run --with fonttools python scripts/generate-pptx-east-asian-defaults.py` |
| `generate-fixture-font.py` | The synthetic test fonts in `tests/fixtures/fonts/` | `uv run --no-project --with fonttools==4.67.0 python scripts/generate-fixture-font.py` |

Other committed generators live beside their output:

- `packages/core/scripts/gen-*.mjs` and `packages/docx/scripts/gen-arabic-joining.mjs`
  — Unicode property tables (`*.generated.ts`) from the UCD files next to them;
  run each with `node <path>`.
- `packages/pptx/scripts/` — preset shape and text-warp geometry JSON; see
  its [README](../packages/pptx/scripts/README.md).
- `tests/patterns/fixtures/build-*.py` — the synthetic PPTX pattern fixtures
  beside them; run with `python3 <path>`.
- `fuzz/generate-part-corpus.py` — part-level fuzz seeds from public demo
  packages; run with `python3 fuzz/generate-part-corpus.py`.
- `packages/xlsx/tests/visual/gen-validation-fixture.py` — the ignored
  data-validation fixture used by `validation-dropdown.spec.ts`.
- `packages/node/scripts/generate-synthetic-xlsx.mjs` — synthetic workbooks
  generated at test time.

## Build steps

| Script | Purpose |
| --- | --- |
| `append-wasm-reinit.mjs` | Adds the `reinit()` export to wasm-bindgen glue (each package's `wasm` script). |
| `build-model-source-sidecars.mjs` | Builds the self-contained optional model-source worker sidecars (`pnpm build`). |
| `bundle-declarations.mjs` | Bundles the published `.d.ts` entries (`pnpm build`). |
| `docx-compatibility-build.mjs` | Freezes DOCX compatibility rule metadata for Vite, Storybook and the site. Its evidence ledger is `docx-compatibility-microsoft-evidence.json`. |
| `optional-renderer-markers.mjs` | Emits the ownership markers that the optional-renderer checks read. |
| `module-stats-plugin.mjs` | Build diagnostics for `check-model-source-lazy-graph.mjs` when `OOXML_MODULE_STATS` is set. |
| `build-worker-consumer-fixture.mjs` | Builds the consumer app for `pnpm test:production-render-workers`. |

## CI checks

Each check has a `pnpm check:*` or `pnpm test:*` script, or is called directly
from `.github/workflows/ci.yml`.

- Public API: `check-public-api.mjs`, `check-public-type-exports.mjs`,
  `check-viewer-api-symmetry.mjs`.
- Bundle shape: `check-no-inline-wasm.mjs`, `check-relative-wasm-urls.mjs`,
  `check-production-render-workers.mjs`, `check-font-source-tree-shaking.mjs`,
  `check-optional-{chart-ex,region-map,three-d,tiff}.mjs`,
  `check-optional-legacy-source.mjs`, `check-legacy-wasm-exports.mjs`.
- Module boundaries: `check-core-legacy-boundary.mjs`,
  `check-docx-layout-boundaries.mjs`, `check-model-source-lazy-graph.mjs`,
  `check-model-source-lazy-runtime.mjs` (with its `model-source-runtime-*.mjs`
  helpers), `check-model-source-ooxml-path.mjs`.
- DOCX compatibility evidence: `check-docx-compatibility-evidence.mjs`.
- Python generator tests: `run-python-tests.py` (`pnpm test:python`).

## Release

| Script | Purpose |
| --- | --- |
| `sync-companion-version.mjs` | Pins the companion packages to the root version (`pnpm sync:versions`; CI runs `pnpm check:versions`). |
| `check-packed-packages.mjs` | Installs the packed tarballs in an isolated consumer (`pnpm test:packed-packages`). |
| `publish-packages.mjs` | Publishes the root and companion packages (`publish.yml`). |

## Measurements

- `measure-bundle-size.mjs` — the asset sizes shown on the site's bundle-size
  page, after `pnpm build:wasm && pnpm build`.
- `measure-ooxml-archive-entries.mjs` — archive entry statistics used to
  calibrate the resource policy (`pnpm measure:archive-entries`).

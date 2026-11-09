# @silurus/ooxml-pptx — for PPTX sessions

Package supplement for PPTX sessions. Read the root `AGENTS.md` first; it is
canonical for workflow, branching, VRT and reference-image rules, and the root
tool-specific files (for example `CLAUDE.md` for autonomous-work hours) add
agent-specific supplements. This file only records PPTX-specific facts.

## Package Layout

```
packages/pptx/
├── parser/            Rust/wasm-pack parser crate (src/lib.rs, shape.rs, text.rs,
│                      master.rs, theme.rs, chart.rs, …); emits the slide model
├── model/             Rust model crate shared by the parser
├── src/
│   ├── index.ts       public entry (re-exports PptxViewer, PptxPresentation, types)
│   ├── viewer.ts      PptxViewer
│   ├── presentation.ts PptxPresentation
│   ├── renderer.ts    Canvas 2D slide renderer
│   ├── worker.ts / render-worker.ts  WASM parse worker and render worker
│   ├── types.ts       TypeScript model types matching the Rust output
│   └── wasm/          wasm-pack output (generated, git-ignored)
├── public/demo/       tracked demo deck (sample-1.pptx)
├── public/private/pptx/  local-only private corpus (git-ignored)
└── tests/visual/      Playwright specs (visual.spec.ts, worker-equivalence, …),
                       fixture.html (default width 1280, `?width=` overrides) and
                       references/demo/<name>/slide-N.png + scores.json
```

## WASM Build

```bash
pnpm --filter @silurus/ooxml-pptx wasm
```

This runs `wasm-pack build --target web --out-dir ../src/wasm` in `parser/` and
then `scripts/append-wasm-reinit.mjs`. There is no manual copy step. Rebuild
whenever Rust parser source changes or `src/wasm/` may be stale.

## Tests

- Unit tests run from the repository root with `pnpm test` (vitest).
- Visual tests: `pnpm --filter @silurus/ooxml-pptx vrt` (self-VRT against a
  baseline checkout), `vrt:private` (private self-VRT plus private worker
  parity), `vrt:fidelity` (demo ratchet against `references/demo/`). See the
  root `AGENTS.md` for the required baseline environment variables.
- Storybook is unified at the repository root (`pnpm storybook`).

## OOXML Units

- Rotation: `rot / 60000` → degrees.
- Font size: `sz / 100` → pt (`2400` → 24 pt).
- `spcPts` spacing is in hundredths of a point.
- Autofit: `normAutofit` (`fontScale`, `lnSpcReduction`) shrinks text;
  `spAutoFit` resizes the shape and does not shrink text (ECMA-376 `a:bodyPr`
  autofit children).

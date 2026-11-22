# @silurus/ooxml-docx — for DOCX sessions

Read the root `AGENTS.md` (canonical) first. Follow it for workflow, branches,
VRT, and reference-image rules. This file records only DOCX-specific facts.

## Edit Scope

- Mainly `packages/docx/**` (including the Rust code in `parser/` and `model/`).
- Change `packages/core` / `packages/ooxml-common` or other formats only as part
  of a coherent fix that follows Cross-Package Integration in the root
  `AGENTS.md`.

## Directory Layout (main entries)

```
packages/docx/
├── parser/src/          Rust parser (lib.rs, parser.rs, styles.rs, numbering.rs, …)
├── model/src/           Rust model crate
├── src/
│   ├── index.ts         public API
│   ├── document.ts      DocxDocument
│   ├── viewer.ts        DocxViewer
│   ├── types.ts         model types (1:1 with the Rust output)
│   ├── layout/          pagination and layout
│   ├── line-breaker/    line breaking
│   ├── paint/           Canvas painting
│   ├── render-worker.ts render worker
│   └── wasm/            wasm-pack output (generated, git-ignored)
├── public/demo/         tracked demo sample (sample-1.docx)
├── public/private/docx/ local-only private corpus (git-ignored)
└── tests/visual/        Playwright specs (visual.spec.ts, conformance.spec.ts, …)
    └── references/demo/<name>/page-N.png + scores.json
                         adopted renderer output (not a Word export). Never update automatically.
```

## WASM Build

```bash
pnpm --filter @silurus/ooxml-docx wasm
```

This runs `wasm-pack build --target web --out-dir ../src/wasm` in `parser/` and
then applies `scripts/append-wasm-reinit.mjs`.

## Tests

- Unit tests: `pnpm test` (vitest) at the repository root.
- VRT: `pnpm --filter @silurus/ooxml-docx vrt` / `vrt:private` / `vrt:fidelity` /
  `conformance`. See the root `AGENTS.md` for the baseline environment variables.
- Storybook: `pnpm storybook` at the repository root.

## No Heuristics (consistent with the OOXML Implementation Policy in the root `AGENTS.md`)

DOCX rendering, especially around anchors, wgp, docGrid, and ruby, often tempts
you to add heuristics. Observe the following:

- **Anchor position**: the meaning of `relativeFrom` and `posOffset` / `align` / `pctPos` in `positionH/V` is fully defined in ECMA-376 §20.4.3.x. Constants or special branches "to match sample-N" are forbidden.
- **wgp group transform**: `lIns/tIns/rIns/bIns` are page-space EMU (measured in the dimensions after the group transform). `anchorXPt/anchorYPt` are made page-relative by multiplying the child-coordinate-space offset by sx/sy. `widthPt/heightPt` are likewise cx/cy × sx/sy. Unless all of these receive consistent scaling, position errors accumulate.
- **Font classification**: use `<w:family w:val=…/>` in `word/fontTable.xml` (§17.8.3.10), not name pattern matching. Keep pattern matching only as a fallback for fonts missing from the font table.
- **Page-break hints**: `<w:lastRenderedPageBreak/>` (§17.3.1.20) is Word's layout cache. "Trusting it only on the particular paths where our own computation is broken" is forbidden. Either fix our own computation or always ignore the hint. Do not mix the two.
- **Specification default values**: `unwrap_or(0.0)` for an omitted attribute is usually wrong. When ECMA-376 defines a default value, such as lIns=91440 EMU in §21.1.2.1.1, always use that value.

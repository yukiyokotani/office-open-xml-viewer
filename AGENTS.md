# Agent Instructions

This file is the canonical workflow for AI coding agents working on this
repository. Tool-specific files such as `CLAUDE.md` may add supplements, but they
must not contradict this file.

## Project

`office-open-xml-viewer` renders OOXML documents (`docx`, `xlsx`, `pptx`) to
browser Canvas. The implementation is a Rust/WASM parser layer plus TypeScript
Canvas renderers and viewers.

Important directories:

- `packages/core/` — shared rendering primitives and shared TypeScript helpers.
- `packages/ooxml-common/` — shared Rust parsing/model helpers.
- `packages/docx/` — DOCX parser, renderer, viewer, and tests.
- `packages/xlsx/` — XLSX parser, renderer, viewer, and tests.
- `packages/pptx/` — PPTX parser, renderer, viewer, and tests.
- `site/` — public site.
- `.storybook/` — unified Storybook config.

## Standard Development Flow

Use this flow for ordinary bug fixes and feature work:

1. Start from `main` and run `git pull --ff-only`.
2. If the checkout is dirty, preserve unrelated work before switching or pulling.
   Prefer `git stash push -u -m <clear-message>` for work that appears to belong
   to another session.
3. Create a feature branch. Use `codex/<topic>` for Codex work when possible;
   otherwise follow the repository's existing branch naming.
4. When creating or entering a Codex/Claude worktree, make the root checkout's
   ignored `spec/` directory available from that worktree as a local symlink.
   For the standard `.claude/worktrees/<name>` layout, run:

   ```bash
   ln -s ../../../spec spec
   ```

   If the worktree lives elsewhere, create `spec` as a symlink to the main local
   checkout's `spec/` directory. Do not copy or commit the specification files.
5. Work with TDD for bug fixes:
   - explore the failing sample or behavior,
   - add a focused failing test,
   - make it pass,
   - refactor only where it reduces real complexity.
6. Verify with the narrowest meaningful tests first, then broader checks when the
   touched surface warrants it.
7. Commit only the intended source/docs/test files. Do not commit private samples
   or generated local artifacts.
8. Push the feature branch, create a PR, wait for checks when practical, and merge
   with a merge commit. Do not squash.

## Git and PR Rules

- Never push directly to `main`.
- Use PRs for integration into `main`.
- Do not squash merge. Use `gh pr merge <number> --merge` unless the user
  explicitly asks for another non-squash strategy.
- Before `git push`, set `git config http.postBuffer 524288000` if large pushes
  are expected.
- Commit messages should match the repository style:
  - subject: `fix(scope): ...`, `test(scope): ...`, `refactor(scope): ...`, etc.
  - body: explain the root cause, the fix, and verification.
  - include specification sections or observed Office behavior when relevant.
- PR titles, PR descriptions, commit messages, and public docs must be safe for
  an OSS repository:
  - do not include local absolute paths, usernames, home directories, machine
    names, or other personal environment details;
  - do not quote or summarize the concrete contents of private or copyrighted
    samples;
  - describe private-sample-driven work in terms of the implementation problem,
    the OOXML/Office behavior mismatch, the general class of affected documents,
    and the verification command or local-only comparison that was performed;
  - keep local-only sample filenames and visual details out of public PR text
    unless the user explicitly approves publishing that information.
- Agent-authored commits should include a matching co-author trailer. Include
  the actual model used for that task; do not hard-code a model name across
  tasks. Examples:
  - Codex: `Co-Authored-By: Codex <model-name> <noreply@openai.com>`
  - Claude/Fable: `Co-Authored-By: Claude <model-name> <noreply@anthropic.com>`
- If Claude starts or delegates work to Codex, include both trailers in the
  resulting commit, including the model each agent actually used for that task.
  This preserves the human-readable provenance of the orchestration layer and
  the implementation agent.

## Verification

Prefer focused verification tied to the change. Common commands:

```bash
pnpm test
pnpm typecheck
pnpm build:wasm
pnpm storybook
pnpm build-storybook
```

For a narrow DOCX renderer fix, a good verification set is:

```bash
./node_modules/.bin/vitest run \
  packages/docx/src/cell-border-conflict.test.ts \
  packages/docx/src/cell-border-conflict-render.test.ts \
  packages/docx/src/column-widths.test.ts

./node_modules/.pnpm/typescript@*/node_modules/typescript/bin/tsc --build --pretty false
git diff --check
```

If a check fails because ignored/generated WASM is stale, rebuild WASM before
assuming the source change is broken.

Before running parser-backed integration tests, Storybook, or VRT, rebuild
generated WASM (`pnpm build:wasm` or the per-package command) whenever parser
(Rust) source has changed or `packages/*/src/wasm/` may be stale — stale
generated WASM can both hide real parser defects and fabricate failures.

## OOXML Implementation Policy

Be specification-first.

- Before implementing OOXML behavior, consult the relevant documents under the
  local `spec/` symlink. Prefer ECMA-376 / ISO-29500 for normative behavior, and
  Microsoft extension notes such as `[MS-DOCX]`, `[MS-XLSX]`, `[MS-PPTX]`, and
  `[MS-ODRAWXML]` when Office-specific behavior or extensions are involved.
- Record the relevant specification section, schema element, or observed Office
  behavior in the commit body or code comment when it materially explains the
  implementation.
- Prefer ECMA-376 / ISO-29500 behavior over sample-specific tuning.
- Do not add heuristics only to improve one VRT/sample number, such as arbitrary
  thresholds, empirical scaling constants, or special-casing a sample path.
- A user report like "this line is too long" or "this object is missing" is a hint
  that an OOXML rule is missing or mis-modeled. Identify the related section and
  implement the rule for the whole class of documents.
- When behavior differs from Word/Excel/PowerPoint, first check whether the parser
  discarded necessary information during parsing, inheritance, or style merging.
  If information is missing, extend the parser/model instead of guessing in the
  renderer.
- If Office behavior is not fully specified and requires reverse engineering,
  explain that explicitly and ask for user approval before adding a compatibility
  heuristic.
- If a temporary heuristic is unavoidable, mark it clearly in code with the
  missing spec/implementation work and track it for removal.

## Cross-Package Integration

DOCX, XLSX, and PPTX share many ECMA-376 concepts: DrawingML images, `srcRect`,
text runs, paragraph properties, themes, fills, effects, and shape presets.

- Treat one package's bug as a signal to inspect the sibling packages.
- If the concept is shared, fix the shared layer (`packages/ooxml-common` or
  `packages/core`) where possible.
- Do not duplicate pure parsing predicates or common type definitions across
  packages unless the format truly diverges.
- Avoid false abstraction. Renderer logic that is tightly coupled to a package's
  layout model may stay local.
- Cross-package fixes may touch `core` / `ooxml-common` and multiple packages in
  one PR when that is the coherent unit of work.

## Mandatory Adversarial Review

Before declaring an OOXML feature or bug fix complete, perform an adversarial
review of the final diff. This review is required even when focused tests and VRT
already pass. Record the conclusion and any material evidence in the task handoff
or PR description; do not silently waive an unresolved item.

- **Specification fidelity:** Verify the implementation against the applicable
  ECMA-376 / ISO-29500 rule and Microsoft extension notes. Distinguish normative
  behavior, documented Office extensions, and implementation policy such as
  resource governance. Do not claim specification support from sample output
  alone.
- **No unexplained heuristics:** Search for sample-specific branches, magic
  thresholds, empirical scale factors, path/name checks, and renderer guesses.
  Remove them or justify them under the OOXML Implementation Policy. A passing
  VRT is not evidence that a heuristic is correct.
- **Bounded inference for implicit behavior:** When Word, Excel, or PowerPoint
  behavior is not fully specified, base the inference on sufficiently varied,
  Office-produced tests that exercise boundaries and counterexamples—not one
  private sample. State the observed behavior, the tested range, and the limits
  of the compatibility claim. Keep the implementation scope no broader than the
  evidence supports, and obtain user approval before adding a compatibility
  heuristic.
- **Performance and resource safety:** Review asymptotic work, allocation size,
  decoded and retained memory, cache ownership/eviction, concurrency, worker
  transfer cost, repeated parsing or decoding, hot paths, and failure cleanup.
  Reject unbounded fan-out or whole-document work when visible/page-local work is
  sufficient. Add budget, concurrency, cache, or stress tests where the risk is
  material.
- **Module boundaries and single responsibility:** Confirm that shared OOXML
  concepts and pure rendering/resource primitives live in `ooxml-common` or
  `core`, while DOCX, XLSX, and PPTX keep only format-specific parsing, layout,
  orchestration, and painting. Check for duplicated policy, false abstractions,
  package cycles, and classes/functions that mix unrelated responsibilities.
- **Cross-format wiring:** For every shared capability, explicitly inspect DOCX,
  XLSX, and PPTX integration points. Verify the relevant parser/model, public
  options, viewer, renderer, cache/loader, worker protocol, error propagation,
  documentation, and tests for each applicable format. If one format is
  intentionally excluded, document the format-level reason.
- **Regression evidence:** Run focused behavioral tests plus the broadest
  proportionate cross-package checks. For renderer changes, use the previous
  renderer as the self-VRT oracle and keep Office-reference fidelity evaluation
  separate. Investigate every unexplained difference; do not weaken comparison
  thresholds merely to make the review pass.

## Private Samples and Storybook

Private Office samples and local-only sample stories are gitignored and must not
be committed.

Typical local-only paths:

- `packages/docx/public/private/docx/`
- `packages/xlsx/public/private/xlsx/`
- `packages/pptx/public/private/pptx/`
- `packages/*/src/*privateDemo.stories.ts`
- `packages/*/src/wasm/`

When using a Codex/Claude worktree, copy local private stories/samples from the
main local checkout if needed. Avoid copying `.DS_Store`, temporary Office lock
files, screenshots, VRT diffs, or private reference images unless the user
explicitly asks.

Before using Storybook for parser-backed samples, rebuild WASM from the current
worktree source:

```bash
pnpm build:wasm
```

or per package:

```bash
cd packages/docx/parser && wasm-pack build --target web --out-dir ../src/wasm && node ../../../scripts/append-wasm-reinit.mjs ../src/wasm/docx_parser.js
cd packages/xlsx/parser && wasm-pack build --target web --out-dir ../src/wasm && node ../../../scripts/append-wasm-reinit.mjs ../src/wasm/xlsx_parser.js
cd packages/pptx/parser && wasm-pack build --target web --out-dir ../src/wasm && node ../../../scripts/append-wasm-reinit.mjs ../src/wasm/pptx_parser.js
```

Storybook is unified at the repository root. Static prefixes are:

- `packages/docx/public/` → `/docx/`
- `packages/xlsx/public/` → `/xlsx/`
- `packages/pptx/public/` → `/pptx/`

Use those prefixes when fetching sample files.

In sandboxed agent environments, Storybook's port detection may fail with
`EPERM` while probing `0.0.0.0`. Running Storybook outside the sandbox for local
verification is acceptable when the user asks to view it.

## Visual Regression Tests

VRT is local-only because private samples are not redistributable.

Regression detection must compare the changed renderer with images produced by
the previous renderer. It must not use Word/Excel/PowerPoint or PDF reference
images as the regression oracle. Keep Office-reference fidelity evaluation as a
separate run:

```bash
pnpm build:wasm
pnpm vrt:snapshot # run from main / the merge-base before the renderer change
pnpm vrt
pnpm --filter @silurus/ooxml-docx vrt
pnpm --filter @silurus/ooxml-xlsx vrt
pnpm --filter @silurus/ooxml-pptx vrt
pnpm vrt:fidelity # separate Office/PDF reference fidelity evaluation
```

If the change is already in progress and no trustworthy self-baseline exists,
capture `vrt:snapshot` in a temporary worktree at the merge-base, then run `vrt`
from the changed worktree against those images. A missing baseline is not a
successful regression check; report or correct stale sample/page manifests
instead of treating their skips as coverage.

For the complete private corpus, bind both runs to the same full merge-base SHA.
Use a clean detached worktree for the previous renderer, rebuild its WASM, and
use distinct ports so a candidate server can never satisfy the baseline run:

```bash
git worktree add --detach /tmp/ooxml-vrt-baseline <merge-base-sha>
cd /tmp/ooxml-vrt-baseline
pnpm build:wasm
VRT_BASELINE_REVISION=<full-merge-base-sha> VRT_PORT=5190 pnpm vrt:private:snapshot

cd <candidate-worktree>
pnpm build:wasm
VRT_BASELINE_REVISION=<full-merge-base-sha> VRT_PORT=5191 pnpm vrt:private
```

Snapshot mode rejects a dirty tracked checkout. For the one-time bootstrap of
this harness against a merge-base that predates it, copy only the allowlisted VRT
harness files listed in `tests/visual/private-corpus.mjs` and set
`VRT_ALLOW_HARNESS_CHANGES=1`; renderer/parser changes remain forbidden. The
baseline manifests record the resolved SHA, each private input's SHA-256, and
the exact page/sheet/slide sets.

Only update references when the user explicitly asks:

```bash
UPDATE_REFS=1 pnpm vrt:fidelity
```

Do not automatically update files under `packages/*/tests/visual/references/`.

An exact private self-VRT difference is a detected regression candidate, even
when every changed document carries the setting being implemented. Enumerate
every changed page, sheet, or slide and adjudicate each one against an
Office-produced reference. A manual review counts only when it explicitly
compares the candidate and previous renderers with that Office output, or when
the user explicitly adjudicates the difference. Never approve a group of
differences solely because the affected files share a feature flag or OOXML
element. If neither an Office reference nor user adjudication is available,
report that item as unadjudicated rather than calling the VRT successful.

## Documentation and Public Examples

Public docs and commit messages should be written in English.

README, docs, and Storybook public-facing code examples should:

- use `as Type` assertions instead of postfix non-null assertions,
- use `canvas` for canvas-backed viewers (`DocxViewer`, `PptxViewer`),
- use `container` for container-backed viewers (`XlsxViewer`),
- avoid documenting private/local-only sample paths as public API.

Keep announcements concise and user-facing:

- write first for readers who do not know the implementation: lead with the
  visible outcome and group low-level fixes by the user problem they solve;
- state explicitly whether migration is required. When it is, give the minimum
  steps and name any changed default or behavior; when it is not, say so plainly;
- omit implementation detail that does not change a user's decision or action.
  If a technical qualification is useful, put it in a final `Technical note`
  section after the outcome and migration guidance;
- keep version-sensitive bundle measurements on the stable `/bundle-size` page,
  not in release announcements;
- do not link API reference details directly to release notes. Use a stable
  reference page when more detail is necessary.

## Release Workflow

When the user asks for a release, prepare a dedicated PR named
`release/<version>` and do all release changes there.

Release checklist:

1. Update README screenshots in `docs/images/{pptx,docx,xlsx}.png`.
2. Review README against the current implementation:
   - ESM-only package format,
   - install/import examples,
   - feature support table,
   - bundle-size notes,
   - deprecated API or stale version references.
3. Sync public API documentation:
   - `site/src/lib/api-reference.ts`,
   - `site/src/components/Capabilities.astro` when capabilities changed.
   - `site/src/pages/bundle-size.astro` with the latest production measurements.
4. Add a new top `CHANGELOG.md` section:
   `## 0.x.0 — YYYY-MM-DD`.
5. Bump all versions consistently:
   - root `package.json`,
   - `packages/{core,pptx,xlsx,docx,markdown,node,vscode-extension}/package.json`,
   - `site/package.json`,
   - `packages/mcp-server/Cargo.toml`.
6. Run `cargo check -p ooxml-mcp-server` after bumping the MCP server so
   `Cargo.lock` follows.
7. Open PR `chore(release): 0.x.0`.
8. Merge with `gh pr merge <number> --merge` or another non-squash strategy.
9. Pull main, create and push the annotated tag:
   `git tag -a v0.x.0 -m "v0.x.0"` and `git push origin v0.x.0`.
10. Create the GitHub Release with notes based on the changelog and a full
    changelog comparison link.

Reference images under `tests/visual/references/` are outside the release
checklist unless the user explicitly asks to update them.

## VS Code Extension Release

The VS Code extension version should match the npm library version. Pushing a
`v*` tag triggers `.github/workflows/publish-vscode-extension.yml` to publish the
extension. For manual packaging checks, use the workflow dispatch dry run.

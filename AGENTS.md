# Agent Instructions

This file is the single source of agent rules for this repository. It applies
to every AI coding agent and every tool.

Where each kind of knowledge lives:

- `AGENTS.md` (this file) — repository-wide rules. Keep it to rules.
- `packages/<name>/AGENTS.md` — facts specific to one package. They add to this
  file and do not repeat it.
- `.agents/skills/<name>/SKILL.md` — procedures (runbooks, checklists),
  loaded on demand. Codex reads `.agents/skills/`; Claude Code reads each one
  through a relative symlink `.claude/skills/<name>` ->
  `../../.agents/skills/<name>`.
- `.claude/CLAUDE.md` (imports `../AGENTS.md`) and `packages/<name>/CLAUDE.md`
  (imports `AGENTS.md`) — one-line files for Claude Code. They hold no rules.

Add or change rules only in an `AGENTS.md`, and procedures only in a shared
skill. Do not add tool-specific rule files. Codex stops reading the
`AGENTS.md` chain at 32 KiB (`scripts/check-agent-instructions.test.mjs`).

## Project

`office-open-xml-viewer` renders OOXML documents (`docx`, `xlsx`, `pptx`) to
browser Canvas. The implementation is a Rust/WASM parser layer plus TypeScript
Canvas renderers and viewers.

- `packages/core/` — shared rendering primitives and TypeScript helpers.
- `packages/ooxml-common/` — shared Rust parsing/model helpers.
- `packages/{docx,xlsx,pptx}/` — per-format parser, renderer, viewer, tests.
- `packages/legacy-office/`, `packages/opendocument/` — opt-in companion
  packages for DOC/XLS/PPT and ODF.
- `site/` — public site.

## Standard Development Flow

1. Start from `main` and run `git pull --ff-only`.
2. If the checkout is dirty, preserve unrelated work before switching or
   pulling. Prefer `git stash push -u -m <clear-message>` for work that
   appears to belong to another session.
3. Create a feature branch named `<agent>/<topic>`, for example
   `codex/<topic>` or `claude/<topic>`.
4. When creating or entering a Git worktree, make the root checkout's ignored
   `spec/` directory available as a local symlink. For a worktree under
   `.claude/worktrees/<name>`, run `ln -s ../../../spec spec`; elsewhere, link
   `spec` to the main checkout's `spec/`. Do not copy or commit it.
5. Work with TDD for bug fixes: explore the failing sample or behavior, add a
   focused failing test, make it pass, and refactor only where it reduces real
   complexity.
6. Verify with the narrowest meaningful tests first, then broader checks when
   the touched surface warrants it.
7. Commit only the intended source/docs/test files. Do not commit private
   samples or generated local artifacts.
8. Push the feature branch and open the PR as a draft. Take it through the
   merge gate below, then merge with a merge commit. Do not squash.

## Worktrees and Parallel Sessions

- Before editing, make sure `spec` exists in the worktree and read the
  `AGENTS.md` of each package you will edit.
- A task scoped to one format edits that format's package. Reading other
  packages is fine. Edit `packages/core`, `packages/ooxml-common`, or another
  format only for a coherent fix described in Cross-Package Integration, or
  when the user asks for a cross-package change.
- Treat unexpected dirty files as work that belongs to the user or to another
  session. Do not revert or overwrite them; preserve them as in step 2.

## Git, PRs and Issues

- Never push directly to `main` or to a release branch. Use PRs.
- Open every PR with `gh pr create --draft`. Keep it draft, including through
  rebases, until the adversarial review has passed and the VRT is done.
- Merge gate, one PR at a time in dependency order: adversarial review rounds
  until PASS → one scoped VRT against the latest `main` (`vrt` skill) →
  `gh pr ready` → CI green on the current head → merge with
  `gh pr merge <n> --merge --match-head-commit <full sha>`. Copy SHAs from
  `gh pr view <n> --json headRefOid` output; never retype or expand one.
- Do not open PRs faster than they can land. Retarget and rebase the next
  stacked PR after each merge.
- A fix that makes some documents better and others worse is not merged by
  default; bring per-file evidence to the user. Layout time and bundle size
  are hard constraints, not prices for small fidelity gains.
- Fix regressions from merged work forward: keep the behavior the original
  change was for and fix the cause. Do not revert.
- Set-aside issues that are in scope get the `deferred` label and a one-line
  comment giving the reason; propose deferral to the user first. Skip
  `deferred` issues unless the user brings one back. Do not create release
  milestones.
- Close an issue once its fix is merged, with a comment naming the merge
  commit. Leave partly fixed issues open.
- When asking the user to decide, give each option's pros and cons and a
  recommendation.
- Before `git push`, set `git config http.postBuffer 524288000` if large
  pushes are expected.
- Commit messages follow the repository style: subject `fix(scope): ...`,
  `test(scope): ...`, `refactor(scope): ...`; the body explains the root
  cause, the fix, and verification, citing specification sections or observed
  Office behavior when relevant.
- PR titles, PR descriptions, commit messages, and public docs must be safe
  for an OSS repository:
  - no local absolute paths, usernames, home directories, machine names, or
    other personal environment details;
  - do not quote or summarize the concrete contents of private or copyrighted
    samples;
  - describe private-sample-driven work by the implementation problem, the
    OOXML/Office behavior mismatch, the class of affected documents, and the
    verification performed;
  - keep local-only sample filenames and visual details out of public PR text
    unless the user explicitly approves publishing them.
- Agent-authored commits include a co-author trailer naming the model
  actually used for the task, e.g.
  `Co-Authored-By: Codex <model-name> <noreply@openai.com>` or
  `Co-Authored-By: Claude <model-name> <noreply@anthropic.com>`. When one
  agent directs, delegates to, or reviews for another, include a trailer for
  each agent.

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

### Test value and consolidation

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

## OOXML Implementation Policy

Be specification-first.

- Before implementing OOXML behavior, consult the relevant documents under the
  local `spec/` symlink. Prefer ECMA-376 / ISO-29500 for normative behavior,
  and Microsoft extension notes (`[MS-DOCX]`, `[MS-XLSX]`, `[MS-PPTX]`,
  `[MS-ODRAWXML]`) when Office-specific behavior or extensions are involved.
- A user report like "this line is too long" or "this object is missing" is a
  hint that an OOXML rule is missing or mis-modeled. Identify the related
  section and implement the rule for the whole class of documents.
- When behavior differs from Office, first check whether the parser discarded
  necessary information during parsing, inheritance, or style merging. If so,
  extend the parser/model instead of guessing in the renderer.
- Do not add heuristics only to improve one VRT/sample number: no arbitrary
  thresholds, empirical scaling constants, or special-cased sample paths.

### Fidelity goal

- The goal is rendering that is honest to the specification and reasonable:
  content correct and complete, typeset properly with the fonts actually
  available, nothing broken. Exact Office parity is not the goal; differences
  from unbridgeable causes (font availability, glyph metrics, rasterization)
  are accepted, not tracked as defects.
- Implement Office's implicit rules when they make real documents render more
  correctly and the rule is simple and cheap. A difference earns a fix only if
  the layout looks broken or misleading and the fix does not cost noticeable
  bundle size or layout time.
- Measure every case of an Office behavior, then implement the whole rule. Do
  not ship the measured part with guesses or fail-closed errors for the rest.
  A case Office itself refuses or repairs is a measured result.
- Render every supported feature. Omitting content with a warning is allowed
  only when the caller deliberately left an opt-in module out (for example
  ChartEx); any other missing feature is unimplemented work.
- Render each document the single best way. Do not add caller-selectable
  fidelity modes (`strict`/`approximate`, reading policies) or APIs that
  report omissions or approximations; a one-time `console.warn` is enough.
  Strings and UI the library invents itself (error placeholders, `parseError`)
  may be customizable.

### Compatibility rules for implicit Office behavior

When Office behavior is not fully specified, an agent may adopt a
compatibility rule without per-rule user approval when:

- it was found by varying one parameter at a time across Office-produced
  controls (`office-reference-export` skill), including boundaries and
  counterexamples, not from one private sample;
- it is the smallest rule that reproduces the behavior, not a fine-grained
  branch chasing pixel parity;
- its reach in real documents justifies its size and performance cost.

List each adopted rule with its evidence in the PR. Reject rules shaped to a
sample's character set, rules gated by many branches for synthetic-only
classes, and rules that cost kilobytes or percent for little reach.
Narrowing or removing such a rule is a normal forward fix. Mark any
unavoidable temporary heuristic in code with the missing work and track
it for removal.

### Fonts

- Never fit logic to a particular font: no font-family name comparisons and no
  per-font empirical constants. Metrics come from font data — the host face's
  own tables, or the reference-font-metrics data in `packages/core/src/fonts/`.
  A font-dependent Office behavior must be a general rule over font data,
  measured across several fonts, sizes and scripts.
- Lay out with the metrics of the face that actually paints on the viewing
  host; if the authored face is absent, use the fallback face's metrics and
  accept pagination differences. Never branch on the viewer's OS. When the
  font data is identical but Mac and Windows Office still differ, follow
  Windows.

### Packaging and resources

- Admit input by actual work only: the parser's per-entry and aggregate
  inflation accounting (`docs/ooxml-resource-governance-design.md`). Never add
  checks on compressed or file size.
- Classic charts ship by default; ChartEx is an opt-in TypeScript module. Keep
  one WASM module per format; do not split parsers into extra lazily loaded
  WASM modules.
- Legacy Office (DOC/XLS/PPT) lives only in `@silurus/ooxml-legacy-office`
  (`packages/legacy-office`), an opt-in package with an exact peer pin on
  `@silurus/ooxml`. Core and format packages gain no legacy-specific fields,
  branches, defaults or imports; any addition there must be a format-generic
  capability. Legacy readers map binaries directly to the existing models,
  never through generated OOXML.

### Source is the durable specification

Express every settled library behavior in production source. Put necessary
decision rationale, tradeoffs, evidence boundaries, and intentionally
unsupported behavior in concise comments next to the responsible
implementation. Distinguish normative specification rules from observed
Office behavior and library policy. For an observed rule, record the tested
input classes and boundary/counterexample results that explain its scope,
without private sample contents or paths.

- Code and these comments must remain understandable without local experiment
  files, handoff notes, chat history, PR descriptions, or a separate design
  document. Tests verify the contract; they must not be its only record.
- When an investigation settles a rule, put it and its essential rationale
  into the responsible source before declaring the work complete or deleting
  the evidence. Do not copy an entire experiment log into comments.
- Parameter sweeps, generated Office samples/PDFs/images, probe scripts, logs,
  comparison builds, and scratch notes are temporary investigation artifacts.
  Create them in a task-owned temporary directory outside the checkout and
  remove them once their conclusions are captured in source and focused
  regression coverage. Do not accumulate `output/` directories as an
  alternative specification store.
- Unresolved observations are not established compatibility rules. Record any
  implemented limitation beside its gate.
- Before cleanup, inspect contents, tracked status, symlinks, and active-task
  ownership. Do not delete another task's evidence, original private samples,
  specification documents, or maintained VRT references. Never follow
  scratch-directory symlinks into dependencies or other worktrees for deletion.

## Cross-Package Integration

DOCX, XLSX, and PPTX share many ECMA-376 concepts: DrawingML images,
`srcRect`, text runs, paragraph properties, themes, fills, effects, and shape
presets.

- Treat one package's bug as a signal to inspect the sibling packages.
- If the concept is shared, fix the shared layer (`packages/ooxml-common` or
  `packages/core`) where possible. Do not duplicate pure parsing predicates or
  common type definitions unless the format truly diverges.
- Avoid false abstraction. Renderer logic tightly coupled to a package's
  layout model may stay local.
- Cross-package fixes may touch `core` / `ooxml-common` and multiple packages
  in one PR when that is the coherent unit of work.

## Mandatory Adversarial Review

Before declaring an OOXML feature or bug fix complete, review the final diff
adversarially, even when focused tests and VRT pass, with a fresh reviewer
that did not author the change, and repeat after fixes until it passes. The
`adversarial-review` skill is the checklist: specification fidelity, no
unexplained heuristics, bounded inference, performance and resource safety,
module boundaries, cross-format wiring, and regression evidence; DOCX
architecture changes also run the `docx-architecture-audit` skill. Record the
conclusion and material evidence in the handoff or PR description; do not
silently waive an unresolved item. A passing VRT is not evidence that a
heuristic is correct.

## Private Samples and Storybook

Private Office samples and local-only sample stories are gitignored and must
not be committed:

- `packages/{docx,xlsx,pptx}/public/private/<fmt>/`
- `packages/legacy-office/public/private/{doc,xls,ppt}/` (with paired OOXML
  controls)
- `packages/opendocument/public/private/{odt,ods,odp}/` (with their PDFs)
- `packages/*/src/*privateDemo.stories.ts`

Generated `packages/*/src/wasm/`, VRT screenshots and diffs stay local too.
A worktree may copy private stories/samples from the main checkout, without
`.DS_Store`, Office lock files, screenshots, VRT diffs, or private
reference images.

Storybook is unified at the repository root and serves
`packages/<fmt>/public/` at `/<fmt>/`; use those prefixes when fetching
samples. In sandboxed environments its port probe may fail with `EPERM`;
running it outside the sandbox is acceptable when the user asks to view it.

## Visual Regression Tests

VRT is local-only because private samples are not redistributable. The `vrt`
skill holds the commands, environment variables, harness bootstrap, and
fidelity-score recording.

- The regression oracle is the previous renderer's output. Never use
  Word/Excel/PowerPoint or PDF images as the oracle; keep Office-reference
  fidelity evaluation as a separate run.
- The baseline for unmerged work is the latest `origin/main`; stacked PRs
  compare adjacent heads. A release compares against the previous release tag.
- A missing or mismatched baseline fails the run; it is not a pass.
- Every exact difference is a regression candidate, even when every changed
  document carries the feature being implemented. Adjudicate each changed
  page, sheet, or slide against an Office-produced reference, or get the
  user's explicit decision; report anything else as unadjudicated. Never
  approve a group of differences because the files share a feature.
- Never update files under `packages/*/tests/visual/references/` unless the
  user explicitly asks. Never loosen `RATCHET_DROP_PCT` or `FAIL_ABOVE_PCT` to
  make a run pass. Run `pnpm vrt:fidelity` with every renderer change, and
  record `scores.json` only from committed renderer code.

## Documentation and Public Examples

Public docs and commit messages are written in English, in a plain, factual
tone: descriptive headings that name the content, no slogans or emotive copy.

README, docs, and Storybook public-facing code examples:

- use `as Type` assertions instead of postfix non-null assertions;
- use `canvas` for canvas-backed viewers (`DocxViewer`, `PptxViewer`) and
  `container` for container-backed viewers (`XlsxViewer`);
- do not document private/local-only sample paths as public API.

Announcements are concise and user-facing:

- lead with the visible outcome, grouping low-level fixes by the user problem
  they solve;
- state whether migration is required; when it is, give the minimum steps and
  name any changed default or behavior;
- omit implementation detail that does not change a user's decision; put a
  useful technical qualification in a final `Technical note` section;
- keep bundle measurements on the `/bundle-size` page, and do not link API
  reference details directly from release notes.

## Releases

Minor releases come from `main`; patch releases come from the `v0.x` release
branch with only the chosen fixes backported by PR. Before tagging, no page
may move farther from Office than in the previous release without an
explanation the user approves. The `release` skill is the procedure.

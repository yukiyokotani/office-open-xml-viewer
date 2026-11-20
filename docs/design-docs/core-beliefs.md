# Core beliefs

The base policy of the library and the implementation rules that follow
from it. These are rules, not suggestions; every OOXML change is reviewed
against them ([adversarial review](../workflow/adversarial-review.md)).

## Base policy

A browser library cannot carry the fonts that Word, Excel and
PowerPoint have, so font resolution differs and the rendering can never match
Office exactly. Exact match is not the goal. Accept the differences that follow
from this, and aim for reasonable rendering whose layout does not break:
content correct and complete, typeset properly with the fonts actually
available. Implement Office's implicit rules when they make real documents
render more correctly. A difference from Office earns a fix only if the layout
looks broken or misleading and the fix is simple and cheap (for example a
compatibility flag or a rounding rule); otherwise it is an accepted
difference, with no follow-up issue.

## Specification first

Within that policy, be specification-first.

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

## Completeness

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

## Compatibility rules for implicit Office behavior

When Office behavior is not fully specified, an agent may adopt a
compatibility rule without per-rule user approval when:

- it was found by varying one parameter at a time across Office-produced
  controls ([Office reference export](../workflow/office-reference-export.md)), including boundaries and
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

## Fonts

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

## Packaging and resources

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

## Source is the durable specification

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

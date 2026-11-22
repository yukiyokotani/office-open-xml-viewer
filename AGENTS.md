# Agent Instructions

This file is the entry point for every AI coding agent and tool working in
this repository. It is a map, not the rulebook: it states the few invariants
that always apply and tells you where to read next. The knowledge itself lives
in `docs/`, which is the system of record.

## Project

`office-open-xml-viewer` renders OOXML documents (`docx`, `xlsx`, `pptx`) to
browser Canvas: a Rust/WASM parser layer plus TypeScript Canvas renderers and
viewers.

- `packages/core/`, `packages/ooxml-common/` — shared TypeScript and Rust layers.
- `packages/{docx,xlsx,pptx}/` — per-format parser, renderer, viewer, tests.
- `packages/legacy-office/`, `packages/opendocument/` — opt-in companion
  packages for DOC/XLS/PPT and ODF.
- `site/` — public site. `.storybook/` — unified Storybook config.

## Invariants

1. Knowledge that is not in this repository does not exist for other agents.
   Express a decision in source code first, then in a comment beside the
   source, then in the issue or PR. Do not add design documents, plans or
   audits; `docs/` holds procedures, policy and user guides only. A change
   that makes a documented statement false corrects or deletes it in the same
   PR ([documentation](docs/workflow/documentation.md)).
2. A browser library cannot carry Office's fonts, so its rendering never
   matches Office exactly, and exact match is not the goal. Aim for
   reasonable rendering whose layout does not break
   ([core beliefs](docs/design-docs/core-beliefs.md)).
3. Be specification-first: consult the local `spec/` symlink (ECMA-376 /
   ISO-29500, then the Microsoft extension notes) before implementing OOXML
   behavior. No heuristics fitted to one sample or one font.
4. Never push to `main` or a release branch. Open PRs as drafts, take them
   through the merge gate, and merge with a merge commit, never a squash
   ([Git, PRs and issues](docs/workflow/git-and-prs.md)).
5. Never commit private samples or generated local artifacts, and keep local paths,
   usernames and private sample contents out of commits, PRs and public docs.
6. Never update the reference images in
   `packages/*/tests/visual/references/` or loosen VRT thresholds unless the
   user explicitly asks. The VRT oracle is the previous renderer
   ([VRT](docs/workflow/vrt.md)).
7. Review every OOXML feature or fix adversarially before calling it complete
   ([adversarial review](docs/workflow/adversarial-review.md)).
8. Other sessions work in parallel worktrees. Treat unexpected dirty files as
   theirs: preserve them, never revert them.
9. Agent-authored commits carry a `Co-Authored-By:` trailer for each agent
   involved, naming the model actually used.

## Where to read next

| When you are… | Read |
| --- | --- |
| implementing or reviewing OOXML behavior | [`docs/design-docs/core-beliefs.md`](docs/design-docs/core-beliefs.md) |
| starting a branch, opening or merging a PR, triaging issues | [`docs/workflow/git-and-prs.md`](docs/workflow/git-and-prs.md) |
| choosing or writing tests | [`docs/workflow/testing.md`](docs/workflow/testing.md) |
| changing anything that renders pixels | [`docs/workflow/vrt.md`](docs/workflow/vrt.md) |
| finishing a feature or fix | [`docs/workflow/adversarial-review.md`](docs/workflow/adversarial-review.md) |
| changing DOCX layout architecture | [`docs/workflow/docx-architecture-audit.md`](docs/workflow/docx-architecture-audit.md) |
| needing Office ground truth (controls, PDF exports) | [`docs/workflow/office-reference-export.md`](docs/workflow/office-reference-export.md) |
| recording a decision, or writing, moving or deleting docs | [`docs/workflow/documentation.md`](docs/workflow/documentation.md) |
| preparing a release or patch release | [`docs/workflow/release.md`](docs/workflow/release.md) |
| looking for design contracts and guides | [`docs/README.md`](docs/README.md) (index of all docs) |
| editing one package | that package's `AGENTS.md` |

## How the knowledge is organized

- `AGENTS.md` (this file) and `packages/<name>/AGENTS.md` — short maps. A
  package file adds package-specific facts and does not repeat this one.
- Source code and the comments beside it — every design decision and its
  rationale. There are no design documents.
- `docs/design-docs/core-beliefs.md` — the policy every change is reviewed
  against.
- `docs/workflow/` — repeated procedures and the rules that govern them.
- `docs/*.md` — guides for library users that a README, the site or shipped
  code links to.
- `docs/README.md` — the index; every document in `docs/` is listed there.
- `.agents/skills/<name>/SKILL.md` — thin pointers that load a
  `docs/workflow/` page on demand. Codex reads `.agents/skills/`; Claude Code
  reads each through the relative symlink `.claude/skills/<name>` ->
  `../../.agents/skills/<name>`. Skills hold no content of their own. On
  Windows the links need `git config core.symlinks true` before checkout.
- `.claude/CLAUDE.md` and `packages/<name>/CLAUDE.md` — one-line imports of
  the matching `AGENTS.md` for Claude Code.
- `ast-grep/` — lint rules (`rules/`), shared matchers (`utils/`) and rule
  tests (`tests/`), configured by `sgconfig.yml`.

Do not add tool-specific rule files. `scripts/check-agent-instructions.test.mjs`
(part of `pnpm test`) checks the limits of every `AGENTS.md`, the index, links,
named paths and scripts, the allowed document locations and the skills;
[documentation](docs/workflow/documentation.md) lists the checks and the
gardening pass that covers what they cannot.

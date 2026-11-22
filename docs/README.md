# Documentation

This directory holds three kinds of document and nothing else. Design decisions
are not documented here: they are expressed in source code and in the comments
beside it ([`workflow/documentation.md`](workflow/documentation.md)).

Documents are maintained in English. Parallel translations are not kept
because they drift from the implementation.

## Guides for library users

The root [`README.md`](../README.md), the package READMEs and the
[site](https://ooxml.silurus.dev/) are the main user documentation. A guide
stays here only while one of them, or a message printed by shipped code, links
to it.

- [`legacy-office-conversion.md`](legacy-office-conversion.md) — the
  experimental `@silurus/ooxml-legacy-office` companion (DOC, XLS, PPT).
- [`opendocument.md`](opendocument.md) — the experimental
  `@silurus/ooxml-opendocument` companion.
- [`migration-0.90.md`](migration-0.90.md) — font sources (`fonts`) replace
  `useGoogleFonts`. The 0.90 deprecation warning prints this path.

## Core beliefs

- [`design-docs/core-beliefs.md`](design-docs/core-beliefs.md) — the base
  fidelity policy and the OOXML implementation rules that follow from it.

## Workflow

Procedures and the rules that govern them. The root
[`AGENTS.md`](../AGENTS.md) is the short entry point that maps into these
pages; the skills in `.agents/skills/` only point here.

- [`workflow/git-and-prs.md`](workflow/git-and-prs.md) — development flow,
  worktrees, draft PRs and the merge gate, issues, commit and PR text.
- [`workflow/testing.md`](workflow/testing.md) — verification commands and
  which tests are worth keeping.
- [`workflow/vrt.md`](workflow/vrt.md) — visual regression rules and runbook.
- [`workflow/adversarial-review.md`](workflow/adversarial-review.md) — the
  mandatory review before a change is complete.
- [`workflow/docx-architecture-audit.md`](workflow/docx-architecture-audit.md)
  — the review procedure for DOCX layout changes.
- [`workflow/office-reference-export.md`](workflow/office-reference-export.md)
  — Office controls, reference PDFs and driving Office.
- [`workflow/documentation.md`](workflow/documentation.md) — where knowledge
  is recorded, what this directory may contain, how documents are kept
  current, and the writing style.
- [`workflow/release.md`](workflow/release.md) — minor and patch releases.

`images/` holds the screenshots and icon used by the READMEs and the site.
[`../scripts/README.md`](../scripts/README.md) lists the repository scripts.

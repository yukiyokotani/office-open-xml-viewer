# Documentation

Where knowledge is recorded, what `docs/` may contain, how documents are kept
current, and the writing style.

## Where knowledge goes

Record a decision in the first place on this list that can hold it:

1. **Source code.** Types, names, module boundaries, tests and lint rules
   (`ast-grep/rules/`, `scripts/check-*.mjs`) state the decision and enforce it.
2. **A comment beside the responsible source**, for the rationale, evidence
   boundary or deliberate omission that code cannot express
   ([core beliefs](../design-docs/core-beliefs.md), "Source is the durable
   specification").
3. **The issue or PR**, for history: the investigation, the alternatives that
   were rejected, the delivery plan, review notes and measurements taken once.

A standalone design document, plan, audit, evidence log or status matrix is not
on the list. Do not add one to the repository. Plans and trackers are GitHub
issues.

## What `docs/` contains

- `docs/workflow/` — procedures agents and contributors follow, with the rules
  that govern them. A page exists only for a procedure that is repeated.
- `docs/design-docs/core-beliefs.md` — the policy every change is reviewed
  against. It is the only file in that directory.
- `docs/images/` — the screenshots and icon the READMEs and the site use.
- `docs/*.md` — guides for library users that a consumer-facing surface links
  to: the root or a package `README.md`, the site, or a message printed by
  shipped code. A guide that repeats a site page is merged into the site page
  and deleted.

Outside `docs/`, Markdown is limited to `README.md`, `AGENTS.md`, `CLAUDE.md`,
`SKILL.md`, `CHANGELOG.md`, `CODE_OF_CONDUCT.md` and `THIRD_PARTY_NOTICES.md`.
A directory `README.md` says what the directory is and how to run it; usage of
one script belongs in that script's header.

## Keeping documents current

- A change that makes a statement in `docs/`, a `README.md` or an `AGENTS.md`
  false corrects or deletes that statement in the same PR.
- Delete instead of archiving. Git history and the PR keep the old text.
- A migration guide is deleted, together with its entry in `docs/README.md`,
  once no consumer-facing surface on `main` links to it and no supported
  release prints a URL that resolves to its path on `main`. Until then it
  keeps its path. A relative link inside a release branch resolves
  within that branch and does not keep the file on `main`. A guide kept only
  for a supported release is listed, with the reason, in
  `RELEASE_RETAINED_GUIDES` in `scripts/check-agent-instructions.test.mjs`.
- Do not cite an issue as open work unless it is open. When the issue closes,
  remove the sentence or state the outcome.
- Do not record dated status, progress checklists, case counts or one-time
  measurements in a document; they are stale after the next change. The
  standing checklist of a repeated procedure is part of that procedure.

`scripts/check-agent-instructions.test.mjs` (part of `pnpm test`) enforces what
can be checked mechanically:

- every `AGENTS.md` chain fits Codex's 32 KiB limit and stays a short map;
- every page under `docs/` is indexed and every relative link resolves;
- `docs/` holds only the pages above and the PNG images under
  `docs/images/`, and a top-level `docs/*.md` guide is a link destination in
  a consumer-facing surface or is retained for a supported release, so a
  guide nothing links to any more fails the check and is deleted;
- no Markdown file exists outside the names listed above;
- repository paths and `pnpm` scripts named in `AGENTS.md`, `docs/workflow/`,
  `docs/design-docs/` and the skills exist;
- no task-list checkbox appears in `docs/`;
- every skill is a pointer to its `docs/workflow/` page.

### Gardening pass

The mechanical check cannot tell whether a sentence is still true. Run this
pass before each minor release, and whenever asked to review the documentation:

1. Run `node --test scripts/check-agent-instructions.test.mjs` and fix what it
   reports.
2. For each page under `docs/` and each `AGENTS.md`, verify its concrete claims
   against the code: API and option names, defaults and limits, commands,
   described behavior. Correct or delete what is false.
3. List the issues and PRs each page cites (`gh issue view <n> --json state`)
   and apply the rule above to the closed ones.
4. Delete every guide that meets the deletion rule, and every sentence that
   repeats what the code, a script header or the site already states.
5. Look for knowledge in the wrong place: rationale in a document that belongs
   in a source comment, or a rule followed in practice that no file records.
   Move the first; write down the second.
6. Open one PR with the result. List what was deleted and where each piece of
   surviving knowledge now lives.

## Style

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

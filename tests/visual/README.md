# Visual regression harness

`AGENTS.md` ("Visual Regression Tests") states the rules: the oracle is the
previous renderer, the baseline is bound to a checkout, and every difference
must be adjudicated. This file holds the procedures that go with them.

## Harness bootstrap exceptions

When a harness change is bootstrapped against an older baseline
(`VRT_ALLOW_HARNESS_CHANGES=1`, with the allowlist in `private-corpus.mjs`),
two more things are allowed:

- `pnpm-lock.yaml`, but only hunks that add or remove the fixtures'
  `@fontsource/*` packages. Any other hunk fails.
- The inputs under `packages/<fmt>/public/corpus/`. Git ignores them before
  the corpus is tracked, so the bootstrap may copy them in.

## New public specimens

A specimen added in the PR under test does not exist at the baseline
revision, so the previous renderer has no image of it. With
`VRT_ALLOW_NEW_CORPUS_INPUTS=1` for both the snapshot and the comparison run,
the new `packages/<fmt>/public/corpus/*.{docx,xlsx,pptx}` files may be copied
into the baseline worktree as untracked files. Nothing else is excused: a
change to a tracked input, any other file in the corpus folders, and any
renderer change still fail, and the manifest's SHA-256 binding rejects a
copied input that differs from the candidate's.

## Fast private tier (PPTX/XLSX)

The fast tier renders only the slides and sheets that a greedy cover of
structural XML features needs. The features are element nesting, attributes,
enum-like values, relationship types and media types, including parts reached
through relationships (`corpus-features.mjs`).

Generate the gitignored selection manifests, then use the same baseline
binding as the full run:

```bash
pnpm vrt:private:fast:select
# baseline worktree
VRT_BASELINE_REVISION=$base VRT_PORT=5190 pnpm vrt:private:fast:snapshot
# candidate
VRT_BASELINE_REVISION=$base VRT_BASELINE_CHECKOUT=/tmp/ooxml-vrt-baseline VRT_PORT=5191 pnpm vrt:private:fast
```

Both worktrees need the same selection manifest. It lives in the gitignored
private corpus, so copy it across.

The run fails closed in these cases:
- the selection manifest is missing;
- a corpus file was added, removed or changed;
- the baseline was captured under a different selection.

Fast-tier baselines are stored separately from full-tier ones. Structural
features do not capture geometry-dependent behaviour such as text wrapping,
autofit, pagination or value formatting. Run the full `vrt:private` before a
release, or whenever a change affects layout geometry. DOCX has no fast tier
yet.

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

## Added or changed public specimens

A specimen that the PR under test adds or rewrites has no image from the
previous renderer. Copy the candidate's specimen files into the baseline
worktree, so both renderers draw identical bytes, and set
`VRT_ALLOW_CORPUS_INPUT_CHANGES=1` for both the snapshot and the comparison
run:

```bash
git diff --name-only --diff-filter=AM "$base" -- 'packages/*/public/corpus/*.docx' \
  'packages/*/public/corpus/*.xlsx' 'packages/*/public/corpus/*.pptx' \
  | while read -r f; do cp "$f" "/tmp/ooxml-vrt-baseline/$f"; done
```

Only a `.docx`/`.xlsx`/`.pptx` directly in its own package's corpus folder
qualifies. It must be a regular file reached without a symlink, and during the
comparison each copy must still equal the candidate's file; the manifest's
SHA-256 binding checks the inputs as well. Any other file, a deleted input and
any renderer change still fail. The specs serve the fixture the input bytes
the corpus manifest recorded (`selfVrtInputBody`), so a specimen replaced
during capture fails instead of lending its image to another input. When the
baseline revision predates this allowance, also copy
`tests/visual/private-corpus.mjs`, `packages/*/tests/visual/visual.spec.ts`
and `packages/*/tests/visual/fixture.html` (the specs and fixtures agree on
the encoded input URL) and set `VRT_ALLOW_HARNESS_CHANGES=1` for both runs (harness bootstrap above).

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

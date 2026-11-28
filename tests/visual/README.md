# Visual regression harness

`AGENTS.md` ("Visual Regression Tests") states the rules: the oracle is the
previous renderer, the baseline is bound to a checkout, and every difference
must be adjudicated. This file holds the procedures that go with them.

## Threat model

The self-VRT bindings protect against stale or mismatched files and operator
mistakes: the wrong checkout, a forgotten specimen copy, leftover images,
gitignored files, symlinked generated directories, and non-regular files such
as named pipes. They do not defend against a local process that deliberately
replaces files while a run is in progress; concurrent races are out of scope.

## Harness bootstrap exceptions

When a harness change is bootstrapped against an older baseline
(`VRT_ALLOW_HARNESS_CHANGES=1`, with the allowlist in `private-corpus.mjs`),
two more things are allowed:

- `pnpm-lock.yaml`, but only hunks that add or remove the fixtures'
  `@fontsource/*` packages. Any other hunk fails.
- The inputs under `packages/<fmt>/public/corpus/`. Git ignores them before
  the corpus is tracked, so the bootstrap may copy them in. The harness lists
  these folders from disk and compares them with the files tracked at the
  baseline revision, so an ignored copy is checked like any other.

## Added or changed public specimens

A specimen that the PR under test adds or rewrites has no image from the
previous renderer. Copy the candidate's specimen files into the baseline
worktree, so both renderers draw identical bytes, and set
`VRT_ALLOW_CORPUS_INPUT_CHANGES=1` for both the snapshot and the comparison
run:

```bash
node tests/visual/copy-corpus-inputs.mjs /tmp/ooxml-vrt-baseline "$base"
```

The script (built-ins only, run from the candidate checkout) lists the added
or modified specimens against `$base`, creates a missing corpus folder, and
refuses to write through a symlink in the destination.

Only a `.docx`/`.xlsx`/`.pptx` directly in its own package's corpus folder
qualifies. It must be a regular file reached without a symlink, and during the
comparison each copy must still equal the candidate's file. The snapshot records the
copies (path and SHA-256) in the corpus manifest; the comparison fails when a
recorded copy is missing or differs from the record or the candidate, or when
the flag is not set for it. The manifest's SHA-256 binding checks the inputs
as well. Any other file, a deleted input and
any renderer change still fail. The specs serve the fixture the input bytes
the corpus manifest recorded (`selfVrtInputBody`), so a specimen replaced
during capture fails instead of lending its image to another input. When the
baseline revision predates this allowance, also copy
`tests/visual/private-corpus.mjs`, `tests/visual/checked-read.mjs`,
`tests/visual/corpus-features.mjs`, `tests/visual/copy-corpus-inputs.mjs`, `packages/*/tests/visual/visual.spec.ts`
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

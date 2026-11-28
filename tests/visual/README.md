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

The comment at the top of `private-corpus.mjs` states what a passing
comparison establishes. In short:

- **Same input.** A fixture never loads a corpus input by file name. The specs
  hand it a harness URL (`/__selfvrt/<corpus>/<index>.<ext>`), and the harness
  answers that request itself: one checked read, required to have the SHA-256
  in the corpus manifest, whose bytes are the ones sent. Each document's
  manifest records that digest beside its image digests, and the comparison
  requires it to equal the candidate's.
- **Previous renderer.** Baseline images are read only from a Git checkout
  root at `VRT_BASELINE_REVISION` that is not the candidate and differs from
  the revision only by what the two `VRT_ALLOW_*` flags admit.
- **Same images.** A baseline document holds exactly the items its manifest
  lists, each with its recorded SHA-256, reached without a symlink.

Inputs are bound by content, so it does not matter whether Git tracks, ignores
or reports a file: a missing, renamed or different input changes the input
list or a digest and fails the comparison. Output is written, and stale output
removed, only below the checkout and never through a symlink; documents the
corpus no longer lists lose their generated files at the next snapshot
(baseline) or comparison (candidate screenshots).

## Harness bootstrap exceptions

When a harness change is bootstrapped against an older baseline
(`VRT_ALLOW_HARNESS_CHANGES=1`, with the allowlist in `private-corpus.mjs`),
two more things are allowed:

- `pnpm-lock.yaml`, but only hunks that add or remove the fixtures'
  `@fontsource/*` packages. Any other hunk fails.
- The inputs under `packages/<fmt>/public/corpus/`. Git ignores them before
  the corpus is tracked, so the bootstrap may copy them in. The manifests
  bind them by SHA-256 like every other input.

## Added or changed public specimens

A specimen that the PR under test adds or rewrites has no image from the
previous renderer. Copy the candidate's specimen files into the baseline
worktree, so both renderers draw identical bytes, and set
`VRT_ALLOW_CORPUS_INPUT_CHANGES=1` for both the snapshot and the comparison
run:

```bash
node tests/visual/copy-corpus-inputs.mjs /tmp/ooxml-vrt-baseline "$base"
```

The command (built-ins only, run from the root of the candidate checkout)
requires the destination to be the root of another Git checkout whose HEAD is
`$base`. It copies the tracked specimens that the candidate's working tree
adds or modifies relative to `$base`, creates a missing corpus folder, and
refuses to write through a symlink. When the candidate deletes or renames a
specimen it stops before writing anything: the two sides would render
different input sets, which the self-VRT cannot compare.

The flag admits only a `.docx`/`.xlsx`/`.pptx` present directly in its own
package's corpus folder. Any other file, a specimen deleted from the baseline
and any renderer change still fail. The copies need no check of their own:
the manifests bind every input by SHA-256 (threat model above), so a
forgotten, renamed or stale copy fails the comparison. When the baseline
revision predates this harness, also copy `tests/visual/private-corpus.mjs`,
`tests/visual/checked-read.mjs`, `tests/visual/corpus-features.mjs` and
`packages/*/tests/visual/visual.spec.ts` and set
`VRT_ALLOW_HARNESS_CHANGES=1` for both runs (harness bootstrap above).

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

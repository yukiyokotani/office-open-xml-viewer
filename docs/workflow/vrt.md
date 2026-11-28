# Visual regression tests

## Rules

VRT is local-only because private samples are not redistributable.

- The regression oracle is the previous renderer's output. Never use
  Word/Excel/PowerPoint or PDF images as the oracle; keep Office-reference
  fidelity evaluation as a separate run.
- The baseline is the previous renderer the change will land on; see
  [Choosing the baseline](#choosing-the-baseline) for unmerged work, backports,
  stacked PRs and releases.
- A missing or mismatched baseline fails the run; it is not a pass.
- Every exact difference is a regression candidate, even when every changed
  document carries the feature being implemented. Adjudicate each changed
  page, sheet, or slide against an Office-produced reference, or get the
  user's explicit decision; report anything else as unadjudicated. Never
  approve a group of differences because the files share a feature.
- Never update the reference images under
  `packages/*/tests/visual/references/` unless the user explicitly asks. The
  `scores.json` files beside them follow the recording rule below. Never loosen `RATCHET_DROP_PCT` or `FAIL_ABOVE_PCT` to
  make a run pass. Run `pnpm vrt:fidelity` with every renderer change, and
  record `scores.json` only from committed renderer code.

## Corpora

There are two self-VRT corpora per package. Both render every file completely
and compare exact decoded pixels with the previous renderer's images:

- **demo**: every file under `packages/<fmt>/public/demo/` (tracked). Captured
  by `pnpm vrt:snapshot` and checked by `pnpm vrt`.
- **private**: every file directly in `packages/<fmt>/public/private/<fmt>/`
  (gitignored). Subdirectories are not read: a control that should take part
  in the self-VRT is placed at the top level of that directory. Captured by `pnpm vrt:private:snapshot` and checked by
  `pnpm vrt:private`.

`pnpm vrt` also runs the package's other visual specs (worker parity, overlays,
acceptance checks). Specs that need a private sample load it from
`public/private/<fmt>/` and skip when it is absent.

## Choosing the baseline

- Work that is not merged yet: the head of the PR's base branch. That is the
  latest `origin/main`, or the release branch (`origin/v0.x`) for a backport.
  Record its full SHA. Differences between the last release and `main` were
  adjudicated when `main` moved, so they are not this work's concern.
- Stacked PRs: compare adjacent heads (main → PR1, PR1 → PR2), so every diff
  maps to exactly one PR.
- A release: the previous release tag (see [release](release.md)).
- Use another baseline only when the user asks for it.

## Baseline and candidate runs

Baselines are produced from a clean, detached worktree at the baseline
revision. Snapshot mode rejects a dirty tracked checkout and writes a manifest
binding the baseline to `VRT_BASELINE_REVISION`, the SHA-256 of each input,
the exact page/sheet/slide set and the browser locale every package
Playwright config pins (`VRT_LOCALE`, default `en-US`; number separators and
built-in date formats follow it). The candidate run verifies that manifest and
reads the images from the baseline worktree named by `VRT_BASELINE_CHECKOUT`.
A baseline revision whose own harness predates the locale field writes none:
it is an `en-US` capture (leave `VRT_LOCALE` unset for it) and compares with
an `en-US` candidate run only, with no bootstrap. The field may be missing for
no other reason.
Use distinct ports so a candidate server can never satisfy the baseline run:

```bash
base=$(git rev-parse origin/main)   # the PR's base branch; origin/v0.x for a backport
git worktree add --detach /tmp/ooxml-vrt-baseline "$base"
cd /tmp/ooxml-vrt-baseline
pnpm install --frozen-lockfile && pnpm build:wasm
# The private corpus is git-ignored, so a new worktree has none: copy the same
# files into each packages/<fmt>/public/private/<fmt>/ here and in the candidate.
VRT_BASELINE_REVISION=$base VRT_PORT=5190 pnpm vrt:snapshot
VRT_BASELINE_REVISION=$base VRT_PORT=5190 pnpm vrt:private:snapshot

cd <candidate-worktree>
pnpm build:wasm
VRT_BASELINE_REVISION=$base VRT_BASELINE_CHECKOUT=/tmp/ooxml-vrt-baseline VRT_PORT=5191 pnpm vrt
VRT_BASELINE_REVISION=$base VRT_BASELINE_CHECKOUT=/tmp/ooxml-vrt-baseline VRT_PORT=5191 pnpm vrt:private
```

The root scripts run every package without stopping at the first failing one,
and still exit non-zero if any package failed. Filter a single package with
`pnpm --filter @silurus/ooxml-<fmt> vrt` (and the matching snapshot script).

Remove the baseline worktree when done.

### What the harness enforces

A comparison run requires `VRT_BASELINE_CHECKOUT` and fails closed unless it
names the root of a Git checkout whose HEAD is exactly `VRT_BASELINE_REVISION`
and whose renderer is unmodified: no tracked change, and no untracked file
outside dependencies and the private corpus. The checkout is re-verified
before every item is read, so a baseline that changes mid-run fails. Snapshot
manifests record the checkout's real path and each image's SHA-256; the
comparison requires both to match, and rejects a symlink on the path of any
baseline file it opens, so copied, replaced or redirected images are never accepted. The
harness runs Git with every `GIT_*` variable removed. A missing or mismatched
baseline fails the run; it is not a successful regression check.

## Scoping a PR's VRT

Run the VRT once per PR, after the adversarial review has passed: every review
fix would invalidate an earlier run.

- Scope it to the formats and files the change can affect: the files named in
  the diff or the findings, the files that exercise the changed code path
  (grep the corpus for the feature), and pages that differed before. For a
  change in `packages/core`, scope it to the samples that exercise that module.
- The full private corpus runs once per release, against the previous tag.
- If the PR's base branch moves after the VRT and touches the PR's files,
  merge the base branch and rerun only the VRT, not the review. The base is
  `main`, or the release branch for a backport; never merge `main` into a
  backport. Merges that touch none of the PR's
  rendering area keep a format-scoped VRT valid.

## Adjudicating differences

Enumerate every changed page, sheet or slide. For each one, compare the
candidate and the previous renderer with the Office-produced PDF, and classify
it as closer to Office, farther, or equal. A difference no Office reference or
user decision covers is reported as unadjudicated, never as a pass.

## Harness bootstrap

For the one-time bootstrap of a harness change against a baseline revision that
predates it, copy only the allowlisted VRT harness files listed in
`tests/visual/private-corpus.mjs` into the baseline worktree and set
`VRT_ALLOW_HARNESS_CHANGES=1` for both the snapshot and the comparison run.
Renderer and parser changes remain forbidden.

## Fidelity scores

`pnpm vrt:fidelity` compares each tracked demo file with its committed images
under `packages/<fmt>/tests/visual/references/demo/` and applies the fidelity
ratchet: a page, sheet or slide fails when its match-% drops more than
`RATCHET_DROP_PCT` below the score recorded in `scores.json` beside the images.
The demo references are adopted renderer outputs, not Office exports, so the
ratchet detects drift from the adopted generation. Office PDFs remain the
reference for judging fidelity, under the base policy in the [core beliefs](../design-docs/core-beliefs.md): exact
match is not the goal, and differences that leave the layout unbroken are
accepted.

Run `pnpm vrt:fidelity` alongside the self-VRT for every renderer change, so
`main` always passes its own ratchet. Record scores only from a clean checkout
of committed renderer code: the latest `origin/main` to resynchronize, or the
head of a renderer PR whose score drop has been adjudicated against the Office
PDF (recorded in that same PR). Never record from uncommitted work:

```bash
UPDATE_SCORES=1 pnpm vrt:fidelity
```

This rewrites only `scores.json`, never the reference images. Commit the
re-recorded scores as their own change that names the source revision.

Only when the user explicitly asks:

```bash
UPDATE_REFS=1 pnpm vrt:fidelity
```

## Running alongside other sessions

- Rebuild WASM in the worktree under test before the run; stale
  `packages/*/src/wasm/` output can both hide and fabricate differences.
- Do not run a production build in a worktree while a VRT script is
  rebuilding its WASM; they race on `src/wasm/*`.
- Other sessions run the same commands. Stop only your own VRT, Playwright or
  Vite processes, selected by working directory (`lsof -a -p <pid> -d cwd`),
  never by a generic `pkill -f` pattern.
- VRT screenshots and diffs stay local.

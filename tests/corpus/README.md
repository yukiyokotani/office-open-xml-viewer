# Public specimen corpus

Self-authored OOXML specimens that anyone may rebuild, inspect and use to check
a renderer change. Each specimen exercises one feature family, states on its
own page what to verify, and is built from hand-written XML so that every byte
is accounted for.

The private samples under `packages/*/public/private/` are not redistributable
and play no part here: no text, data, structure or measurement in this corpus
is derived from them.

## Layout

```
tests/corpus/
  build.mjs            rebuild every case into packages/<fmt>/public/corpus/
  gap.mjs              coverage gap against a local private corpus (stdout only)
  corpus.test.mjs      node:test suite (part of `pnpm test`)
  cases/<fmt>-NNN-<slug>.mjs   one module per specimen
  lib/grid.mjs         the grid: every length used by any case
  lib/docx.mjs, xlsx.mjs, pptx.mjs   per-format composers
  lib/opc.mjs, xml.mjs, zip.mjs      packaging
  lib/png.mjs          deterministic PNG pictures drawn by the builder
  lib/chart.mjs, drawing.mjs, iso216.mjs   shared chart, picture and data helpers
  lib/features.mjs     `covers` vocabulary (extraction shared with tests/visual/corpus-features.mjs)
  lib/package-check.mjs  content types, relationships, XML well-formedness
  render.mjs           render specimens to PNG and contact sheets for review
  references.mjs       status and recording of the Office reference PDFs
  references.json      source and PDF hashes of every reference PDF
packages/<fmt>/public/corpus/<fmt>-NNN-<slug>.<ext>   committed output
```

Storybook and the VRT fixtures serve `packages/<fmt>/public/` at `/<fmt>/`, so
a specimen is fetched as `/<fmt>/corpus/<file>`.

## Commands

```bash
pnpm corpus:build                         # write every specimen
node tests/corpus/build.mjs --check       # exit 1 if a committed file is stale
node --test tests/corpus/corpus.test.mjs  # rebuild in memory and verify
pnpm corpus:gap [--format docx] [--part 'word/document'] [--top N]
pnpm corpus:coverage --format docx [--dir SUBDIR|.] [--top N]
node tests/corpus/render.mjs --format docx [--cases docx-024,docx-031] --out DIR
```

`corpus:gap` compares the public corpus with a local private corpus under
`packages/<fmt>/public/private/<fmt>/` (top level), if one exists. It reduces
both to structural schema vocabulary only (element nesting, attribute names,
relationship types, binary part extensions; every `attr=value` feature is
dropped), prints covered / total per format and the uncovered features
grouped by part and element, ranked by how many private files use them. It
prints no file names and writes nothing. Without a private corpus it says so
and exits 0. Use it to choose the next specimens.

`corpus:coverage` measures the renderer instead of the XML. It renders every
public specimen and every private file (recursively; `--dir .` for the top
level only) through the package's VRT fixture in Chromium with V8 block
coverage, maps the served TypeScript back to `packages/*/src` lines, and lists
the lines some private file executes but no specimen does, grouped into hunks
and ranked by the number of private files reaching them. This finds behaviour
that structural vocabulary misses: combinations (a table inside balanced
columns), compatibility modes and value-dependent branches. Font-resolution
modules are excluded by default (`--exclude REGEX` to change), because the
public corpus loads its own two families while private files fall back to
host fonts. Parsing runs in a worker and is not measured; use `corpus:gap` for
it. Lines executed by the same code path in two different layout states (one
balanced page against several) look the same, so pair it with reading the
private files. Stdout shows source lines and counts only; the full report,
which names private files, is written to the ignored
`packages/<fmt>/public/private/coverage-gap.json`.

The test fails when a committed file differs by one byte from its rebuild,
when a package is malformed, when a feature a case declares in `covers` is
missing from the package, or when a page margin, indent, tab stop, column,
slide offset or extent, text inset or row height leaves the grid.

## Reproducibility

Packages are zipped by `lib/zip.mjs`: entries sorted by name, every entry
stored (no compression, so output does not depend on the zlib build), every
timestamp 1980-01-01 00:00, no extra fields. The same sources therefore give
the same bytes on any platform and Node version. Builders must not read the
clock, the environment or the locale.

## Office reference PDFs

Each DOCX and PPTX specimen, and each XLSX specimen that holds a chart, is
kept with the PDFs Office exports from it, next to it in
`packages/<fmt>/public/corpus/`: `<stem>.win.pdf` from Office for Windows and
`<stem>.mac.pdf` from Office for Mac. `tests/corpus/references.json` records
for each PDF the SHA-256 of the specimen it was exported from and of the PDF.
The test fails when a present PDF is unrecorded, differs from its record, or
was exported from a specimen that has since changed.

```bash
node tests/corpus/references.mjs                 # list missing and stale PDFs
node tests/corpus/references.mjs --record [ids]  # record freshly exported PDFs
```

The exporting machine must have the corpus fonts installed (Noto Sans, Noto
Sans JP, Noto Sans Arabic, Noto Sans Hebrew), except for the Office font
specimen, which tests substitution by design.

## Authoring

The rules every specimen follows (grid, type, font, colour, text, minimal
XML), the procedure for adding one and the review loop live in the
`corpus-specimen` skill, `.agents/skills/corpus-specimen/SKILL.md` (shared by Codex; `.claude/skills/corpus-specimen` links to it). Read it
before adding or changing a case.

## Visual regression

Each package's `tests/visual/visual.spec.ts` has a `public corpus self
regression` group, enumerated from `public/corpus/`, that renders every page,
slide or sheet through the existing VRT fixture and compares decoded pixels
exactly with a baseline captured from the previous renderer. It needs no
private data, so anyone can run it:

```bash
base=$(git merge-base HEAD origin/main)
git worktree add --detach /tmp/ooxml-vrt-baseline "$base"            # previous renderer
(cd /tmp/ooxml-vrt-baseline && pnpm install --frozen-lockfile && pnpm build:wasm \
  && VRT_BASELINE_REVISION=$base VRT_PORT=5190 pnpm vrt:corpus:snapshot)
pnpm build:wasm                                                      # changed renderer
VRT_BASELINE_REVISION=$base VRT_BASELINE_CHECKOUT=/tmp/ooxml-vrt-baseline VRT_PORT=5191 pnpm vrt:corpus
```

The comparison reads the baseline images from the baseline worktree itself,
which must be a clean checkout at that revision (see `AGENTS.md`, "Visual
regression"). Baselines live in the gitignored
`tests/visual/baseline/public-corpus/`. The fixtures load the Noto Sans faces
from the packages' `@fontsource` dependencies (`?fonts=corpus`), so the run
is offline, and they fail if any declared face does not load rather than
comparing fallback glyphs. The private fast and full tiers stay
maintainer-only.

## Licence

TODO(maintainer): choose the licence for the specimen sources and generated
files before publishing (for example CC0-1.0 or the repository's MIT licence).

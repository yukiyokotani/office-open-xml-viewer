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
```

`corpus:gap` compares the public corpus with a local private corpus under
`packages/<fmt>/public/private/<fmt>/` (top level), if one exists. It reduces
both to structural schema vocabulary only (element nesting, attribute names,
relationship types, binary part extensions; every `attr=value` feature is
dropped), prints covered / total per format and the uncovered features
grouped by part and element, ranked by how many private files use them. It
prints no file names and writes nothing. Without a private corpus it says so
and exits 0. Use it to choose the next specimens.

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

## Adding a case

1. Pick the next number for the format and a short slug; create
   `cases/<fmt>-NNN-<slug>.mjs` exporting
   `{ id, slug, format, title, covers, parts }`, where `parts()` returns
   part name → XML string (use the composer in `lib/<fmt>.mjs`).
2. List in `covers` the XML features the case exists to exercise, as
   `part|parent>child`, `part|element@attr` or `part|element@attr=value`
   with the prefixes of `lib/xml.mjs` (`w`, `x`, `a`, `p`, `r`). Values are
   checked only when they are enum-like (letters only); numeric attributes are
   covered by presence.
3. Run `pnpm corpus:build`, open the file in Word, Excel or PowerPoint (it
   must open without repair), and commit the case and its output together.

## Design rules

The corpus follows a grid in the sense of Josef Müller-Brockmann: necessary
and sufficient, no decoration.

- **One feature family per file.** One page, sheet or slide unless pagination
  is the subject.
- **Furniture.** A header in the top-left field: id, title, one line saying
  what to verify. Nothing else.
- **Grid.** All lengths come from `lib/grid.mjs`; a case never writes a
  free-standing number for a position, indent, tab stop or width.
  - DOCX: A4 portrait; margins left 25, right 20, top 20, bottom 23 mm; live
    area 165 × 254 mm; 12 pt baseline grid of 60 lines; 4 columns of 37.5 mm
    with 5 mm gutters; 6 rows of 9 lines separated by one empty line. Line
    spacing is exact, paragraph spacing a whole number of lines.
  - PPTX: 16:9 (12 192 000 × 6 858 000 EMU); margin 457 200 EMU (0.5 in)
    on all sides; 12 columns of 800 100 EMU with 152 400 EMU (12 pt)
    gutters; 12 pt baseline; 4 rows of 9 lines separated by one line.
    The indent module is one column plus one gutter, 952 500 EMU.
  - XLSX: column A is a 2-character margin; the column module is 12
    characters (half and whole multiples); rows are 12 pt, one baseline, and
    taller rows are whole baselines (title row 36 pt). A 2-character gutter
    column separates a flush-right column from a following flush-left one.
  - Headers and footers (DOCX) sit two baselines outside the live area;
    pictures, charts and shapes occupy whole grid fields. Text insets of
    text boxes and table cells are zero so the first glyph sits on the edge;
    a shape whose outline is drawn insets its text by one gutter (12 pt,
    `OUTLINED_INSET`) so the text never touches the rule.
- **Type.** One sans family in two weights; at most three sizes per file,
  each on the baseline (DOCX 8/12, 9/12, 24/24; PPTX 10/12, 20/24, 40/48;
  XLSX 8, 9 and 24 pt). On slides, 20/24 is for display lines and single-line
  labels (paragraph levels, placeholders, table cells); running prose is set
  in the text size, 10/12, like the Verify line. Flush left, ragged right. No centring, justification,
  italics, underline, colour or effects unless that is the feature under
  test. Line breaks in running text are written into the text, so wrapping
  never moves a block off the grid; a specimen about wrapping says so.
- **Font.** Roboto (Latin) and Noto Sans JP (Japanese), both SIL Open Font
  License 1.1. They are the open faces this repository's Google Fonts
  registries already resolve (`packages/core/src/fonts/google-fonts.ts`,
  `packages/core/src/fonts/scripts.ts`); see the note on VRT below.
- **Colour.** Black on white; one accent, cyan `12BFD8`, only when a feature
  needs a colour. Fills are solid; gradients appear only in the gradient
  specimen (pptx-009).
- **Text.** Original, sober and self-describing: a specimen names its own
  values (a tab stop line names its stop positions). Data are timeless,
  verifiable facts: ISO 216 sizes, unit definitions, typographic point
  systems, calendar and number facts. English first; Japanese plain prose
  (である体) only where East Asian layout is the feature. No lorem ipsum, no
  placeholder words, no quoted or copyrighted material.
- **Minimal XML.** No rsids, no unused styles, no document properties; only
  the parts a consumer needs to open the file without repair.

## Where specimens leave the grid on purpose

- PPTX group children use the group's child coordinate space, and connector
  ends sit at the middle of shape edges (half a row); the grid test skips them.
- DOCX inline pictures sit in a paragraph with single (auto) line spacing,
  so that line is as tall as the picture plus the font's descent.
- DOCX vertical writing (docx-007): the 165 mm live width is not a whole
  number of 12 pt lines (38.98).
- DOCX paragraph side borders stand in the gutters, outside the columns.
- Where Word decides placement (column balancing, the footnote area), the
  result is Word's, not the grid's.

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
`tests/visual/baseline/public-corpus/`. The fixtures load Roboto and Noto Sans
JP from the packages' `@fontsource` dependencies (`?fonts=corpus`), so the run
is offline, and they fail if any declared face does not load rather than
comparing fallback glyphs. The private fast and full tiers stay
maintainer-only.

## Licence

TODO(maintainer): choose the licence for the specimen sources and generated
files before publishing (for example CC0-1.0 or the repository's MIT licence).

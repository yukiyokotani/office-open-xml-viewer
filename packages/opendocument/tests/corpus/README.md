# OpenDocument validation corpus (local only)

Ground truth for the OpenDocument readers (#1934) is LibreOffice's PDF export of
the same ODF file. These scripts build that corpus from the private OOXML
samples and measure a checkout against it. Nothing they produce is committed:
the corpus and every report live under the main checkout's git-ignored
`packages/opendocument/public/private/`.

## Requirements

- LibreOffice at `/Applications/LibreOffice.app` (or `SOFFICE=<path>`).
- Homebrew `fontconfig` (`brew install fontconfig`). Headless LibreOffice on
  macOS enumerates fonts through fontconfig; without
  `/opt/homebrew/etc/fonts/fonts.conf` it cannot see the system CJK fonts and
  exports CJK text as blank glyphs.
- Poppler (`pdftoppm`, `pdfinfo`), Chrome, Rust with `wasm-pack`.

## 1. Build the corpus (once; resumable)

```bash
node packages/opendocument/tests/corpus/convert.mjs --jobs 6
```

Converts the OOXML files under `packages/{docx,xlsx,pptx,legacy-office}/public/private/**` into
`packages/opendocument/public/private/{odt,ods,odp}/<same relative path>`,
plus `<name>.pdf` exported by LibreOffice from the converted ODF file. Existing
outputs are kept. Failures go to `convert-failures.local.jsonl` and are skipped
on later runs unless you pass `--retry-failed`.

## 2. Run the ODT harness against this checkout

```bash
node packages/opendocument/tests/corpus/run.mjs
```

Options:

- `--build`: rebuild the DOCX and ODT WASM first.
- `--filter <substring>`, `--list <file of corpus-relative paths>` and
  `--limit N`: run a subset.
- `--jobs N`: the number of Chrome workers.
- `--fresh`: discard the previous results under the same label.

Results go to `packages/opendocument/public/private/harness/<sha>[-dirty]/ (or `--label`)`:

- `report.md` and `summary.json`: the share of files opened, page-count
  matches, any rejection codes, the page-diff distribution and the worst
  pages.
- `worst/`: side-by-side PNGs of the worst pages (LibreOffice | ours | diff).
- `results/`: one JSON per file. A rerun with the same label resumes.

The reader rejects only corrupt, encrypted or oversized input; the native
`inspect_odt` example, run on the same files, records the reason for any
rejection, and the report lists files where the browser result and the
native result disagree. Set `ODF_HARNESS_MAX_PAGES=N` to compare only the
first N pages of each document for a quick sample.

The renderer loads the opt-in equation and TIFF renderers, as the DOCX visual
fixture does. `ODF_HARNESS_LO_FACES=1` also serves, through the `fonts` load
option, the faces LibreOffice drew each PDF with (`fonts.mjs`: read with
`pdffonts`/`pdftohtml`; LibreOffice's Carlito/Caladea replacements, its
fallback faces for families this machine lacks, and its bundled fonts). It
is off by default: the DOCX layout measures a served face's line box without
the face's hhea line gap, so served faces skew line pitch.

Page scores:

- **Layout diff** (primary): the share of inked pixels, on either side, with
  no ink on the other side within 2 px. It tracks missing, extra and moved
  content, and ignores antialiasing.
- **Ink diff** and **page diff**: the raw pixelmatch counts, normalized by the
  inked area and by the page area.

## 3. Run the ODP harness

```bash
node packages/opendocument/tests/corpus/run-odp.mjs [--filter substr] [--jobs N]
```

It classifies every ODP natively (`inspect_odp`: admission and hidden
slides, which LibreOffice's PDF export skips), renders each shown slide in
Chrome through the `odp` model source and the PPTX renderer
(`odp-corpus.spec.ts`), and compares it with its PDF page. Slides are mostly
full-bleed colour, so the layout score compares edge structure (pixels whose
colour differs from a neighbour) instead of ink: the share of edge pixels,
on either side, with no edge on the other side within 2 px. Results go to
`harness/odp-<sha>/` (`report.md`, `summary.json`, `worst/`). The run holds
`/private/tmp/ooxml-vrt.lock`; `--no-lock` is for small subsets only.

ODS has no reader yet.

---
name: corpus-specimen
description: Author, change or review public OOXML specimens in tests/corpus (docx/pptx/xlsx test files built from hand-written XML). Use whenever issue, bug-fix or feature work exposes an unimplemented or broken OOXML feature (to hand a minimal pattern to the corpus), and when adding a specimen, turning a private-sample finding or a coverage gap into a public test file, writing a binary helper (PNG, EMF) for the builder, or checking how specimens render.
---

# Public specimen authoring

The specimens are built by code in the repository (`tests/corpus/`), so
anyone and CI can rebuild them byte for byte (`build.mjs --check`). This skill
holds the rules and the procedure; it holds no generator code. Put new
builders and binary writers in `tests/corpus/lib/`, never here.

Read `tests/corpus/README.md` for the layout, the commands and the VRT run.

## Procedure

1. **Choose the gap.** Use one of:
   - `pnpm corpus:gap --format <fmt>`: XML vocabulary the private corpus
     uses and no specimen does.
   - `pnpm corpus:coverage --format <fmt>`: renderer lines only private
     files reach (V8 coverage through the VRT fixture; takes minutes on the
     full private corpus; the JSON report names private files and stays
     local).
   - A maintainer's account of what a private sample exercised (see
     "Specimens from a private sample" below).
2. **Write the case.** `tests/corpus/cases/<fmt>-NNN-<slug>.mjs` exporting
   `{ id, slug, format, title, covers, parts }`; `parts()` returns part name
   → XML string through the composer in `lib/<fmt>.mjs`. Every length comes
   from `lib/grid.mjs`. In `covers` list the features the case exists to
   exercise as `part|parent>child`, `part|element@attr` or
   `part|element@attr=value`, with concrete part names
   (`word/charts/chart1.xml`) and the prefixes in `lib/xml.mjs` `NS` (add a
   prefix there when a new namespace appears). Values are checked only when
   enum-like; numeric attributes are covered by presence.
3. **Build and test.** `pnpm corpus:build`, then
   `node --test tests/corpus/*.test.mjs` and
   `node tests/corpus/build.mjs --check`. Builders must be deterministic:
   reset any module-level counter inside `parts()`, never read the clock,
   environment or locale.
4. **Look at it.**
   `node tests/corpus/render.mjs --format <fmt> --cases <id,...> --out <scratch dir>`
   writes every page, slide or sheet and 6 × 2 contact sheets
   (`sheet-N.png`). Read the contact sheet first, then single pages where
   something looks off. Check that the page does what its Verify line says,
   that no text overruns its field or wraps where the text has no break,
   and that blocks do not overlap. Fix the specimen when the specimen is
   wrong; record a renderer difference as a finding instead of bending the
   specimen to the renderer.
5. **Office check.** Each specimen must open in Word, Excel or PowerPoint
   without repair. Batch these checks (usually delegated); note any repair
   prompt as a bug in the case.
6. **Reference PDFs.** A new or changed DOCX/PPTX specimen (and an XLSX
   one with a chart) needs fresh `<stem>.win.pdf` and `<stem>.mac.pdf`
   exported by Office (Windows: Word/PowerPoint/Excel export to PDF; Mac:
   the same with "Best for printing", the Mac's own engine), with the
   corpus fonts installed. Then `node tests/corpus/references.mjs --record
   <id>`. Until then `references.mjs` lists it as missing; a changed
   specimen whose old PDFs remain fails the test as stale, so export again
   or delete them.
7. **Commit** the case, any lib change and the regenerated
   `packages/<fmt>/public/corpus/` file together, on a feature branch. No
   private file names, content or paths in messages or PR text.

## Patterns from issue work

The public corpus replaces accumulated private samples with the smallest set of
files that still covers every behaviour, one that outside developers can use
too. Every issue, bug fix or feature task feeds it.

1. Write the pattern yourself. When the work exposes a feature that is
   unimplemented or broken, keep only the elements, attributes, values and
   geometry that trigger the behaviour. Never copy text, names, images or file
   names from private, third-party or copyrighted samples.
2. Check for overlap. Look at the existing specimens' `covers` lists and the
   feature report. Extend a specimen rather than adding one that overlaps it.
3. Hand it over in one of two ways:
   - Add or extend the specimen yourself, following this skill. That includes
     exporting the Windows and Mac Office PDFs.
   - If another session owns the corpus, send it four things:
     - the minimal XML;
     - what Office does, as observed facts from an Office-exported PDF: page,
       position and size, line and page breaks;
     - what the viewer does now: an error, or the visible difference;
     - the branch or PR with the fix, if there is one.
4. In the PR or handoff, say which pattern went where. The fix still needs its
   own focused regression test; a specimen does not replace it.

## Specimens from a private sample

Private samples are not redistributable; nothing in a specimen (text, data,
structure, measurement) is copied from one. What carries over is the
*pattern*: the feature combination and the layout state that broke the
renderer, e.g. one-column abstract and two-column body on one page, columns
balanced at a section end, a table split by rows at a page foot, a heading
moved by keepNext, vertical text with upright digits. Build the pattern
fresh on the grid with self-describing text, and name in the Verify line the
state the reader should check. Keep interview notes in the ignored
`packages/<fmt>/public/private/sample-notes.md`, never in the repository.

## Office font specimens

The normal specimens use only the Noto Sans family, which the VRT fixtures
load (`?fonts=corpus`), so they test layout with exact fonts. A separate,
clearly titled set names Office-only fonts (Calibri, Cambria, Aptos,
Yu Mincho, Yu Gothic, MS Mincho, MS Gothic, Meiryo, Times New Roman, Arial
and the like) without embedding them. Viewed outside Office they render with
substitutes, so line breaks and page counts will differ from Office; that is
expected. Their purpose is regression: with a substitute the layout must not
break: no text outside its box or cell, no overlapping lines, no clipped
page foot, tables and columns intact, and a page count within a line or two
of the authored one. Their Verify line says so. Put dense packed pages
(document grid with tight line pitch, tables near the page foot) in this set
when the pattern depends on the font's metrics.

## Binary parts

Pictures and metafiles are written by deterministic helpers in
`tests/corpus/lib/` (`png.mjs`; `emf.mjs` for Enhanced Metafiles). A helper
writes records explicitly (no external tools, no compression that varies by
platform) and draws on the grid in device units that map to whole grid
fields. A metafile specimen states its clip, transform and record list in
its caption so a reader can predict the drawing.

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
- **Type.** One family in two weights; at most three sizes per file,
  each on the baseline (DOCX 8/12, 9/12, 24/24; PPTX 10/12, 20/24, 40/48;
  XLSX 8, 9 and 24 pt). On slides, 20/24 is for display lines and single-line
  labels (paragraph levels, placeholders, table cells); running prose is set
  in the text size, 10/12, like the Verify line. Flush left, ragged right. No centring, justification,
  italics, underline, colour or effects unless that is the feature under
  test. Line breaks in running text are written into the text, so wrapping
  never moves a block off the grid; a specimen about wrapping says so.
- **Font.** The Noto Sans family for every script: Noto Sans (Latin), Noto
  Sans JP, Noto Sans Arabic and Noto Sans Hebrew, all SIL Open Font License
  1.1. They are the faces this repository's script registry already resolves
  (`packages/core/src/fonts/scripts.ts`); see the note on VRT below.
  Every DOCX specimen carries a `word/fontTable.xml` listing them with their
  charsets (`docxPackage` writes it). Without it, Word for Windows replaces
  the docDefaults East Asian font with the theme font, so the Japanese text
  comes out in Yu Mincho. A case that writes its own font table must start
  it with `FONT_TABLE_ENTRIES`.
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

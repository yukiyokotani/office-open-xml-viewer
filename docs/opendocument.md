# Experimental OpenDocument sources

Install the viewer and its optional companion at matching versions:

```sh
npm i @silurus/ooxml @silurus/ooxml-opendocument
```

The companion pins its `@silurus/ooxml` peer to the exact same version. A
stale runtime or model-source protocol fails before input loading with
`CompanionRuntimeMismatchError` (`code: 'companion-runtime-mismatch'`).

OpenDocument text can be opened with the ordinary DOCX viewers and loaders
through an optional model source:

| Input | Entry | Reader / factory | Option | Loaders and viewers |
| --- | --- | --- | --- | --- |
| ODT, OTT | `@silurus/ooxml-opendocument/odt` | `odt` / `createOdt()` | `odt` | `DocxViewer`, `DocxScrollViewer`, `DocxDocument.load`, `openDocxDocument`, `materializeDocxDocument` |

```typescript
import { DocxViewer } from '@silurus/ooxml/docx';
import { odt } from '@silurus/ooxml-opendocument/odt';

const canvas = document.querySelector('canvas') as HTMLCanvasElement;
const viewer = new DocxViewer(canvas, { odt });
await viewer.load(odtOrDocxBytes);
```

The reader maps the document directly into the DOCX document model, without
generating an OOXML package; the ordinary layout and Canvas renderer draw it.
Importing the entry or calling `createOdt()` fetches nothing. When an ODT file
is opened, the parser Worker (or Node) imports the reader's self-contained ES
module (`odt-source-module.mjs`) and its WASM (`opendocument_odt_bg.wasm`).
Serve these files with the other package assets and allow the module URL in a
Content Security Policy. Applications with a custom asset pipeline pass
absolute URLs:

```typescript
import { createOdt } from '@silurus/ooxml-opendocument/odt';

const odt = createOdt({
  moduleUrl: new URL('/assets/odt-source-module.mjs', location.href).href,
  wasmUrl: new URL('/assets/opendocument_odt_bg.wasm', location.href).href,
});
```

## Admission

The source claims a package only by its `mimetype` file, which a packaged
OpenDocument text document must carry (ODF 1.3 Part 2 §3.3, Part 3 §2.2.3).
Producers store it first and uncompressed; a package re-zipped by a general
archiver, with the file compressed, listed later or given extra fields, is
claimed and read as well, as LibreOffice does (the claim compares the
central directory's checksum and size of the entry with the media types, so
nothing is inflated in the calling realm). Other inputs, including other
OpenDocument families, continue on the OOXML path unchanged. The ZIP is read
with the same structural checks and inflation limits as OOXML packages,
including the caller's `resourceLimits`; item names are compared exactly.
The manifest is read only to detect encryption: encrypted packages reject
with `unsupported-encryption`, and the manifest's version and root entry are
not checked. The declared ODF version is not checked either: older documents,
and documents that do not declare one, are read by the ODF 1.3 rules.
A style definition that repeats a name within a family replaces the earlier
one instead of refusing the document; a style or list
style without a name, or a style without a family, cannot be referenced and
is skipped. List level definitions outside levels 1 to 10 are never used.

## Rendering policy

The reader renders every OpenDocument text feature it meets. Each one is
read as ODF 1.3 Part 3 specifies and mapped onto the DOCX document model;
where the specification leaves a choice to the consumer, or LibreOffice
reads it differently, LibreOffice is the reference. A document is rejected
only when it is corrupt, encrypted or unsafe, or when it exceeds a resource
limit (the caller's ZIP limits, the XML limits and the reader's own bounds on
the memory it holds and the work it does, which its allocator measures). A feature the DOCX model cannot express exactly is drawn with its
nearest DOCX equivalent instead of refusing the document, so a document may
show small differences from LibreOffice. The reader never executes scripts,
macros or forms, and never fetches external resources (linked files outside
the package are not drawn).

## What is rendered

- Text: paragraphs, headings, spans, white space (`text:s`, tabs, line
  breaks, soft hyphens), hyperlinks, bookmarks, fields (page number and
  page count are computed by the layout; every other field shows its
  stored value), rubies, hidden text and paragraphs, and the accepted state
  of tracked changes (insertions shown, deletions left out).
- Character properties: font faces of the Western, Asian and complex script
  classes (each class keeps its own face, size, weight and posture), sizes
  including percentages and relative sizes, weight, posture, colours with
  transparency, background, underline styles, types, widths and colours,
  strike-through, letter spacing, horizontal scaling, raised and lowered
  text, upper case, small capitals, lower case and title case, pair
  kerning, emphasis marks, two lines in one, rotated characters, character
  borders and languages.
- Paragraph properties: indents including automatic first-line indents,
  spacing (adjacent margins add, as in LibreOffice), contextual spacing,
  alignment (start/end follow the writing direction; a justified last
  line), proportional, fixed, minimum and leading line heights, borders
  and padding, backgrounds and area fills, tab stops and leaders,
  keep-with-next and keep-together, widows and orphans, page and column
  breaks, right-to-left paragraphs, Asian line breaking, hanging
  punctuation, autospacing, layout-grid snapping and drop caps.
- Lists and heading numbering: every list and outline numbering style,
  continued and restarted lists, list headers, nested lists with their own
  styles, numbered paragraphs, label alignment and the older label width
  and position mode, bullets, pictures as bullets and LibreOffice's number
  formats (Roman, letters, circled numbers, Asian counting systems, other
  alphabets and digit sets, ordinals and words).
- Pages: page sizes, margins, gutters, borders and padding, page number
  formats and restarts, page columns with separators, layout grids,
  vertical writing, left/right page usage, several page styles in one
  document (each change starts a DOCX section), first-page styles,
  headers and footers including left-page and first-page variants, notes
  configuration and footnote separators, and line numbering.
- Sections: columns (balanced at the section end, as DOCX continuous
  sections are), section indents, and hidden sections.
- Notes: footnotes and endnotes with automatic or custom citations, in the
  configured number format.
- Tables: column grids from absolute or relative widths, alignment and
  margins, spanned and covered cells, repeated header rows, row heights,
  rows kept together, cell borders (all line styles), padding,
  backgrounds, vertical alignment and text direction, and nested tables.
- Drawings: pictures (PNG, JPEG, GIF, BMP, TIFF, WebP, SVG, EMF and WMF,
  linked in the package or embedded as base64) with cropping, mirroring,
  opacity and colour modes; text frames; custom shapes, whose enhanced
  geometry (equations, modifiers and the full path command set) is
  evaluated; rectangles, ellipses, lines, polylines, polygons, paths and
  groups; fills (solid, gradient, hatch and bitmap), lines with dashes,
  joins, caps and arrowheads; anchoring as character, to the character,
  paragraph or page, with every horizontal and vertical position and
  relation, wrapping modes, distances, stacking and overlap control; and
  formula objects, whose MathML is drawn as math. A drawing anchored to a
  physical page (`text:anchor-page-number`) is drawn on that page and wraps
  that page's text, wherever it is written (the body, a paragraph, or a
  master page's header or footer, as LibreOffice reads it), and the document
  runs to that page with empty pages when it is past the content.
- Embedded fonts (`svg:font-face-src`) are registered for the document.

## LibreOffice compatibility rules

Where ODF leaves the behaviour to the consumer, or LibreOffice reads the
specification differently, the reader follows LibreOffice:

- proportional line height scales the font's natural line height;
- spaces, punctuation, symbols and emoji outside ODF's script table take the
  script class of the preceding character (the following one at paragraph
  start);
- the nearest style that sets a font colour or the automatic colour decides
  the colour;
- automatic text colour over a character or paragraph background is black
  or white by LibreOffice's `Color::IsDark()` rule;
- list labels are numbered and placed as LibreOffice numbers and places
  them (continued lists, phantom parents, label alignment, the list tab
  stop and tab grid);
- `BalanceSpacesAndIdeographicSpaces`, `DoNotJustifyLinesWithManualBreak`,
  `JustifyLinesWithShrinking` and `TabsRelativeToIndent` from
  `settings.xml` are applied;
- table columns are fitted to the table as LibreOffice fits them, rows are
  cut at a page as LibreOffice cuts them, and shared cell edges keep the
  wider rule;
- superscript and subscript without an explicit raise use LibreOffice's
  automatic escapement.
- a line snapped to a layout grid takes as many cells as the face's own
  line box needs (without Word's 1.3 East Asian box), and the page admits it
  only when all of its cells fit.

## Known differences from LibreOffice

The DOCX layout draws some features differently from LibreOffice:

- automatic hyphenation (`fo:hyphenate`) is not applied;
- LibreOffice's fixed and minimum line placement (a fixed line's baseline
  four fifths of the way down the line; a minimum line's extra space above
  the text) applies only on horizontal pages, to paragraphs that do not snap
  to a layout grid;
- LibreOffice spreads a justified line's spare width in whole layout units;
- section backgrounds, character shadows, outlines and reliefs, overlines
  and Graphite font shaping are not drawn (page backgrounds are);
- charts and other embedded objects without a formula are drawn only when
  their replacement picture is a format the renderer draws;
- widow and orphan counts above two are treated as two;
- the text box chain (`draw:chain-next-name`) is not flowed across boxes.

## Presentations (ODP)

OpenDocument presentations open with the ordinary PPTX viewers and loaders:

| Input | Entry | Reader / factory | Option | Loaders and viewers |
| --- | --- | --- | --- | --- |
| ODP, OTP | `@silurus/ooxml-opendocument/odp` | `odp` / `createOdp()` | `odp` | `PptxViewer`, `PptxScrollViewer`, `PptxPresentation.load`, `openPptxPresentation`, `materializePptxPresentation` |

```typescript
import { PptxViewer } from '@silurus/ooxml/pptx';
import { odp } from '@silurus/ooxml-opendocument/odp';

const viewer = new PptxViewer(canvas, { odp });
await viewer.load(odpOrPptxBytes);
```

The ODP reader maps a presentation directly into the PPTX slide model; the
ordinary PPTX renderer draws it. Its assets are `odp-source-module.mjs` and
`opendocument_odp_bg.wasm`; `createOdp({ moduleUrl, wasmUrl })` relocates
them as for ODT. The source claims a package by its `mimetype` file
(`application/vnd.oasis.opendocument.presentation` and `-template`), and
admits packages, versions and styles as the ODT reader does.

It follows the same rendering policy as the ODT reader: every presentation
feature is read from its ODF definition and mapped onto the slide model, and
only corrupt, encrypted or unsafe packages, or packages over the resource
limits, reject. Where ODF leaves a choice to the application, LibreOffice
Impress is the reference. It reads:

- slides and master pages: master background objects, the header, footer,
  date/time and page-number frames with the slide's display settings and
  declarations, slide and master backgrounds, hidden slides and speaker
  notes;
- the style cascade of graphic and presentation styles (with the per-level
  outline styles of outline placeholders), paragraph and text styles, and
  list styles with their indents, bullet characters and pictures, and
  number labels;
- custom shapes from their enhanced geometry (formulas, modifiers, every
  path command, text areas, mirroring), rectangles, ellipses and arcs,
  polygons, paths, lines, connectors, captions, groups, transforms and shape
  hyperlinks;
- area fills (solid, every gradient style with borders and multi-colour
  stops, uniform and gradient transparency, hatches as the nearest pattern,
  bitmaps stretched, tiled or placed once), lines with dashes, caps, joins
  and markers, shadows, glow and soft edges;
- text: fonts per script, sizes, weight, posture, colours (including
  automatic colour), underline and strike-through, raised and lowered text,
  capitals, letter spacing, highlighting, outline and shadowed text, tabs,
  line breaks, fields and hyperlinks; paragraph alignment, indents,
  spacing, line height, tab stops, columns, text anchoring, padding,
  wrapping and shrink-on-overflow;
- pictures (cropping, mirroring, colour modes, luminance/contrast,
  transparency), StarView metafile pictures (converted to SVG), embedded
  objects and OLE objects through their replacement pictures, and tables
  (column widths, row heights, spans, cell fills, borders, templates).

Charts (`draw:object` frames holding an `office:chart` document) are read
with their own data table into the shared chart model; an object the chart
reader finds structurally invalid is drawn from its replacement picture.
Known gaps: 3D extrusion of custom shapes and fontwork text paths are drawn
flat; animations and transitions are not played; text measurement follows
the fonts available to the browser, so text set in fonts LibreOffice
substitutes (for example metric-compatible replacements) can wrap
differently.

Resource limits: while it builds the deck, the reader may grow its heap by
at most 512 MiB beyond the package's parsed XML and take a bounded number of
steps, as measured by the same allocator as the other readers; the embedded
charts of a deck may allocate at most 1 GiB together, and the deck's
elements, slides (10,000), group nesting and retained pictures and slide
data have bounds of their own. A presentation over a limit is rejected with
an error that names the limit, instead of being shown in part.

## Spreadsheets (ODS)

OpenDocument spreadsheets open with the ordinary XLSX viewers and loaders:

| Input | Entry | Reader / factory | Option | Loaders and viewers |
| --- | --- | --- | --- | --- |
| ODS, OTS | `@silurus/ooxml-opendocument/ods` | `ods` / `createOds()` | `ods` | `XlsxViewer`, `XlsxWorkbook.load`, `openXlsxWorkbook`, `materializeXlsxWorkbook` |

```typescript
import { XlsxViewer } from '@silurus/ooxml/xlsx';
import { ods } from '@silurus/ooxml-opendocument/ods';

const container = document.getElementById('viewer') as HTMLElement;
const viewer = new XlsxViewer(container, { ods });
await viewer.load(odsOrXlsxBytes);
```

The source claims a package by its `mimetype` file
(`application/vnd.oasis.opendocument.spreadsheet` or `-template`), admits
packages as the ODT reader does, and loads
`ods-source-module.mjs` and `opendocument_ods_bg.wasm` the same way as the ODT
reader. It follows the same rendering policy as the ODT reader: every
spreadsheet feature is read from its ODF definition and mapped onto the
workbook model, and only corrupt, encrypted or unsafe packages, or packages
over the resource limits, reject. Where ODF leaves a choice to the application, LibreOffice Calc is the
reference.

What is read:

- tables as sheets (hidden tables, tab colours, right-to-left tables),
  columns and rows with their repeat counts (kept as runs, so LibreOffice's
  million-row fillers cost nothing), widths, heights, hidden and grouped
  (outlined) rows and columns, frozen panes, grid and zero-value display;
- cell values of every `office:value-type` (float, percentage, currency,
  date, time, boolean, string), error results, and the cached value of
  formula cells (formulas are kept as text and never calculated);
- merged cells (`table:number-columns-spanned` / `-rows-spanned`), rich text
  spans, line breaks and spaces inside a cell, hyperlinks and annotations;
- cell styles: fonts, colours, bold, italic, underline, strike-through,
  super/subscript, fills, borders (including diagonals), alignment, wrap,
  shrink to fit, rotation, stacked text and indents;
- data styles (number, currency, percentage, scientific, fraction, date,
  time, boolean and text styles, with `style:map` conditions), converted to
  the format codes the workbook renderer evaluates, including the quarter
  and week-of-year date fields;
- named ranges and expressions, print ranges, database ranges with filter
  buttons, content validations, conditional formats (cell conditions, colour
  scales, data bars, icon sets) and sparklines;
- pictures, text boxes and shapes anchored to cells or to the sheet, and
  embedded charts (bar, column, line, area, pie, ring, scatter, bubble,
  radar, filled radar and stock charts and their combinations, with titles,
  legends, axes, gridlines, data labels, series and point colours) read from
  the sheet's cells or the chart's own data table.

Column widths are authored lengths; the reader asks the renderer for the
default cell font's digit width and encodes each width so that the grid is
painted at the authored size.

Where the specification and LibreOffice differ, or the workbook model
cannot hold what LibreOffice shows, the reader does the following:

- Names in dates (months, weekdays, the long quarter `1st quarter`) are
  English, or Japanese in a Japanese-calendar format, as everywhere in the
  workbook renderer; LibreOffice uses the format's language.
- Weeks of the year are numbered as LibreOffice numbers them for the
  format's language and country (for example from Monday with at least four
  days in the first week for `de-DE`, from Sunday with 1 January in week 1
  for `en-US`). A format without a language uses the `en-US` numbering;
  LibreOffice would use the locale it runs in.
- A date before 1 March 1900 is shown as the same calendar date. The
  workbook date system has no date before 31 December 1899: an earlier date
  shows its time of day only, or the invalid-date marker.
- A formula result that LibreOffice shows as nothing (an empty-cell or
  empty-string result saved as the number 0) is shown as nothing and stays
  the number 0.
- A cell with a style of its own takes nothing from its row's or column's
  default cell style, as in LibreOffice (ODF 1.3 Part 3 §16.2 continues the
  search for unset properties there).
- A number style without `number:decimal-places` is shown in the General
  format, as LibreOffice shows it.
- A minutes field that stands next to neither hours nor seconds is shown as
  LibreOffice shows it: the month, or all the minutes of a duration when the
  style does not truncate.
- Numbers are formatted by the workbook renderer's formatter, which follows
  Excel where the two differ: a negative number that rounds to zero keeps
  its minus sign (`-0` where LibreOffice shows `0`), and General shows at
  most eleven characters.
- Hatch lines turn counter-clockwise from the horizontal by the hatch's
  rotation, as LibreOffice draws them (ODF 1.3 Part 3 §19.209 words the
  rotation as clockwise from the vertical axis).

Not representable in the workbook model, and therefore not shown: page
layout, headers and footers (the workbook viewer shows the grid, not printed
pages; print ranges become `_xlnm.Print_Area` names). Data pilot tables are
shown as the cells LibreOffice saved for them. Shape area fills are solid
colours, gradients (linear, axial, radial, ellipsoid, square, rectangular,
with border and centre; the rotation of a non-linear gradient is not
applied), hatches, and pictures that are stretched, repeated (at their
size, from their reference point, with its offsets and the tile stagger) or
drawn once. Fills and lines keep their opacity; an opacity gradient is shown
over a solid colour and over a colour gradient of the same geometry, and as
its mean opacity over other fills. Chart axis and label number formats
include scientific and fraction formats; a label's percentage has its own
format, and chart colours keep their opacity. Of the conditions of one
conditional format, the first that holds for a cell gives it its style, as in
LibreOffice; overlapping conditional formats all apply.

The workbook grid has 1,048,576 rows and 16,384 columns: rows and columns
of a table beyond it are not shown, as when LibreOffice saves such a table
as a workbook.

Resource limits: while it builds the workbook, the reader may grow its heap
by at most 512 MiB beyond the package's parsed XML, materialize at most 8
million cells and take a bounded number of steps. A spreadsheet over a
limit is rejected with an error that names the limit, instead of being shown
in part. That measured limit covers building the workbook. Before it, the
package directory and the text and parsed form of
`styles.xml`, `content.xml` and `settings.xml` are bounded by the package
limits (entries, bytes per entry, total inflated bytes) and the XML parsing
limits; after it, the workbook is handed to the viewer through the bounded
workbook and worksheet writers, a few rows at a time. White space is read as LibreOffice reads it: the text of a cell is
kept as written (Calc does not collapse it), and comments, shape text and
chart titles collapse runs of white space, keeping the space that ends a
paragraph. Connectors are drawn along their path (bent and curved ones too),
and horizontal and vertical lines as lines; callouts (`draw:caption`) are
their text box and the line to their point; ellipse arcs, sections and cuts
are drawn as such. Dimension lines (`draw:measure`), regular polygons
(`draw:regular-polygon`, which LibreOffice writes as polygons) and form
controls are not drawn. Custom shapes are drawn from
their preset or their enhanced-geometry path (equations are evaluated; arcs
are flattened to short segments). Embedded objects that are not charts are
drawn from their replacement image
when the package stores it in a browser-decodable format.

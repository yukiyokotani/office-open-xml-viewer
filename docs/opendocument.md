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

The source claims a package only by its leading `mimetype` file, which a
packaged OpenDocument text document must carry (ODF 1.3 Part 2 §3.3, Part 3
§2.2.3). Other inputs, including other OpenDocument families, continue on the
OOXML path unchanged. The ZIP is read with the same structural checks and
inflation limits as OOXML packages, including the caller's `resourceLimits`;
item names are compared exactly. Encrypted packages reject with
`unsupported-encryption`.

## Supported content

The reader is fail-closed: a document that uses anything outside the admitted
subset rejects as a whole with `OoxmlError.code === 'unsupported-content'`
instead of being shown partially. It never executes scripts or fetches
external resources.

Currently admitted:

- paragraphs, headings (without heading numbering), plain sections, spans,
  spaces, soft hyphens and line breaks between visible text set in the same
  font;
- page size and margins of a single page style without displayed headers or
  footers;
- paragraph indents, left/centre/right/justified alignment (a justified
  paragraph with a line break only where LibreOffice leaves that line
  unjustified, and not with space shrinking), proportional line height
  from 100% to 300% on text of one font and size, solid paragraph backgrounds
  (except two consecutive filled paragraphs with spacing between them), page
  breaks, keep-with-next / keep-together, widow/orphan control of two lines,
  and paragraph spacing including contextual spacing;
- character fonts, sizes, bold, italic, single/double underline and
  strike-through, colours, background colour and pair kerning on or off
  (where a style sets it; LibreOffice kerns text whose styles never do). Where LibreOffice's
  settings make a space not size its line, the space must be in the face of
  the adjacent text on its line and no larger, and decorated spaces that end
  a paragraph or precede a line break are admitted only where LibreOffice
  leaves them undecorated (its Word-compatible trailing-blank setting); a
  trailing no-break or other Unicode space keeps its decoration in both;
- text in the Latin script class of ODF Part 3 Table 22, including the
  spaces, punctuation, symbols and emoji the table leaves unassigned.
- lists in label alignment mode, numbered as LibreOffice numbers them
  (continued lists, items nested without a paragraph of their own, list
  headers), with decimal, letter and Roman labels or text bullets, labels
  aligned at their start, end or centre, the level's character style and
  bullet face, and a tab, space or nothing after the label. Pictures as
  bullets are admitted when the package holds them as PNG, GIF or baseline
  JPEG files that decode completely, without rotation or animation.
  JPEG additionally admits the exact known 3,144-byte sRGB IEC61966-2.1
  profiles with perceptual or media-relative colorimetric rendering intent;
  other JPEG profiles and all PNG/GIF colour profiles remain unsupported.
  Pictures are placed in the middle of their line, no taller than the font
  size of the paragraph's first character and followed by a tab or nothing. LibreOffice's other number formats, headings
  in lists, a second sublist within one item, empty labelled paragraphs,
  decorated labels, labels that could wrap and labelled paragraphs that are
  centred or right-aligned reject for now.

Four behaviours follow LibreOffice where ODF leaves the choice to the
consumer or LibreOffice reads the specification differently; each applies
only within its measured scope:

- proportional line height scales the font's natural line height;
- spaces, punctuation, symbols and emoji outside Table 22 take the script
  class of the preceding character (the following one at paragraph start);
- the nearest style that sets a font colour or the automatic colour decides
  the colour;
- automatic text colour over a character or paragraph background is black
  or white by LibreOffice's `Color::IsDark()` rule.

Three known differences from LibreOffice remain in admitted documents:

- LibreOffice spreads a justified line's spare width in whole layout
  units, so its justified lines can end up to about 1.6 pt short of the
  margin; the viewer ends them at the margin;
- with LibreOffice's default settings, a decorated space at a wrapped
  line's end is underlined or struck through in LibreOffice and left
  undecorated by the viewer;
- LibreOffice shapes Graphite fonts (such as Linux Libertine G) with their
  own Graphite rules, which keep pair kerning even where a style turns it
  off; browsers do not implement Graphite and shape these fonts from their
  OpenType tables.

LibreOffice's layout settings (`settings.xml`) that Word-derived documents
carry are admitted where they were measured to change nothing the reader
admits; the others must keep LibreOffice's new-document values, except
`BalanceSpacesAndIdeographicSpaces`, which is applied as LibreOffice applies
it: a space at the start or end of a text portion (portions end at
formatting changes, bookmarks, reference marks, soft hyphens, no-break
spaces and hyphens, zero-width spaces, word joiners and line breaks), or
next to another space, advances half its Asian font size. Where the reader
cannot tell whether LibreOffice starts a portion next to such a space (for
example the same formatting spelled differently in two spans, or spans that
differ only in language), the document rejects. A
justified paragraph with such a space anywhere except where it ends the
paragraph or precedes a line break rejects,
because LibreOffice then lays out its lines differently.

Tabs, letter spacing, horizontal scaling,
raised/lowered text, fixed and minimum line heights, proportional line
heights below 100%, tables, frames and images, notes, fields,
comments, tracked changes, headers and footers, Asian and complex scripts,
and other formatting reject for now. The subset grows release by release.

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

The source claims a package by its leading `mimetype`
(`application/vnd.oasis.opendocument.spreadsheet` or `-template`) and loads
`ods-source-module.mjs` and `opendocument_ods_bg.wasm` the same way as the ODT
reader. Unlike the ODT reader it is not fail-closed: every spreadsheet feature
the workbook model can express is read from its ODF definition, and only
corrupt, encrypted or unsafe packages, or packages over the resource limits,
reject. Where ODF leaves a choice to the application, LibreOffice Calc is the
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
drawn once.

The workbook grid has 1,048,576 rows and 16,384 columns: rows and columns
of a table beyond it are not shown, as when LibreOffice saves such a table
as a workbook.

Resource limits: while it builds the workbook, the reader may grow its heap
by at most 512 MiB beyond the package's parsed XML, materialize at most 8
million cells and take a bounded number of steps. A spreadsheet over a
limit rejects with the resource-limit error instead of being shown in part. Custom shapes are drawn from
their preset or their enhanced-geometry path (equations are evaluated; arcs
are flattened to short segments). Embedded objects that are not charts are
drawn from their replacement image
when the package stores it in a browser-decodable format.

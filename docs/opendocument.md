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
`unsupported-encryption`. The declared ODF version is not checked: older
documents are read by the ODF 1.3 rules.

## Rendering policy

The reader renders every OpenDocument text feature it meets. Each one is
read as ODF 1.3 Part 3 specifies and mapped onto the DOCX document model;
where the specification leaves a choice to the consumer, or LibreOffice
reads it differently, LibreOffice is the reference. A document is rejected
only when it is corrupt, encrypted or unsafe, or when it exceeds a resource
limit (the caller's ZIP limits, the XML limits and the reader's construction
budget). A feature the DOCX model cannot express exactly is drawn with its
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
  formula objects, whose MathML is drawn as math.
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

## Known differences from LibreOffice

The DOCX layout draws some features differently from LibreOffice:

- automatic hyphenation (`fo:hyphenate`) is not applied;
- fixed and minimum line heights place extra space above the text, as
  LibreOffice does, only on horizontal pages without a layout grid;
- LibreOffice spreads a justified line's spare width in whole layout units;
- section backgrounds, character shadows, outlines and reliefs, overlines
  and Graphite font shaping are not drawn (page backgrounds are);
- charts and other embedded objects without a formula are drawn only when
  their replacement picture is a format the renderer draws;
- widow and orphan counts above two are treated as two;
- the text box chain (`draw:chain-next-name`) is not flowed across boxes.

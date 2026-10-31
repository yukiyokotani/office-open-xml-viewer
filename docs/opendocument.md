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
- tables of paragraphs with absolute or full-width relative column widths,
  left/centre/right alignment and a left margin (also past the text area),
  cell padding, single solid borders (collapsed as LibreOffice draws them),
  cell backgrounds, vertical alignment, fixed and minimum row heights,
  rows kept together, repeated header rows and a page break before. Merged
  or repeated cells, nested tables, lists and headings in cells, double or
  dashed borders and vertical cell text reject for now;
- lists in label alignment mode, numbered as LibreOffice numbers them
  (continued lists, items nested without a paragraph of their own, list
  headers), with decimal, letter and Roman labels or text bullets, labels
  aligned at their start, end or centre, the level's character style and
  bullet face, and a tab, space or nothing after the label. Pictures as
  bullets are admitted when the package holds them as PNG, GIF or baseline
  JPEG files that decode completely, without colour profiles, rotation or
  animation, placed in the middle of their line, no taller than the font
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
heights below 100%, frames and images, notes, fields,
comments, tracked changes, headers and footers, Asian and complex scripts,
and other formatting reject for now. The subset grows release by release.

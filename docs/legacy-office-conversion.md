# Experimental legacy Office sources

Legacy binary Office files can be opened with the ordinary viewers and
loaders through optional model sources:

- `.ppt` with the PPTX loaders and `PptxViewer`

Each reader is a separate opt-in entry. It reads a supported binary subset
directly into the existing document, workbook or presentation model, without
generating an OOXML package, and the ordinary layout and Canvas renderers
draw the result. Importing the ordinary DOCX, XLSX or PPTX entries alone does
not import, fetch or initialize any legacy reader. Without a matching source,
legacy input continues to reject with
`OoxmlError.code === 'legacy-binary-format'`, so no migration is required for
applications that do not opt in.

These readers are experimental and deliberately narrow. They reject input
they cannot represent instead of showing a partial document, and none of them
executes macros, fields, formulas, actions or embedded objects.

## Enabling a source

Pass the source in the format-generic `modelSources` option. It is available
on the DOCX, XLSX and PPTX `load()` options, on the viewers, and on the Node
session options (`openDocxDocument`, `materializeDocxDocument`,
`openXlsxWorkbook`, `openPptxPresentation` and their siblings).

```typescript
import { PptxViewer } from '@silurus/ooxml/pptx';
import { legacyPptSource } from '@silurus/ooxml/legacy-ppt';

const canvas = document.querySelector('canvas') as HTMLCanvasElement;
const viewer = new PptxViewer(canvas, { modelSources: [legacyPptSource()] });
await viewer.load(pptOrPptxBytes);
```

| Input | Entry | Factory | Loaders and viewers |
| --- | --- | --- | --- |
| PPT | `@silurus/ooxml/legacy-ppt` | `legacyPptSource()` | `PptxViewer`, `PptxPresentation.load`, `openPptxPresentation` |

The factory accepts these optional settings:

```typescript
interface LegacySourceOptions {
  wasmUrl?: string; // absolute URL of the reader's WASM
  moduleUrl?: string; // absolute URL of the reader's source module
  maxInputBytes?: number; // defaults to, and may not exceed, 256 MiB
}
```

Creating a source fetches nothing. When a claimed file loads, the parser
Worker (or Node) imports the source's self-contained ES module, emitted as
`legacy-<format>-source-module*.js` next to the package files, and initializes
the reader's dedicated WASM, `legacy_ppt_direct_bg.wasm`. Serve these files
with the other package assets, and allow the module URL wherever a Content
Security Policy restricts `script-src` or Worker imports. Applications with a
custom asset pipeline pass absolute `moduleUrl` and `wasmUrl` values:

```typescript
const source = legacyPptSource({
  moduleUrl: new URL('/assets/legacy-ppt-source-module.js', location.href).href,
  wasmUrl: new URL('/assets/legacy_ppt_direct_bg.wasm', location.href).href,
});
```

Module URLs come only from application options; document content never
selects or rewrites them.

A source can report the document's own view preferences. Precedence is: an
explicit caller option, then the source's view default, then the renderer
default. Capabilities that a source lacks degrade the same way for every
format: resource metrics are reported without a ZIP usage snapshot, and
`toMarkdown()` rejects with an "... is unsupported for this source" error.
The legacy PPT reader lacks Markdown export and ZIP accounting.

To cancel a load, destroy the document or viewer, or start another load in
its place. Node sessions keep their `signal` option.

## Admission and failure behavior

A source's synchronous `claim(bytes)` accepts only its own [MS-CFB] family:
a compound file whose directory names `PowerPoint Document` (PPT). A container
that names more than one binary family (`WordDocument` for DOC and `Workbook`
or `Book` for XLS count too), or that carries `EncryptionInfo`, is not
claimed. Every input a
source does not claim takes the unchanged OOXML path, so:

- a legacy file without a matching source rejects with the typed
  `legacy-binary-format` error;
- encrypted OOXML packages keep their existing encryption errors;
- configuring `legacyPptSource()` does not enable DOC or XLS input.

Claimed input larger than `maxInputBytes` throws a `RangeError`. The limit is
resource policy, not an Office format limit. After a source claims a file, a
reader failure rejects the load: there is no fallback to another source or to
the OOXML path. Password-protected legacy binaries, pre-CFB Office formats and
unsupported binary structures are rejected by the readers. These checks are
structural, never filename-based.

## Experimental direct PPT source

Browser:

```typescript
import { PptxPresentation } from '@silurus/ooxml/pptx';
import { legacyPptSource } from '@silurus/ooxml/legacy-ppt';

const presentation = await PptxPresentation.load(legacyPptArrayBuffer, {
  modelSources: [legacyPptSource()],
});
const canvas = document.querySelector('canvas') as HTMLCanvasElement;

try {
  await presentation.renderSlide(canvas, 0, { width: 960 });
} finally {
  presentation.destroy();
}
```

Node:

```typescript
import { openPptxPresentation } from '@silurus/ooxml/node';
import { legacyPptSource } from '@silurus/ooxml/legacy-ppt';

const session = await openPptxPresentation(legacyPptBytes, {
  modelSources: [legacyPptSource()],
});

try {
  for await (const slide of session.slides()) {
    // Consume each ordinary shared slide model.
  }
} finally {
  await session.close();
}
```

The direct source is an experimental, bounded subset, not a full-fidelity
PowerPoint implementation. Besides Markdown and ZIP accounting, it currently
lacks audio/video media and embedded fonts. It rejects constructs it cannot
represent, including unresolved paragraph margin or indentation and positive
paragraph before/after percentages. Its explicit unsupported diagnostics are
authoritative.

A slide's own `SlideShowSlideInfoAtom.fHidden` marks it hidden; hidden slides
and their content remain in the presentation. Display follows the existing
`PptxViewer` `hiddenSlideMode` option, whose default remains `'show'`.

The direct reader shows an embedded OLE object, such as an Excel or Graph
chart, as the presentation picture the file stores for it: an OLE shape is a
picture frame whose `pib` names the BLIP to display (MS-ODRAW 2.2.40 and
2.3.23.5), resolved through the shape's `ExObjRefAtom` to the document's
external object list (MS-PPT 2.7.7 and 2.10.1). The object storage is never
read or activated. PowerPoint's own PDF exports show that stored picture
unchanged for embedded objects drawn as content. Icon or thumbnail aspects,
linked objects and ActiveX controls, pictures without a supported BLIP, and
pattern, texture or non-stretched picture fills are rejected instead of being
drawn without them.

PowerPoint displays GIF data that a producer stored in a PNG picture slot,
so the direct PPT reader identifies such a slot by its GIF87a/GIF89a
signature and emits it as `image/gif`; other mismatched content stays
rejected.

Classic linear gradients (MS-ODRAW `msofillShade` and `msofillShadeScale`)
keep their ordered shade colours, signed focus and 16.16 angle. Explicit
shade-array stops at either endpoint take precedence over scalar front or
back colours; a missing final endpoint uses the scalar back colour. Linear,
scaled, two-colour and translucent shades are projected, including on rotated
shapes and inside rotated or flipped groups. Path and title shades, other
shade types, custom fill rectangles and opacity combined with a shade-colour
array fail closed. Twenty controlled Office comparisons covered focus, angle,
endpoint conflicts and leaf flips; the projected positions were within one
DrawingML position unit of Office's serialized integers.

Pattern fills on unrotated shapes become tiled picture fills that follow
PowerPoint's own output: the 8x8 area of the stored 10x10 pattern bitmap,
one pattern pixel per point, white pixels in the fill colour and black
pixels in the background colour. Pattern fills on rotated or flipped shapes,
other pattern bitmap sizes, translucent pattern colours, background
patterns and texture fills stay rejected.

Picture colour settings follow how PowerPoint itself reads the binary
properties when it saves a binary deck as PPTX: "Black and White" becomes
DrawingML `grayscl` plus `biLevel` at 50%, and a transparent colour becomes a
`clrChange` to the same colour with zero alpha. The presentation renderer
applies these blip effects in document order for PPTX files as well.
Brightness/contrast (washout) on a picture becomes DrawingML `lum`: bright is
the brightness over 0x8000, and contrast is k - 1 or 1 - 1/k for the stored
16.16 slope k. PowerPoint writes exactly these values when it saves the
binary deck as PPTX. The renderer's `lum` formula reproduces PowerPoint's PDF
of a gray ramp under a bright × contrast grid, from both PPTX and .ppt, to
within 0.7 of 255. Brightness/contrast combined with a transparent colour or
black-and-white (whose order has no evidence), grayscale or black-and-white
alone, recolouring and adjustments on picture fills stay rejected until
Office output confirms their rendering.

A modern Office-saved PPT can retain paragraph properties in the OfficeArt
`metroBlob` alternative shape XML rather than in its classic text ruler. The
direct PPT source adopts that XML under the rule described with the release
gap inventory below.

## Current implementation boundary

The repository contains only the direct readers; legacy input is never
converted to OOXML. The Rust crate `legacy-office-converter` builds one reader
per feature; this package has `direct-ppt`. Building it for `wasm32` without
it is a compile error. The `inspection` feature adds a native-only PPT
inspection example, and `fuzzing` exposes the fuzz entry points.

The local direct-render survey described below renders every installed
legacy sample through its source and pairs it with the Office-exported PDF.
The corpus is deliberately not redistributed. Broader binary-record coverage,
visual fidelity against Office, fuzzing and resource measurements remain part
of [issue #1472](https://github.com/yukiyokotani/office-open-xml-viewer/issues/1472).

## Best-effort fidelity evaluation

Loading without an error is a smoke test, **not reader completion**. The
target is useful best-effort preservation of the binary input's content and
display, with missing content and visual differences explicitly reported.
Pixel equality is not required for each incremental improvement. Pairing a
legacy file with its original OOXML is useful for investigation, but does not
prove fidelity: saving to an old format can itself change or remove features.
Use Office opening the actual legacy file as the visual reference. Office's
upgraded OOXML is useful for mapping binary records to XML, but Office's own
conversion can change layout and is not an absolute visual oracle. In
particular, rebuilt or down-saved corpus members must not silently be treated
as lossless copies of their original OOXML.

Compare the direct reader's Canvas output with Office-exported PDFs through
the local direct-render survey. Keep renderer self-regression tests against the
previous renderer separate from this fidelity comparison. Neither whole-corpus
Office equality nor Canvas display equality has been reached.

Treat the rendered model as a derived view of the original binary, keep the
binary as the authoritative source, and gate production use on a corpus
representative of the documents being ingested.

## Local direct-render survey

`packages/legacy-converter/tests/survey/ppt.spec.ts` renders each
local private legacy sample through the direct source, on the PPTX package's
own VRT fixture and dev server. Each sample is written beside its
same-named Office PDF export as paired PNGs and a summary. The survey reports
only: it never gates, updates references, or generates OOXML. Run it with an
output directory outside the checkout (`VRT_PORT` serves DOC, `+1` PPT and
`+2` XLS; `LEGACY_CORPUS_FORMATS` and `LEGACY_CORPUS_FILTER` narrow the run):

```bash
LEGACY_CORPUS=1 LEGACY_CORPUS_OUT=/tmp/legacy-survey LEGACY_CORPUS_FORMATS=ppt \
  pnpm --filter @silurus/ooxml-legacy-converter survey
```

Pixel percentages are only a triage signal. For example, a slide can score
above 95% while a chart or autoshape is missing, so review the pairs visually.

## Release gap inventory

The direct-render survey was reviewed visually against the Office PDF
exports, and the findings are grouped here. Counts are local private samples
affected. They record open work, not supported behavior. Omission is
acceptable only where the caller did not enable an opt-in module such as
chartex. Every other gap below is unimplemented behavior or a bug that must
be closed before an experimental release.

| Area | Gap | Samples |
| --- | --- | --- |
| PPT | ~~Only seven MS-ODRAW shape types map to presets~~ 100+ shape types map as PowerPoint converts them, with evidenced adjust formulas (officeart::preset); adjusted callout2/3 families, arrow callouts, curved arrows, ribbons and tall cubes/hexagons/parallelograms still fail closed | several |
| PPT | ~~Native/OLE charts are missing~~ Resolved: embedded OLE objects show their stored presentation picture (bfc835d3) | 3 |
| PPT | ~~Rotation by multiples of 90 degrees and combined flips use the wrong bounds or order~~ Resolved from the 120-case PowerPoint control (aa9dc5c1) | 1 |
| PPT | ~~Slide gradient backgrounds~~ linear/scaled/two-colour/translucent shades resolved (95b74d19); path (5, 6) and title (8) shades now fail closed. Bullets resolved through master levels; letter spacing, shrink-to-fit and per-paragraph indents come from adopted alternative shape XML where it agrees with the binary | several |
| PPT | ~~Gradients on rotated shapes (or inside rotated/flipped groups) are replaced by the solid fill colour~~ Resolved (ef41f03a) | several |
| PPT | Custom geometry with per-path fill/stroke flags is rejected; the PPTX model has no per-path `fill`/`stroke` (ECMA-376 §20.1.9.15), a generic PPTX gap | 1 |
| PPT | ~~Unmapped shape types are dropped silently~~ Now rejected | several |
| PPT | ~~Picture brightness/contrast (washout)~~ projected as `lum` from the gray-ramp control; pattern fills (including on rotated shapes) are projected; pattern fills on flipped shapes, texture fills, and OLE icons, links and controls are rejected | several |
| PPT | Implicit paragraph margin/indent and percentage spacing are rejected | 12 of 34 load failures |
| PPT | Alternative shape XML (metroBlob) that cannot be verified against the binary shape fails closed: placeholders (no slide layout), a preset against freeform geometry, and fills stated in non-comparable forms | first error of 14 of 34 (10 previously loaded) |

PowerPoint 2007+ also stores a `metroBlob` (MS-ODRAW 2.3.4.41, an OPC
package with the shape's DrawingML, `drs/shapexml.xml`) on most shapes. The
specification says it SHOULD be ignored; implementation note 32 says Office
2007 and 2010 use it. PowerPoint 16 controls settle how it is used:

- With only the alternative XML edited (binary byte-identical, the package
  rebuilt at the same length), PowerPoint's PDF renders the edited XML: a card
  fill changed in the XML only is drawn in the new color, and character
  spacing raised in the XML only is drawn raised.
- With a shape's binary adjust and fill edited but its alternative XML kept,
  PowerPoint's PDF follows the binary, identically to a copy whose XML was
  removed.
- A corpus deck whose alternative states 10.5 pt text over a binary 10 pt run
  renders without the alternative's other run properties, while decks whose
  sizes agree render them.

The alternative therefore carries display information the binary lacks, so
the direct PPT source resolves each shape's alternative to one of three
outcomes:

- The package names no alternative part (only the `downRev` checksums), or
  the alternative verifiably disagrees with the binary on a compared
  attribute: the binary projection is used, as PowerPoint does.
- The alternative agrees on every compared attribute: it is adopted. Its
  text characters are masked, so the adopted shape takes its characters,
  transform and identifier from the binary. The XML is parsed by the ordinary
  PPTX shape parser, resolving theme references against the master's
  round-trip theme and color map; no OOXML is generated. Its serialized size
  is charged to the session's model budget.
- Otherwise agreement cannot be established and the presentation fails
  closed as unsupported input: an oversized, over-budget, duplicated or
  unreadable blob, package or round-trip theme; an alternative part that is
  not a shape or connector (pictures, groups, SmartArt, ink); a placeholder,
  whose alternative inherits from a slide layout the binary file does not
  have; relationship references; and any compared attribute the two forms
  state in ways the reader cannot equate.

The compared attributes are those the controls cover, each in one unit
system. Presets compare by name and by adjust values within the range the
binary inputs' rounding allows (a whole 21600-based value on an anchor that
rounds PowerPoint's extent to one master unit), an omitted value standing
for the preset default of ECMA-376 `presetShapeDefinitions.xml`; a preset
against custom geometry is not comparable, because PowerPoint saves a preset
without an MS-ODRAW shape type as a freeform. Custom geometry compares path
by path in normalized path coordinates, including the authored per-path fill
and stroke flags, after dropping a closing line to the subpath start that the
binary stores explicitly. Position, size, rotation and flips compare in the
binary anchor's unit (master units, or a group's unscaled child units) within
one unit. The recorded fill compares before PowerPoint's open-path display
rule: solid colors within one unit per channel (independent rounding of a
theme color transform), gradients and patterns only when identical within
their fixed-point rounding. Run font size, bold and italic compare where both
state them, and paragraphs, runs and line breaks must line up at equal UTF-16
lengths. XML 1.0 cannot carry the vertical tab that breaks a line in a binary
paragraph; the masked alternative states it as one masked character of the
enclosing run, which becomes a line break in the adopted shape. The downrev
checksums beside the XML are not recomputable (they are not checksums of the
binary records), so structural agreement stands in for them; a binary edit
that preserves every compared property would still adopt a stale XML.
Glyph shadow and emboss, which the binary cannot express, reject a shape only
when its alternative XML is not adopted.

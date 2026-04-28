# Changelog

All notable changes to @silurus/ooxml are documented here. The project follows
semantic versioning; minor releases add spec-compliant features or behavior
changes that remain compatible with existing API surfaces.

## 0.21.0 — 2026-04-29

Patch-level release. XLSX overflow rendering fixes and refreshed project icon.

- **XLSX engine** (`packages/xlsx`):
  - **`centerContinuous` overflow**: when text was wider than the
    selection range, it was clipped at the range boundary. Excel keeps
    overflowing the text symmetrically into adjacent empty cells, so the
    renderer now extends the draw rect using the same logic as
    `center` alignment (PR #148, ECMA-376 §18.18.40).
  - **Neighbour fill no longer overpaints overflow text**: cell text is
    drawn in a deferred second pass after every cell's background, so
    e.g. a left-aligned overflow stays visible on top of an adjacent
    cell with a non-default fill — matching Excel's z-order (PR #148).
- **Assets**:
  - Refreshed `docs/images/icon.png` and the VS Code extension icon
    (PR #149).

## 0.20.0 — 2026-04-28

Minor release. Built-in PowerPoint table styles and additional XLSX cell alignments.

- **PPTX engine** (`packages/pptx`):
  - **74 built-in table style presets**: all PowerPoint ribbon table styles
    (Themed 1/2, Light 1/2/3, Medium 1/2/3/4, Dark 1/2 × accent variants)
    are now resolved from a hard-coded GUID catalog when no
    `ppt/tableStyles.xml` definition is present. Fills and borders are
    computed directly from the presentation theme (PR #146,
    ECMA-376 §19.3.1.39; GUID catalog from LibreOffice MPL 2.0 source).
- **XLSX engine** (`packages/xlsx`):
  - **Cell alignment — `centerContinuous`, `fill`, `distributed`,
    `justify`, `readingOrder`**: five additional `<xf alignment>` modes
    are now parsed and applied during cell rendering (PR #145).

## 0.19.0 — 2026-04-27

Minor release. Spec-faithful improvements to scatter chart rendering
on the XLSX side and an Excel-matching font cascade for cell text.

- **PPTX engine** (`packages/xlsx`, `packages/core`):
  - **Scatter X-axis position**: parse `<c:catAx><c:crosses>` /
    `<c:crossesAt>` and draw the X-axis line at the resolved Y
    coordinate (`autoZero` → `toY(0)`). The Vertex42 "Project
    Timeline" template puts milestones above and tasks below the
    timeline ruler — that finally renders correctly because the
    ruler now sits at y=0 instead of the bottom of the plot rect
    (PR #140).
  - **Bold and font sizes from `<c:txPr>` / `<a:defRPr>`**: chart
    title, both axis tick labels, series-level `<c:dLbls>`
    defaults (size 1200/1400 = 12 / 14 pt), and per-idx `<c:dLbl>`
    rich text now honor `b="1"` and `sz="..."` (PR #140).
  - **Axis line styling**: `<c:catAx | valAx><c:spPr><a:ln>`
    resolved color and width are applied to the axis line stroke
    (sample-26's 5 pt 50 % gray timeline ruler now renders as
    intended) (PR #141).
  - **Auto-derived axis range snaps to nice round values**: when
    `<c:scaling><c:min/max>` aren't set, we expand both ends to a
    multiple of `niceStep`. Sample-26 X-axis now spans
    2018/2/19 .. 2018/10/27 in 50-day steps, matching Excel
    (PR #141).
  - **Tick marks honor `<c:majorTickMark>` / `<c:minorTickMark>`**:
    the XLSX adapter was hard-coding `'cross'`; replaced with the
    parsed value (default `'out'` per ECMA-376 §21.2.2.49). Scatter
    renderer now actually calls `drawAxisTick` at every major tick
    on both axes; tick stroke inherits the axis line color and
    width (PR #142).
- **XLSX cell font** (`packages/xlsx`):
  - Calibri-styled cells fell back to system Arial / Helvetica on
    macOS / Linux, which is ~10–15 % wider per character at every
    weight × size — visible on `demo/sample-1` where the B2 title
    overflowed past column F instead of stopping inside column E.
    Add an opt-in `useGoogleFonts` to `XlsxWorkbook.load` (and
    forwarded from `XlsxViewer`'s constructor option) that loads
    [Carlito](https://fonts.google.com/specimen/Carlito) and
    [Caladea](https://fonts.google.com/specimen/Caladea) — Google's
    metric-compatible substitutes for Calibri and Cambria. With the
    substitutes loaded, advance widths match Excel and column
    layout shows the title fitting inside the same cells Excel
    does. Default off — zero third-party requests until the host
    opts in (PR #143).

The unified `ChartModel` and `ChartSeries` gained matching optional
fields. PPTX charts continue rendering unchanged — the PPTX parser
will pick up the same fields in a follow-up release.

## 0.18.2 — 2026-04-27

Patch release. Sparkline visual polish.

- **Core sparkline** (`packages/core/src/sparkline/renderer.ts`):
  - Vertical padding inside the cell switched from a 2 px cap to a
    proportional 20 % of the cell height (with a 2 px floor). The peak
    / trough of the line and high / low marker dots no longer overlap
    the row separators, and the breathing room stays consistent across
    zoom levels.

## 0.18.1 — 2026-04-27

Patch release. One sparkline correctness fix.

- **Core sparkline** (`packages/core/src/sparkline/renderer.ts`):
  - `computeFlagged` now marks **every** point tied for the high or low
    value, not just the first occurrence. Excel does the same. Visible
    on `private/sample-7.xlsx` Q10 (3 non-zero leading values + 9 zeros
    under `low="1"`): all 9 zero points are now dotted. Other
    highlights (`first` / `last` / `negative`) are unchanged — only
    `high` / `low` had the tied-value case.

## 0.18.0 — 2026-04-27

Minor release. Excel scatter / bubble charts gain a spec-faithful set of
features so Gantt-style "Project Timeline" templates render against the
reference. Sparkline release reused — no docx/pptx behavioral changes.

- **PPTX engine** (`packages/xlsx`, `packages/core`):
  - Per-point `<c:dPt>` overrides (color, marker shape / size / fill /
    line) — ECMA-376 §21.2.2.39.
  - `<c:marker>` symbol / size resolution with all 10 ECMA shapes
    (circle / square / diamond / triangle / x / plus / star / dot /
    dash / picture). `picture` falls back to circle pending image-marker
    resolution.
  - `<c:dLbl>` per-point custom data labels (§21.2.2.45) with rich-text
    flattening and `<a:fld type="CELLRANGE">` substitution from the
    series' `<c15:datalabelsRange>` cache. Position (`l`/`r`/`t`/`b`/
    `ctr`/`outEnd`) is honored.
  - `<c:errBars>` (§21.2.2.20) X / Y direction with all five
    `errValType` modes — `cust` (cell-range), `fixedVal`, `stdErr`,
    `stdDev`, `percentage`. Cust values can be signed (Vertex42 Gantt
    uses negative minus values to flip stems toward the X-axis).
    Dashed strokes and end caps via `<a:prstDash>` / `<c:noEndCap>`.
  - `<c:title><c:layout>` and `<c:plotArea><c:layout>` manual layout
    (§21.2.2.27) — when present, used directly for absolute placement.
  - Scatter's dual `<c:valAx>` blocks are now disambiguated by
    `<c:axPos>` so X-axis (`b`/`t`) and Y-axis (`l`/`r`) settings end
    up in their respective slots. `<c:scaling><c:min/max>` and
    `<c:numFmt>` are picked up per axis. `<c:delete val="1"/>` on
    either axis hides it correctly.
  - `formatChartValWithCode` recognises date format codes (m/d/y/h/s
    outside quotes) and routes through a new `formatExcelDate` so
    scatter X-axis tick labels for date series come out as
    `2018/4/12` instead of raw serial numbers.

The unified `ChartModel` / `ChartSeries` gained the corresponding
optional fields. PPTX charts continue rendering unchanged — the PPTX
parser hasn't been updated to populate them yet (follow-up).

## 0.17.0 — 2026-04-27

Minor release. Handwritten ink strokes now render via PowerPoint's
rasterized fallback, plus rendering accuracy fixes for the PowerPoint
engine.

- **PPTX engine** (`packages/pptx`):
  - `mc:AlternateContent` now walks `mc:Fallback` when `mc:Choice`
    produces no output (previously Choice was always taken and Fallback
    silently discarded). PowerPoint embeds ink / handwriting as
    `p:contentPart` (InkML) inside Choice with a rasterized `p:pic`
    inside Fallback; this restores those strokes.
  - When the Choice subtree is an ink `p:contentPart`, the fallback PNG
    is rendered at its natural pixel size centered in the bounding box
    rather than always stretched. Empty / single-tap strokes (whose
    fallback PNG is only a few pixels) no longer blow up into blocky
    artifacts. Visible strokes are unaffected.
  - Fix preset-shape arc visual-to-parametric angle conversion for
    non-square shape boxes (ECMA-376 §20.1.9.18 `<a:arcTo>`). The
    path-executor was using canvas-scaled radii where path-local radii
    were required, skewing every arc segment when `sx ≠ sy`. Visible
    on `cloudCallout` placed in a landscape box, where the inner cloud
    detail arcs were misaligned even though the outline looked plausible.
  - `PptxViewer` wrapper sets `vertical-align: top` so the inline-block
    line-box descender (~6 px on default font metrics) no longer leaks
    the host container's background through below the canvas.
- **VS Code extension** (`packages/vscode-extension`):
  - Downsize `icon.png` to 512×512 to match the practical Marketplace
    icon range. No functional change.

## 0.16.1 — 2026-04-26

Patch release. Project icon refresh.

- **Project icon**: replace `docs/images/icon.png` with a refreshed master
  and resync `packages/vscode-extension/icon.png` from it. README and the
  VS Code Marketplace listing now show the updated artwork.

## 0.16.0 — 2026-04-26

Minor release. Audio / video–embedded files now open in the VS Code
extension and play back interactively.

- **VS Code extension** (`packages/vscode-extension`):
  - Replace the `Array.from(bytes)` + `webview.postMessage` data path with
    `webview.asWebviewUri()` + `fetch().arrayBuffer()` in the webview. The
    previous IPC route serialized file bytes as a JSON number array, which
    hung the spinner indefinitely on media-embedded pptx (50–200 MB) and,
    after a stop-gap that sent `Uint8Array` directly, returned zero-byte
    buffers (`Could not find EOCD` for every file) because VS Code's
    webview `postMessage` does not reliably structured-clone typed arrays.
    Same approach the bundled PDF viewer uses; native binary path, no IPC
    size or type cliff.
  - Switch the scroll-stack pptx renderer from `PptxPresentation.renderSlide`
    to `presentSlide` so embedded audio / video become clickable with a
    canvas-native play / pause / progress bar.
- **PPTX engine** (`packages/pptx`):
  - `presentSlide` now forwards `opts.onTextRun` to its inner `renderSlide`
    call so the transparent text-selection layer keeps working when
    interactive playback is enabled. The same bug existed when calling
    `PptxViewer` with both `enableMediaPlayback: true` and
    `enableTextSelection: true`; fixed at the engine level so both paths
    benefit.
  - `createPresentationHandle` now skips its `requestAnimationFrame` loop
    and pointer wiring for slides without media, so a 50-slide deck no
    longer spawns 50 idle animation loops.

## 0.15.2 — 2026-04-26

Patch release. Project icon refresh.

- **Project icon**:
  - Adopt a new high-resolution master at `docs/images/icon.png` (2048×2048)
    and reference it from the root README.
  - **VS Code extension** (`packages/vscode-extension`):
    - Replace the previous 128×128 Marketplace icon with the new master so the
      Marketplace listing renders crisply on retina displays.
    - Show the icon at the top of the extension README (via the GitHub raw
      URL — Marketplace ignores relative image paths).
    - Add a `cp ../../docs/images/icon.png ./icon.png` step to
      `vscode:prepublish` so the bundled `.vsix` icon is always re-synced from
      the master at publish time and never drifts.

## 0.15.1 — 2026-04-26

Patch release. Mobile UX fix for the XLSX viewer plus Storybook tidy-up.

- **XLSX viewer** (`packages/xlsx`):
  - Distinguish tap from swipe on touch / pen input. `pointerdown` no longer
    commits a cell selection for non-mouse pointers; the gesture is buffered
    and only commits on `pointerup` if the pointer stayed within an 8 px slop.
    A swipe to scroll on a phone or tablet now leaves the selected cell
    alone. Mouse input is unchanged so drag-to-extend keeps working.
- **Storybook**:
  - Drop the `Selectable — file upload` / `Selectable — sample-1.xlsx`
    stories. The public demo already exercises cell selection, so the stories
    were redundant.

## 0.15.0 — 2026-04-25

VS Code extension polish + selection overlay accuracy fix. No new format
support compared to 0.14.1; library packages are bumped to 0.15.0 so the
tag-driven CI keeps the npm versions in sync with the VS Code Marketplace
release.

- **VS Code extension** (`packages/vscode-extension`):
  - Add Marketplace icon (`icon.png`, 128×128) and wire it up via `package.json#icon`.
  - Shorten `displayName` from `Office Viewer — DOCX, XLSX, PPTX` to `Office Viewer`; supported formats remain in the description and feature list.
  - Replace the plain text loading status with a CSS-only spinner (#107).
  - Center the loading spinner and error status on the viewport.
  - Open documents in the currently focused column instead of forcing a split (#114).
- **Viewer / selection overlay** (`packages/docx`, `packages/pptx`, VS Code webview):
  - Carry the canvas `ctx.font` shorthand (font-family / weight / style) through `TextRunInfo` / `DocxTextRunInfo` and apply it on the transparent selection `<span>` so its width tracks the drawn glyphs. Previously the overlay relied on a fallback font, which drifted at the trailing edge of European text. Kerning / ligatures are intentionally left at the browser default to match canvas behavior.
- **Docs**:
  - Add a forward-looking note explaining the dual-layer (canvas + transparent DOM) selection architecture as a deliberate stop-gap, with a reference to the WICG `html-in-canvas` `drawElement` API as the planned unified replacement.
  - VS Code extension README: 3-column screenshot table for the Marketplace listing.
  - Standardize format ordering across READMEs as DOCX → XLSX → PPTX.
  - Add Marketplace badges to the root README.

## 0.14.1 — 2026-04-25

VS Code Marketplace metadata fix. The Marketplace `vsce publish` of 0.14.0
failed because the extension `name` (`ooxml-viewer`) collided with an existing
listing. This release renames the extension and broadens the displayName so
that users searching the Marketplace for "Office", "XLSX", "DOCX", or "PPTX"
discover it.

- **Rename** (`packages/vscode-extension`):
  - `name`: `ooxml-viewer` → `office-open-xml-viewer`
  - `displayName`: `OOXML Viewer` → `Office Viewer — XLSX, DOCX, PPTX`
  - `description`: emphasizes Office file support and the local-only privacy
    posture for Marketplace searchers.
- Library packages (`@silurus/ooxml{,-pptx,-xlsx,-docx,-core}`) are bumped to
  0.14.1 to keep tag-driven CI in sync; **no code changes for the libraries.**

## 0.14.0 — 2026-04-25

VS Code extension UX overhaul. The `.docx` and `.pptx` editors switch from a
prev/next pager to a **continuous scroll-stack** that renders every page or
slide at once with a transparent text layer (PDF.js style). The Webview chrome
now follows the active VS Code theme (light / dark / high-contrast) via
`--vscode-editor-background` and `--vscode-foreground`. The Marketplace README
gains screenshots and a privacy statement asserting zero network access.

### vscode-extension

- **Scroll-stack viewer** (`packages/vscode-extension`) — replaces the
  page-by-page navigation for docx/pptx. Every page/slide is rendered
  vertically with its own transparent text layer; selection and copy work
  across the whole document.
- **Theme-aware backgrounds** — body/foreground driven by VS Code CSS
  variables; the chrome around documents follows the active theme without
  hardcoded fallbacks.
- **CSP + handshake** — workers accept a `data:`-URL wasm asset (decoded
  inside the worker) so the Webview CSP can stay strict; the editor waits for
  a `webview-ready` ping before posting the file payload, fixing an init
  ordering race.
- **Marketplace README** — adds screenshots (absolute raw URLs so they render
  on the Marketplace page), a "Privacy & Security" section, and reflects the
  scroll-view UX.

### tooling

- **pptx wasm script** (`packages/pptx`) — switch to `wasm-pack build
  --out-dir`, matching xlsx/docx, so the generated `pptx_parser.d.ts` is
  written into `src/wasm/`. Resolves the CI `Build library packages` failure
  where per-package `tsc --build` errored on `worker.ts` with TS7016.
- `.gitignore`: exclude `*.vsix` build output.

## 0.13.0 — 2026-04-25

UX and tooling release. The core viewer packages gain **text and cell
selection** (PDF.js-style transparent overlay so the browser's native
selection/copy work on top of Canvas). Two new companion packages ship
alongside: a **VS Code extension** (`ooxml-viewer`) that registers custom
editors for `.xlsx` / `.docx` / `.pptx`, and a **Rust MCP server**
(`ooxml-mcp-server`) that exposes the parsers as structured tools for AI
agents. No rendering-fidelity changes.

### viewer UX

- **Text selection overlay (pptx/docx/xlsx)** — each viewer now emits an
  `onTextRun` stream from the renderer and mounts an absolute-positioned
  `<span>` per text run above the canvas with `color: transparent`. The
  browser's native selection, copy, and `::selection` styling all work
  against the overlay, so users can select, Ctrl+C copy, or drag text
  exactly as they would in a DOM-rendered document. pptx handles rotated
  and vertical text via `transform: rotate(...)` on the overlay spans.
- **xlsx cell selection** (`packages/xlsx`):
  - `getCellAt(clientX, clientY)` on `XlsxViewer` hit-tests canvas
    coordinates to row/col addresses (respects merged cells and freeze
    panes).
  - Four selection modes: single cell, range, row (click row header),
    column (click col header), all (corner click). Drag to extend.
    Shift+click extends from the current anchor.
  - `Ctrl+C` copies the selected range as tab-separated text (TSV) to the
    clipboard, mode-aware (full row → entire row; range → block).
  - `onSelectionChange` callback on `XlsxViewerOptions`; `selection`
    getter on the viewer. New exports: `CellAddress`, `CellRange`,
    `SelectionMode`, `TextRunInfo`.
  - Selection overlay clamps to the header/freeze-pane boundaries so the
    highlight doesn't bleed over the sticky row/column bands.

### New package: VS Code extension (`packages/vscode-extension`)

- Registers `CustomEditorProvider` for `.xlsx`, `.docx`, and `.pptx`, so
  double-clicking an Office file in the VS Code explorer opens it in the
  same Canvas viewer used by the Storybook demo.
- Webview bundles the existing `XlsxViewer` / `DocxViewer` / `PptxViewer`
  classes; selection events can be relayed to the extension host via
  `acquireVsCodeApi().postMessage()`.

### New package: Rust MCP server (`packages/mcp-server`)

- Exposes the existing xlsx/docx/pptx parsers as an MCP server so agents
  (Claude, Copilot, Codex, …) can query OOXML files without shelling out
  to `unzip` + ad-hoc Python. Structured tools include
  `xlsx_get_cell_range`, `xlsx_get_formulas`, `docx_get_structure`,
  `docx_get_tables`, `pptx_get_slide_structure`, and format-specific
  search helpers.
- Built natively from the same Rust crates (the `rlib` output of
  `packages/{xlsx,docx,pptx}/parser`), so the parser logic is shared
  with the browser build one-to-one.

## 0.12.0 — 2026-04-25

xlsx fidelity release focused on sample-1 ("Holiday shopping budget") and
sample-10 ("Calendar"): static pivot/table **slicers** now render from the
Office 2010 extension, **chart data labels** honor per-series `txPr`
(white-on-bar) and the `<c:dLblPos>` / `<c:numFmt>` / `<c:gapWidth>` /
`<c:overlap>` chart attributes, and several conditional-formatting / text
layout fixes land for the calendar sample.

### xlsx

- **Static slicers** (Office 2010 extension `x14:slicerList`, §A.5): parse
  `xl/slicers*.xml` + `slicerCaches/` and render the slicer button array
  with its header and theme-resolved accent fill. Slicers for pivot and
  table sources both lay out correctly; "in-filter" vs "out-of-filter"
  button colors come from the slicerStyle dxfs.
- **Chart bar gap + overlap** (§21.2.2.13 `c:gapWidth`, §21.2.2.25
  `c:overlap`): bar cluster geometry now uses the spec formula
  `clusterWidth = barW · (1 + (N-1)·(1-overlap/100) + gapWidth/100)` so
  paired bars in a two-series chart show the expected gap between
  category clusters instead of flush-packed bars.
- **Chart data-label position and number formats** (§21.2.2.16
  `c:dLblPos`, §21.2.2.21 / .35 / .37): value-axis tick labels honor
  `<c:valAx><c:numFmt>`, data labels honor `<c:dLbls><c:numFmt>` with a
  per-series `<c:val><c:numRef><c:formatCode>` fallback, and labels
  render at the requested position (`inBase` / `inEnd` / `ctr` /
  `outEnd`) with collision-safe placement on horizontal bars.
- **Per-series data-label font color** (§21.2.2.47 `c:ser/c:dLbls`):
  Excel frequently writes the label `txPr` (including `schemeClr
  val="bg1"` → white) on each series rather than on the chart-level
  `<c:dLbls>`. The parser now falls back to the first series's dLbls
  when the chart-level block omits the color, fixing white-on-bar
  labels.
- **Horizontal bar series ordering** (§21.2.2.28 `c:order`,
  §21.2.2.40 `c:delete`): series are sorted by their declared order
  and the visual stack is reversed for horizontal bars so the first
  series appears on top (matches Excel). `<c:catAx><c:delete val="1"/>`
  and `<c:valAx><c:delete>` hide the corresponding axis band, freeing
  padding for the chart itself.
- **Pie/doughnut, radar, waterfall data-label formats**: the same
  `valAxisFormatCode` / `dataLabelFormatCode` plumbing flows through
  non-bar renderers, so value labels on those types pick up the file's
  Excel number-format code (e.g. `¥#,##0.00`).
- **Transparent chart space + theme-palette series colors**
  (§21.2.2.39 `c:chartSpace/c:spPr`): `<a:noFill>` on the chart space
  keeps the underlying cell grid visible behind the chart (Excel's
  default). Series `<c:spPr>` with `<a:schemeClr val="accent1"/>` etc.
  now resolves against the file theme instead of falling through to
  the renderer's built-in palette.
- **Legend manual layout** (§21.2.2.31 `c:legend/c:manualLayout`):
  absolute `x`/`y`/`w`/`h` placement fractions override the default
  side-of-plot legend rectangle while `legendPos` still chooses which
  side of the plot gets the reserved band.
- **dxf numFmt override from conditional formatting** (§18.3.1.10
  `dxf/numFmt`): `cellIs` / `top10` / `aboveAverage` etc. CF rules that
  point to a dxf with a `<numFmt>` now apply that format code when the
  rule matches, not just the fill and font overrides.
- **dxf patternType=none as explicit fill clear**: treats `<patternFill
  patternType="none"/>` inside a dxf as an explicit override that
  unsets the base cell fill, not as "inherit base fill". Matches Excel
  UI where the CF explicitly removes the background.
- **4th format-section (text) honored** (§18.8.30, `;;;` idiom):
  `#,##0;[Red](#,##0);0;@` now applies the fourth section to text-typed
  cells; a `;;;` code correctly hides both numeric *and* text cells.
- **`notContainsBlanks` conditional formatting** (§18.18.15
  `ST_CfType`): the opposite of `containsBlanks`; rules of this type
  now paint non-empty cells instead of silently skipping all cells.
- **`<xdr:grpSp>` custom geometry** (§20.5.2.17): group-shape children
  with `<a:custGeom>` inherit the group's frame transform, so grouped
  freeform icons draw at the correct position and scale.
- **Japanese calendar date format `ge.m.d`** (§18.8.30): era-prefixed
  numeric dates (`R7.4.25`) render alongside the existing era name /
  era year / weekday codes landed in 0.10.0.
- **Image in grouped anchor** (`<xdr:grpSp>` + `<xdr:pic>`): pictures
  nested inside a group anchor no longer drop; the group transform is
  applied to the embedded image frame before rendering.
- **CJK wrap on wrapText cells**: break opportunities between Kanji /
  Hiragana / Katakana characters are recognized when `wrapText="1"`
  is set, matching Excel's line-break behavior for Japanese text.
- **CF over empty cells**: rules that previously required a cached
  `<v>` now also evaluate empty cells against the `containsBlanks` /
  `notContainsBlanks` / text operators.
- **Scroll-flicker fix**: the virtual-scroll frame awaits the next
  animation tick before clearing the canvas, eliminating the flash
  of blank cells during fast scroll.
- **`ShapeGeom::Image` JSON field**: serialized as `dataUrl` (camelCase)
  to match the rest of the parser's JSON surface; fixes shape-image
  rendering in downstream renderers that weren't converting the
  snake_case variant.

## 0.11.0 — 2026-04-22

xlsx fidelity release focused on sample-9 ("Gift budget and tracker"):
stacked combo charts keep their stacking, chart series honor theme
accent colors, custom `<tableStyle>` elements actually style their
cells, and `cellIs` conditional-formatting rules match text operands.

### xlsx

- **Stacked combo charts** (§21.2.2.17): locking `grouping` once a
  non-line series sets it prevents a trailing `lineChart grouping=
  "standard"` from overwriting the bar's `stacked` / `percentStacked`,
  so bar+line combos keep stacked bars.
- **Chart series `<a:schemeClr>` resolution** (§21.2.2.35 `c:spPr`):
  series colors declared as `accent1`..`accent6` / `dk*` / `lt*` are
  resolved against the file's theme color table instead of falling
  back to palette defaults.
- **Custom `<tableStyle>` elements** (§18.8.40): parse `wholeTable`
  and `headerRow` dxf indices from `xl/styles.xml/tableStyles`, then
  overlay the resolved dxf fill, font color, and horizontal / vertical
  borders on top of cells. Built-in style names keep the existing
  accent-based renderer unchanged. `Border` gains `horizontal` /
  `vertical` to carry the inner-rule edges emitted only by tableStyle
  dxfs.
- **Text operands in `cellIs` CF rules** (§18.18.15 `ST_CfOperator`):
  `cellIs` previously only evaluated numeric cells, so text rules like
  `equal "Birthdays"` silently skipped every non-numeric row. Now
  parses each `<formula>` as a quoted string literal or number, and
  compares case-insensitively for equal / notEqual / containsText /
  notContains / beginsWith / endsWith / between / notBetween.

## 0.10.0 — 2026-04-22

xlsx number-format and volatile-function release. Cells with `TODAY()` /
`NOW()` formulas now show today's date at render time instead of the
cached `<v>` from when the file was last saved, and the format-code
renderer gains Japanese weekday / imperial era support plus several
internationally important codes (elapsed time, literal preservation,
scientific notation).

### xlsx

- **Volatile formula recompute** (§18.3.1.40): the parser now carries
  each cell's `<f>` text, and the renderer detects `TODAY()` / `NOW()`
  and substitutes the live serial before formatting. Dates no longer
  appear frozen to the file's last-save date.
- **Japanese weekday format codes** (§18.8.30): `aaa` → 水, `aaaa` →
  水曜日. Detected as date formats even without a `y`/`d` specifier.
- **Japanese imperial era format codes** (§18.8.30): `g` / `gg` / `ggg`
  render the era name (R / 令 / 令和) and `e` / `ee` / `r` / `rr` render
  the era year. Era table covers Meiji through Reiwa; no runtime
  dependency added.
- **Elapsed-time brackets** `[h]` / `[m]` / `[s]` (§18.8.30): render the
  full duration instead of wrapping at 24h / 60m / 60s, so a 54-hour
  value formatted `[h]:mm` reads `54:00`.
- **Literal text preservation in number formats**: quoted strings
  (`"$"#,##0.00`) and backslash-escaped characters (`\$#,##0`), as well
  as non-placeholder currency glyphs like `¥` / `€`, are now kept around
  the formatted number instead of being stripped.
- **Scientific notation** `0.00E+00` / `0.00E-00`: honors the exponent
  sign placeholder and pads the exponent to at least two digits.

## 0.9.0 — 2026-04-22

Focused xlsx release: conditional formatting now evaluates formula-based
rules, resolves defined names, overlays `<dxf>` borders per edge, and
honors Excel's `x14:dataBar@gradient="0"` for solid bars. The CF formula
evaluator is broadened to cover the functions most commonly used in
`expression` rules.

### xlsx

- **Conditional formatting — `expression` rules** (§18.3.1.10): the
  formula is tokenized, references are shifted by the sqref anchor, and
  the AST is walked to a boolean. `stopIfTrue` and rule priority are
  honored so later rules can't mask earlier hits.
- **Defined-name resolution** (§18.2.5): sheet-scoped names used inside
  CF expressions (e.g. `task_start`, `today`) are resolved by inlining
  the formula and shifting embedded relative refs from A1.
- **CF `<dxf>` borders** (§18.8.17): per-edge overlay — a CF rule can
  draw a red left/right stripe without erasing the cell's existing
  top/bottom border.
- **Data-bar gradient flag**: `x14:dataBar@gradient="0"` (living in a
  separate worksheet-level `<extLst>` linked by GUID) now produces a
  solid fill. Previously bars always rendered with a gradient.
- **Data-bar / color-scale theme colors**: `<color theme="…" tint="…">`
  inside `<dataBar>`/`<colorScale>` is now resolved through the workbook
  theme (was srgb/indexed only).
- **`sheetView showGridLines`** (§18.3.1.83): when unchecked in Excel's
  View tab, the default `#d0d0d0` grid lines are no longer drawn.
- **Formula evaluator broadening** (for CF `expression` rules): `A1:B5`
  ranges; `&` concatenation; IFERROR/IFS; type checks (ISTEXT, ISERROR,
  ISNA, …); math (TRUNC, CEILING, FLOOR, MOD, POWER, SQRT, SIGN, EXP,
  LN, LOG10); aggregates (AVERAGE, COUNT/COUNTA/COUNTBLANK, COUNTIF,
  SUMIF, AVERAGEIF with operator-prefixed criteria); text (LEN, LEFT,
  RIGHT, MID, UPPER, LOWER, TRIM, EXACT, FIND, SEARCH, CONCATENATE, T,
  N, VALUE); ROW/COLUMN; date (TODAY, NOW, DATE, YEAR, MONTH, DAY,
  WEEKDAY, with the 1900 leap-year serial compensation).

## 0.8.1 — 2026-04-21

### Infrastructure

- **Demo URL changed to `https://ooxml.silurus.dev`** — custom domain updated
  from `demo.silurus.dev` to `ooxml.silurus.dev` for scalability across future libraries.

## 0.8.0 — 2026-04-21

### Infrastructure

- **Demo URL changed to `https://demo.silurus.dev`** — GitHub Pages now served
  from a custom domain. README and npm `homepage` field updated accordingly.
- Storybook build base path simplified to `/` (was `/office-open-xml-viewer/`
  on CI); `CNAME` file is now written into the artifact on every deploy.

## 0.7.0 — 2026-04-21

Quality pass across pptx shape rendering and chart legends — no new
feature categories, but several existing ✅ features now match the
Office output more faithfully.

### pptx

- **`cxnSp` connectors honor `<p:style><a:lnRef idx="N">`** as a stroke
  fallback (#74). Previously a connector that only declared
  `headEnd` / `tailEnd` on `<a:ln>` (no `solidFill`) rendered invisible;
  the style-level stroke now fills in color and width.
- **`<p:style><a:lnRef>` stroke width resolves from the theme's
  `fmtScheme > lnStyleLst`** for both `<p:cxnSp>` (#74) and `<p:sp>`
  (#76). The previous hard-coded 9525 EMU (0.75 pt) under-weighted
  every idx ≥ 2 stroke — idx=2 is 19050 EMU (1.5 pt) and idx=3 is
  25400 EMU (2 pt) in the Office default theme. Brackets, braces, and
  arcs that inherited the style line now render at the thickness
  PowerPoint shows.
- **`<a:tint>` mixes in linear sRGB** (IEC 61966-2-1) rather than
  straight sRGB (#74). Sampling the PDF export of the reference
  SmartArt arrow (#156082 + tint=60000) yields ~#D1D6DB, which the
  linear-sRGB lerp now reproduces pixel-for-pixel.
- **`bentConnector{2-5}` / `curvedConnector{2-5}` routed through the
  ECMA-376 preset path evaluator** (#74), and `getConnectorAnchors()`
  walks the preset cmd list so arrow heads sit on the true tangent
  angle instead of the bounding-box diagonal.
- **`rtTriangle` prstGeom** (right-angle at bottom-left) gained a
  proper path (#74); previously fell back to `rect`.
- **`adj5`–`adj8` threaded through parser → renderer → preset
  evaluator** (#74) for callouts whose gdLst references them
  (e.g. `accentBorderCallout3`).

### charts

- **`c:legendPos` and marker visibility** now drive legend placement
  and series point rendering across the chart families (#72); radar
  charts also honor the value-axis scale instead of defaulting to
  `0–max`.

### xlsx

- **Data bar conditional formatting** renders with the Excel 2010+
  gradient fill instead of the flat solid color (#73), matching the
  in-cell gradient Excel draws.

### Docs

- README screenshots refreshed for the release.
- CLAUDE.md codifies two workflow rules: squash merges to `main` are
  forbidden (use `--merge` or `--rebase`), and the release process
  (README screenshots + support table + CHANGELOG + version bump)
  is documented as a single PR procedure.

## 0.6.0 — 2026-04-21

### docx

Layout improvements driven by cross-referencing Word's PDF export of
demo/sample-1 with our paginator / line-layout output. Unless noted, the
work lands as strict ECMA-376 reading of the relevant sections — empirical
tolerance knobs were deliberately avoided per the project's spec-first
rule.

- **Line spacing, explicit vs inherited (ECMA-376 §17.6.5 + §17.3.1.33).**
  `line_spacing_explicit` now flows through the style cascade. A paragraph
  whose `w:spacing/@w:line` is inherited only from docDefault snaps to one
  grid pitch per line in a `w:docGrid`-enabled section; a paragraph that
  sets `w:line` on its own pPr or a named style multiplies against the
  pitch. Fixes body labels like ESSAY / BY THE EDITORS advancing at
  `pitch × 1.15` instead of the `pitch` Word uses.
- **Paragraph margin collapsing.** The gap between two paragraphs is now
  `max(prev.spaceAfter, this.spaceBefore)` rather than the sum (CSS-style
  collapsing margins). Matches Word's observed 18 pt gap between
  `after=360` → `before=240` paragraphs.
- **spaceAfter may overflow the bottom margin.** A paragraph fits when
  `y + (h − spaceAfter) ≤ contentH`; the trailing whitespace is suppressed
  at page boundaries. Lets a closing paragraph with a large `after` land
  flush against the bottom margin.
- **Knuth-Plass-style shrink tolerance on wrap-fit.** ECMA-376 doesn't
  prescribe a line-breaking algorithm; we adopt the standard typographic
  policy used by TeX / InDesign / Word — each inter-word space may
  compress by up to 25 % of its natural width when testing fit. Absorbs
  the ~0.1–0.3 px/glyph advance difference between Chromium's canvas
  and Word's internal text layout, so long paragraphs wrap like Word's.
- **Implicit `w:keepNext` on heading paragraphs (w:outlineLvl 0–8).** Word's
  built-in Heading 1–9 styles carry an implicit keepNext even when
  styles.xml omits it; parser now sets `keep_next=true` when a paragraph's
  effective style declares `w:outlineLvl`.
- **Table style `w:pPr` cascade (§17.7.6).** A table's `w:tblStyle` now
  contributes its paragraph formatting to every cell paragraph, resolved
  between docDefault and the paragraph's own style. For the default
  "Table Grid" style (`line=240 auto`, `after=0`), this tightens cell
  line spacing from ~28 pt to ~18 pt per line, matching Word.
- **docGrid per-grid-line computation (§17.6.5).** Parsing
  `w:docGrid/@w:type` and `@w:linePitch` on the section now feeds into
  the line-box formula. Headings authored with oversized `lineRule="auto"`
  values (e.g. `line="1040"` on a 56 pt title) no longer blow up into
  ~300 pt tall lines — they snap to the section's grid pitch times the
  multiplier.
- **Inter-word compression on justified lines.** When canvas measurement
  forces a line slightly over `availW`, the final render compresses
  inter-word spaces (capped at ~¼ of the line's ascent) instead of
  overflowing the right margin.

### Stories / samples

- xlsx viewer: active sheet tab is now visually smaller than inactive
  tabs, which matches the project's layout preference.
- pptx interactive playback: media play / pause badge style unified; the
  story now explicitly opts into `presentSlide` so static rendering and
  playback paths share identical chrome.

### Known limitation

- Word chains `w:keepNext` transitively through "heading cluster"
  paragraphs (kicker label → title) that are not themselves marked with
  `w:keepNext` or `w:outlineLvl`. ECMA-376 §17.3.1.15 defines keepNext as
  1-hop only, so we follow the spec and accept that a kicker paragraph
  can land at the bottom of page N while the matching title heads page
  N+1 in layouts like demo/sample-1 page 3's "FIELD NOTES · CHAPTER THREE".

## 0.5.0 — 2026-04-21

### xlsx

- **Pattern fills (ECMA-376 §18.8.20).** `gray125` and hatch patterns
  (`darkGrid`, `lightGrid`, `darkHorizontal`, `lightHorizontal`, etc.) now
  render as a blend of the cell's `fgColor` and `bgColor`, and the hatch
  varieties draw a repeating `CanvasPattern` tile instead of a flat blended
  color.
- **Gradient fills (§18.8.24).** Parse `<gradientFill>` (linear `degree` +
  path-style bounding box) and render via `createLinearGradient` /
  `createRadialGradient` with multi-stop color interpolation.
- **Comment indicators (§18.7.3).** Commented cells get a small red triangle
  in the top-right corner to mirror Excel's visual cue. Parsed from
  `xl/comments*.xml` and exposed as `Worksheet.commentRefs`.

### docx

- **Page margins respected in pagination.** The paginator's per-paragraph
  height estimator now builds the same `WrapLayoutCtx` as the renderer when
  anchor-image floats are active, so float-aware line wrapping and
  estimation agree. Pages whose content wraps around floats no longer
  overshoot or undershoot the bottom margin.
- **Line-box metrics and vertical centering (§17.3.1.33).** Replace
  per-character `actualBoundingBoxAscent` with font-metric
  `fontBoundingBoxAscent` / `fontBoundingBoxDescent`, so every line in a
  paragraph — and every paragraph that shares a font/size — sits on a
  consistent baseline. `lineBoxHeight` now reads: `auto` = natural × value,
  `exact` = pt × scale, `atLeast` = max(natural, pt × scale). Glyphs are
  centered within the line box (extra spacing split above and below),
  fixing text that previously rendered top-aligned inside wide auto-spaced
  paragraphs.

## 0.4.0 — 2026-04-21

### pptx

- **Audio/video playback.** `PptxPresentation.presentSlide(canvas, index, opts?)`
  returns a disposable `PresentationHandle` that layers the current video
  frame and self-drawn play/pause + seek chrome on top of the statically
  rendered slide. Click on a media element to toggle playback; drag the
  progress handle to scrub. Audio gets a capsule-shaped pill with time on
  the left and a thin seekable bar on the right. `renderSlide` stays pure
  and stateless.
- **Lazy media extraction.** The parse output no longer inlines poster images
  as base64 data URLs; `PptxPresentation.getMedia(path)` fetches bytes on
  demand via a new `extract_media` WASM export. Sample decks with 200 MB
  video now have a <1 KB parse JSON instead of a 16 MB one.
- **Charts, spec-faithful.**
  - Chart space vs. plot area fill distinguished (`<c:chartSpace><c:spPr>`
    separate from `<c:plotArea><c:spPr>`). Transparent outer chart lets the
    slide background show through.
  - Legend visibility driven by `<c:legend>` presence.
  - `<c:crossBetween>` honored (0.5-step category padding for "between").
  - Value-axis line + `majorTickMark` drawn.
  - Title / axis / data-label sizes come from XML `<c:txPr>` `sz` in hpt,
    scaled via `ptToPx = 12700 × slideScale` so fonts track the viewport.
  - Line width and marker radius also scale by pt-per-px.
- **Text rendering.**
  - Line-height maximum is computed from actual run sizes, not placeholder
    `defRPr sz="30000"` prompt markers — fixes 24pt text rendering ~260px
    below its anchor in demo/sample-1 slide 8.
  - Bullet size derives from the first run's font size (ECMA-376
    §21.1.2.4.13) instead of the layout default — fixes em-dash bullets
    overlapping text in demo/sample-1 slide 7.
  - **Text overflow no longer clipped** at shape bounds per §20.1.2.3.6.
- SmartArt connector shapes with `cy=0` no longer inflated by the body-text
  auto-height fallback (horizontal timelines now render horizontal).

### docx

- **Pagination properties (ECMA-376 §17.3.1.14 / .15 / .44).** Parse and
  honor `w:keepNext`, `w:keepLines`, `w:widowControl` through the full
  style cascade (docDefaults → style → pPr). `widowControl` defaults to
  true when absent. `keepNext` now causes a page break before the current
  paragraph when the chain of kept-together paragraphs wouldn't fit.
- **Justified / distribute alignment (§17.18.44).** Inter-word whitespace
  is expanded so `<w:jc w:val="both"/>` paragraphs fill the content width;
  `distribute` also stretches the last line. Previously everything outside
  `right`/`center` collapsed to left-aligned — big visual change for docs
  whose docDefaults declare `both`, which includes most Word-authored
  documents.
- **Line spacing (§17.3.1.33).** `lineSpacingMultiplier` now respects the
  pt value on `atLeast` and `exact` rules (previously both collapsed to
  1.2× font). Decorative titles that encode `lineRule="auto"` with very
  large values (~720+) now render with correct line height.
- **Indentation (§17.3.1.12).** Accept logical `start`/`end` aliases in
  addition to `left`/`right` on `<w:ind>`. `hanging` still wins over
  `firstLine` per spec.
- **rFonts theme references (§17.3.2.26 / §20.1.4.1.14).** theme.xml's
  `<a:fontScheme>` is parsed and rFonts `asciiTheme` / `hAnsiTheme` /
  `eastAsiaTheme` refs are resolved against it at run assembly. Direct
  typeface attributes still take precedence per spec.
- **Default paragraph style.** Fall back to the document's `w:default="1"`
  style ID (e.g. `a`, `標準`) instead of the hardcoded literal `Normal`.
  Matters for contextualSpacing grouping on non-English templates.
- **ST_OnOff (§17.3.2.22).** `bool_prop` now recognises `off` — previously
  interpreted as `true`.
- **Footnote / endnote markers (§17.11.16 / §17.11.7).** Render the
  reference number as a superscript marker inline (previously dropped).
  Full page-bottom footnote layout is deferred.
- **Table cell widths (§17.18.87).** `w:type` now defaults to `dxa` when
  absent; non-dxa types fall back to grid allocation.

### Stories / samples

- Interactive pptx sample story auto-disposes the `PresentationHandle` +
  `PptxPresentation` when its root detaches from the DOM. Storybook
  story-swap no longer leaks playing audio.
- CSS spinner overlay (`createCanvasSpinner`) shows while a sample is
  loading. Both the opinionated `buildViewerUI` and the interactive
  `presentSlide` story use the same helper.

### Known follow-up

- Word's auto-rule rendering for very large multipliers (e.g. `line="640"
  lineRule="auto"` on 28pt headings) still diverges from spec — Word
  Desktop and Word Web themselves disagree here, so we stick to the letter
  of the spec instead of empirical tuning.
- Full bottom-of-page footnote layout.
- Tab alignment variants beyond `pos` (center / right / decimal).
- `cstheme` font axis (only `ascii` / `hAnsi` / `eastAsia` resolve today).

## 0.3.0 — 2026-04-20

DOCX shape rendering (solid / gradient fill, lumMod/lumOff, z-order),
anchor image text wrap (Square + TopAndBottom), default paragraph style
fallback. PPTX placeholder image alpha (`a:alphaModFix`) and master
`txStyles` bold/italic inheritance. Shape helpers extracted to
`@silurus/ooxml-core`.

## 0.2.0 and earlier

See git history.

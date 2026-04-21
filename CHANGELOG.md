# Changelog

All notable changes to @silurus/ooxml are documented here. The project follows
semantic versioning; minor releases add spec-compliant features or behavior
changes that remain compatible with existing API surfaces.

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

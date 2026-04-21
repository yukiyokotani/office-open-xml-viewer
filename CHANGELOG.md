# Changelog

All notable changes to @silurus/ooxml are documented here. The project follows
semantic versioning; minor releases add spec-compliant features or behavior
changes that remain compatible with existing API surfaces.

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

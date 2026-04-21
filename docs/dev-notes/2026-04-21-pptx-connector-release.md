# 2026-04-21 — pptx connector & SmartArt release

Session record for the work that landed as
[PR #74](https://github.com/yukiyokotani/office-open-xml-viewer/pull/74)
(squash merge `f2f82e5`). Captures what changed and what remains open, so a
later session can pick up without re-reading the transcript.

## What shipped

- `cxnSp` connectors now honor `<p:style><a:lnRef idx="N"/>` as a stroke
  fallback. Previously a connector with `<a:ln>` that carried only
  `headEnd`/`tailEnd` (no `solidFill`) rendered invisible — sample-3 slide-17
  was the repro. `parse_connector` merges `<a:ln>` attributes with the
  style-resolved base.
- `lnRef` stroke width resolves from `fmtScheme > lnStyleLst` (theme) instead
  of the hardcoded 9525 EMU. idx=2 is 19050 EMU in the Office theme, not
  9525 — the old value under-weighted idx≥2 strokes.
- `<a:tint>` is applied as a lerp toward white in **linear sRGB**
  (IEC 61966-2-1), not straight sRGB. Sampling the PDF export of the slide-8
  SmartArt arrow (#156082 + tint=60000) gave ~#D1D6DB; linear-RGB lerp
  reproduces that pixel-for-pixel, straight-sRGB lerp produced the more
  saturated #73A0B4 that was previously landing on canvas.
- Preset engine routes `bentConnector{2-5}` / `curvedConnector{2-5}` through
  the ECMA-376 `presetShapeDefinitions.xml` path evaluator, and
  `getConnectorAnchors()` walks the preset cmd list to place arrow heads at
  the correct tangent angle (not the bounding-box diagonal).
- `rtTriangle` prstGeom (right-angle at bottom-left) gained a proper
  implementation; previously fell back to `rect`.
- `adj5`-`adj8` thread through parser → renderer → preset evaluator for
  callouts whose gdLst references them (e.g. `accentBorderCallout3`).

## Issue #68 follow-up verification

[#68](https://github.com/yukiyokotani/office-open-xml-viewer/issues/68) asks
for VRT re-confirmation of title padding on area / pie / doughnut / radar /
scatter renderers after [#67](https://github.com/yukiyokotani/office-open-xml-viewer/pull/67).

- **pptx VRT:** 70/70 pass at the 20 % pixel-diff threshold. No regression,
  but coverage is `demo/sample-1` (lineChart) + `private/sample-2` (barChart)
  only, so the paths changed in #67 are not exercised here.
- **xlsx radar** (`private/sample-1`, *Biodiversity Index* sheet):
  Canvas `fillText` hook confirms *Taxonomic Richness by Region* now renders
  at `bold 18.6667px Calibri` = 14 pt × 4/3, matching the spec-driven size
  #67 introduced. Storybook spot-check shows healthy layout — legend below
  the plot, no title clipping.
- **Unverified:** area / pie / doughnut / scatter. None of the current
  pptx or xlsx samples contain these chart families, and xlsx has no VRT
  harness. A narrower follow-up should wait until either a representative
  sample or an xlsx VRT harness exists.

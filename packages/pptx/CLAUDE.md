# CLAUDE.md

## Principles for Autonomous Work

- From 1:00 AM to 9:00 AM, do not ask the user for confirmation. Asking for confirmation stops the work, so proceed autonomously with everything except destructive operations.
- Work that may proceed without confirmation: code changes, WASM builds, test execution, commits and pushes, Python scripts, and npm scripts.
- Commit and push as appropriate after confirming that the changes are improvements.
- Before running `git push`, set `http.postBuffer 524288000` (large packs can cause HTTP 400 errors).
- Update reference images (`tests/visual/references/`) only when the user explicitly requests it. Never update them automatically.

## Project Overview

This library renders OOXML PowerPoint (`.pptx`) files to browser Canvas.
It consists of a Rust/WASM parser and a TypeScript Canvas renderer.

## Directory Structure

```
pptx-parser/          <- Rust/wasm-pack parser
  src/lib.rs          <- Core OOXML parser; outputs camelCase JSON with serde
  pkg/                <- wasm-pack build output

src/
  wasm/               <- Manual copy destination for pkg/; the application reads from here
  types.ts            <- TypeScript types matching the Rust JSON output one-to-one
  renderer.ts         <- Renders slides with the Canvas 2D API
  index.ts            <- Public PptxViewer API
  worker.ts           <- Calls WASM from a Web Worker

public/sample.pptx    <- Test PPTX (5 slides)

tests/visual/
  visual.spec.ts      <- Playwright visual regression tests
  fixture.html        <- Test HTML (width=1920)
  references/         <- Expected images, slide-1.png through slide-5.png
  screenshots/        <- Screenshots updated on every run
  diffs/              <- Pixel-difference images
```

## WASM Build Procedure (Important)

```bash
cd pptx-parser && wasm-pack build --target web

# Always copy the files to src/wasm/ (these are separate and are not synchronized automatically)
cp pptx-parser/pkg/pptx_parser_bg.wasm pptx-parser/pkg/pptx_parser.js src/wasm/
```

If you forget to copy the files, the old WASM build will continue to be used.

## Storybook

Storybook is unified at the repository root, so do not start it from the package directory.
Run `pnpm storybook` from the repository root to access stories from every package.

## Running Tests

```bash
npx playwright test --reporter=list
# Example: slide 4: match=93.8%  diff=6.2%  (127,638 / 2,073,600 px)
```

## Current Test Results (Session 3, 2026-04-16)

| Slide | match% | Notes |
|-------|--------|-------|
| 1 | 99.6% | |
| 2 | 100.0% | |
| 3 | 99.4% | |
| 4 | 99.0% | |
| 5 | 98.8% | |

## Fixed Bugs (Session 2)

### Tab stops (right-alignment offset for "22%" on slide 4)

- Parse `pPr > tabLst > tab` and store it in `Paragraph.tabStops: TabStop[]`.
- When `layoutParagraph` detects a `\t` token, collect the following text in `tabStop.segments`.
- During painting, right-align the text using `tabAbsX - totalTabW`.

### `grpFill` inheritance (flipped objects are not filled)

- Added `group_fill: Option<&Fill>` to `parse_sp_tree_node` and `parse_shape`.
- Shapes with `spPr > grpFill` now inherit their parent group's `solidFill`.
- The wreath leaves on the award badges on slide 5 (gold award, etc.) are now filled with the accent4 gold color (`#EBC83C`).

## Fixed Bugs (Session 3)

### Group rotation is not applied to child shapes (misaligned wreath-leaf rotation)

- Added a `rot: f64` field to `GroupTransform`.
- Read `rot / 60000` from the `grpSp` `xfrm`.
- `apply_to_transform`: rotate the child's center around the group center using clockwise screen coordinates.
- Correct formula for the child's rotation: `child.rot = group.rot + (group.flipH XOR group.flipV ? -t.rot : t.rot)`.
  - When the group has a net flip (`flipH XOR flipV`), negate `t.rot` because the child's rotation direction is reversed.
  - Simply using `t.rot + group.rot` is incorrect because the direction is reversed when `flipH` is set.
- Added `s.rotation = nt.rot` and `p.rotation = nt.rot` to `apply_group_transform_to_element`; the rotation was previously discarded.

### Incorrect inheritance of layout-placeholder outlines (black title outline)

- A slide layout's `spPr > ln` is an editing-mode indicator and should not be painted.
- Removed the `by_type_stroke` field and `lookup_stroke()`.
- Removed the call to `lph.lookup_stroke()` from `parse_shape`.

### Trapezoid adjustment (corner decoration on slide 5 becomes a triangle)

- OOXML rule: `ss = min(w, h)`, `inset = adj / 100000 * ss`.
- Fix: `const ss = Math.min(w, h); const inset = Math.min(w/2, adj/100000 * ss)`.
- With `adj=99828`, `w=159`, and `h=31`, the inset is `30.95px`, producing the correct trapezoid.

## Remaining Work

### Font sizes are too small on slides 3 and 5

- Placeholder shapes do not inherit font sizes from the slide layout or master.
- Investigate `lstStyle > lvl1pPr > defRPr sz` in `ppt/slideLayouts/` and `ppt/slideMasters/`.
- `parse_text_body` does not yet read defaults from the layout or master.

### Autofit is not implemented

- `bodyPr > spAutoFit` requires the font size to shrink when the text does not fit.
- Currently, the text is only clipped.

### Accuracy of `lumMod`/`lumOff` color conversion

- The approximation of scheme-color modifiers such as `tx2 + lumMod=50000` needs improvement.

### Preset shapes such as `leftBracket`

- Unsupported `prstGeom` values currently fall back to `rect`; slide 5 uses one of these shapes.

## Important Technical Notes

### OOXML Units

- Rotation: `rot / 60000` -> degrees
- Font size: `sz / 100` -> pt (for example, `2400` -> `24pt`)
- Spacing: `spaceBefore` and `spaceAfter` are in hundredths of a point; convert to px with `/ 100 * PT_TO_EMU * scale`.

### Theme Colors (`sample.pptx`)

| name | hex |
|------|-----|
| dk2 / tx2 | #196ECA |
| accent1 | #E46970 |
| accent4 | #EBC83C (gold) |
| accent5 | #00A08C |

### `layoutParagraph` Signature

```typescript
layoutParagraph(ctx, para, maxWidthPx, defaultFontSizePx, defaultColor, scale, marLPx)
//                                                                                ^ Used for tab-stop calculations
```

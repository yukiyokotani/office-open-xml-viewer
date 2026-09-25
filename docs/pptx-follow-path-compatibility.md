# PPTX Follow Path multiline compatibility

Multiline arch and circle WordArt now keeps its lines outside the authored
ellipse, with the correct line order and authored spacing. No migration is
required. Single-line and paired-edge WordArt rendering are unchanged.

## Specification boundary

ECMA-376 Part 1 §20.1.9.19 (`a:prstTxWarp`) describes text-envelope mapping.
The preset definitions supply the authored path and adjustment angles. The
single-edge, multiline placement convention below is **observed PowerPoint
compatibility**, not a normative consequence of the paired-edge mapping rule.

## Office evidence

PowerPoint for Mac 16.112.3 opened a synthetic 38-slide, 76-case matrix and
exported PDF and a separately saved PPTX. The audited shape transforms, warp
adjustments, body settings, paragraph spacing, breaks and run properties were
unchanged after saving. Six additional flat-text control slides isolate line
spacing from the curve transformation. These are local verification artifacts,
not redistributable regression baselines.

The matrix covers:

- One through four lines, paragraphs and manual breaks.
- Arch Up, Arch Down and Circle; default and 135°/225° adjustment angles;
  shape rotations of ±70°.
- 50–200% and 18/36 pt line spacing, paragraph before/after spacing.
- Arial and Times New Roman; 12/24/36 pt and mixed-size lines.
- Widths 200/400/600 px, heights 80/180/320 px, alignment and vertical anchors.
- Unequal line widths, descenders, wrapping, and short/overlength controls.

For clockwise paths, the last baseline is anchored to the authored ellipse;
preceding lines expand outward. For counterclockwise paths, the first baseline
is anchored and subsequent lines expand outward. Direction comes from path
winding, not the screen position or a preset-name rule.

Each line is evaluated on its own concentric ellipse, adding its baseline
outset to both radii. Its natural text width and paragraph alignment are then
mapped against that line's arc length. Reusing the original curve and shifting
glyphs along its normals does not reproduce the aspect-ratio controls.

The flat controls support reusing the renderer's existing natural line-box
policy (1.2 times text size, scaled by percentage spacing, or authored point
spacing) and baseline placement. Paragraph gaps are retained between paragraphs.
The patch removes the unrelated paired-edge vertical-band fraction from
single-edge baseline placement.

## Limits of the compatibility claim

This fixes line ordering, radial placement and spacing; it is not a claim of
pixel-identical Office typography. Native output retains small coordinate and
paragraph-versus-manual-break differences, including up to approximately
1.33 CSS px in equal-size controls. The existing polyline tangent sampling,
font substitutions and glyph metrics remain unchanged. Mixed-font and mixed-size
placement uses the existing line-box approximation, not newly recovered font
shaping metrics.

The overlength controls also expose a pre-existing glyph-shrink discrepancy.
This patch does not change overlength fitting or infer its conditions from a
single case. It does not change paired-edge deformation, vertical text, font
loading, or the shared chart renderer.

## Adversarial review and regression evidence

- Specification: the normative envelope rule and observed single-edge policy
  are explicitly separated. No sample names, path checks, new empirical scale
  factors or changed VRT thresholds occur in the implementation.
- Design: paragraph layout stays in PPTX; ellipse evaluation and arc-length
  mapping reuse the existing pure core geometry. DOCX/XLSX have no integration
  with this PPTX WordArt routine; their shared chart and text code is untouched.
- Resources: the existing line list gains one scalar baseline per line. There
  is at most one fixed-size envelope evaluation per expanded line, outside the
  glyph loop. No new cache, bitmap, worker transfer, asynchronous fan-out or
  document-wide work is introduced.
- Tests: focused cases cover both windings, circles, adjusted angles, manual
  breaks, percentage/point/paragraph spacing, mixed sizes and independently
  expanded ellipses. Test coordinate units now match actual font units.
- Self-VRT: a clean detached previous-renderer checkout at
  `ca5a05083375751760b58be0b7f3a1c79ee37aa1`, with rebuilt WASM and a distinct
  port, supplies the immutable baseline. All 35 private PPTX inputs and all
  361 slides were compared. Exactly one slide changed; previous, candidate and
  Office output were explicitly compared, confirming corrected multiline
  ordering, separation and visibility. The other 360 slides are pixel-equal.
  All 35 complete-corpus worker-parity tests pass. The exact self-VRT command
  deliberately remains red for the one adjudicated intended change; references
  were not updated to conceal it.

# Preset shape definitions

`preset-shape-definitions.xml` is a verbatim copy of LibreOffice's
[`oox/source/drawingml/customshapes/presetShapeDefinitions.xml`](https://github.com/LibreOffice/core/blob/master/oox/source/drawingml/customshapes/presetShapeDefinitions.xml),
licensed under MPLv2.

It is the canonical transcription of the 180-plus preset-geometry entries in
ECMA-376 Part 1 §20.1.9 (DrawingML preset shapes): the adjustment defaults
(`avLst`), derived guides (`gdLst`), and path list (`pathLst`) for each
shape. Rendering these at runtime from the spec's own formulas is what
lets PowerPoint-compatible toolchains (LibreOffice, Aspose, …) reproduce
preset geometry pixel-for-pixel.

## Regenerating `presets.json`

```bash
node scripts/extract-presets.mjs scripts/preset-shape-definitions.xml ../core/src/shape/preset-geometry/presets.json
```

The runtime engine (`packages/core/src/shape/preset-geometry/`) consumes only the generated JSON;
the XML is kept in-tree so anyone can re-run the extractor and audit the
compact output against the source.

# Preset text-warp definitions

`preset-text-warp-definitions.xml` is the ECMA-376 Part 1 §20.1.9.19
`presetTextWarpDefinitions.xml` file. `extract-text-warps.mjs` converts it to
the compact JSON that the shared preset-geometry evaluator reads:

```bash
node scripts/extract-text-warps.mjs scripts/preset-text-warp-definitions.xml ../core/src/shape/preset-geometry/text-warp-presets.json
```

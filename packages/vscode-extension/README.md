# OOXML Viewer for VS Code

A high-fidelity viewer for `.docx`, `.xlsx`, and `.pptx` files — powered by a Rust/WASM parser and an HTML Canvas renderer.

> **Private by design.** All parsing and rendering happens locally inside the VS Code Webview via WebAssembly. **No file contents, no metadata, and no telemetry leave your machine.** The extension makes no network requests.

## Screenshots

### DOCX

![DOCX viewer](https://raw.githubusercontent.com/yukiyokotani/office-open-xml-viewer/main/docs/images/docx.png)

### XLSX

![XLSX viewer](https://raw.githubusercontent.com/yukiyokotani/office-open-xml-viewer/main/docs/images/xlsx.png)

### PPTX

![PPTX viewer](https://raw.githubusercontent.com/yukiyokotani/office-open-xml-viewer/main/docs/images/pptx.png)

## Features

- **DOCX** — Continuous **scroll view** of every page with a transparent text layer (PDF.js-style) — drag to select, copy as plain text.
- **XLSX** — Spreadsheet viewer with cell / row / column / range selection, tab-separated copy (Ctrl+C / Cmd+C), freeze-pane support, and a multi-sheet tab bar.
- **PPTX** — Continuous **scroll view** of every slide with a transparent text layer that handles rotated text boxes correctly.
- **Theme-aware** — Background and foreground follow the active VS Code theme (light / dark / high-contrast).
- **High fidelity** — Charts, conditional formatting, theme colors, custom geometry shapes, and more rendered straight from the OOXML spec.

All three formats share the same Rust parser (`wasm-pack`) for accuracy and speed.

## Usage

Open any `.docx`, `.xlsx`, or `.pptx` file in VS Code — the OOXML Viewer takes over as the default editor for those file types.

If a different editor opens by default, right-click the file → **Reopen Editor With…** → select **OOXML Viewer**, then optionally **Configure default editor** to make it the default.

### Selection & copy

- **DOCX / PPTX** — Drag across rendered text to select, then **Ctrl+C / Cmd+C** to copy as plain text. The transparent overlay matches the canvas glyph positions, so selection feels native.
- **XLSX** — Click a cell to select it, drag for a range, click row/column headers for full-row/column selection, click the corner box for sheet-wide selection. **Ctrl+C / Cmd+C** copies as TSV.

## Privacy & Security

- **Zero network access.** The Webview's Content Security Policy disallows outbound connections to any origin other than the extension itself. There is no analytics, no font CDN, no remote API.
- **Local file I/O only.** The extension reads bytes via `vscode.workspace.fs.readFile` and never writes back — files are opened read-only.
- **Open source.** Source code at [github.com/yukiyokotani/office-open-xml-viewer](https://github.com/yukiyokotani/office-open-xml-viewer).

VS Code's own telemetry is independent of this extension and can be controlled via the `telemetry.telemetryLevel` setting.

## Known Limitations

- XLSX: formula evaluation is not yet supported (raw cached values are shown).
- DOCX: image-anchored float wrap, footnotes, and header/footer rendering may differ slightly from Word.
- PPTX: a small number of obscure preset shapes fall back to a rectangle placeholder.
- Media playback (audio / video) is not supported in the Webview.

## Issues & Contributions

Report bugs or request features at [github.com/yukiyokotani/office-open-xml-viewer/issues](https://github.com/yukiyokotani/office-open-xml-viewer/issues).

## License

MIT

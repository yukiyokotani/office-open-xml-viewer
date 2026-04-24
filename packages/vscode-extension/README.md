# OOXML Viewer for VS Code

A high-fidelity viewer for `.xlsx`, `.docx`, and `.pptx` files powered by a Rust/WASM parser and HTML Canvas renderer.

## Features

- **XLSX** — Spreadsheet viewer with cell/row/column selection, range copy (Ctrl+C), freeze-pane support, and multi-sheet tab bar.
- **DOCX** — Word document viewer with transparent text selection overlay (native drag-to-select + copy).
- **PPTX** — Presentation viewer with slide navigation, transparent text selection overlay, and media poster rendering.

All three formats share the same underlying Rust parser (`wasm-pack`) for maximum accuracy and speed.

## Usage

Simply open any `.xlsx`, `.docx`, or `.pptx` file in VS Code — the OOXML Viewer activates automatically as the default editor for those file types.

### Navigation

| Format | Shortcut |
|--------|---------|
| PPTX | Prev / Next slide buttons in toolbar |
| DOCX | Prev / Next page buttons in toolbar |
| XLSX | Click tabs at the bottom to switch sheets |

### Text Selection

PPTX and DOCX files support transparent text overlays. Drag to select text, then use **Ctrl+C** (or **Cmd+C** on macOS) to copy.

For XLSX, click a cell to select it, or drag across a range. **Ctrl+C** copies the selected cells as tab-separated values.

## Extension Settings

No settings are required. The extension activates automatically for the registered file types.

## Known Limitations

- XLSX: formula evaluation is not yet supported (raw values are shown).
- DOCX: automatic image float wrap, footnotes, and header/footer rendering may differ slightly from Word.
- PPTX: some obscure preset shapes fall back to a rectangle placeholder.
- Media playback (audio/video) is not supported in the Webview.

## Development

```bash
# Build extension + webview bundle
cd packages/vscode-extension
node esbuild.config.mjs

# Watch mode
node esbuild.config.mjs --watch

# Package as .vsix
pnpm package
```

To test in the Extension Development Host:
1. Open this repo in VS Code.
2. Press **F5** — a new VS Code window opens with the extension loaded.
3. Open any `.xlsx`, `.docx`, or `.pptx` sample file from `packages/*/public/`.

> **This entire codebase — Rust parsers, TypeScript renderers, tests, and tooling — was implemented by [Claude](https://claude.ai)** (Anthropic's AI assistant) through iterative prompting. No human-written application code exists in this repository.

<p align="center">
  <img src="docs/images/icon.png" alt="office-open-xml-viewer" width="160" height="160">
</p>

# office-open-xml-viewer

[![npm version](https://img.shields.io/npm/v/@silurus/ooxml.svg)](https://www.npmjs.com/package/@silurus/ooxml)
[![npm downloads](https://img.shields.io/npm/dm/@silurus/ooxml.svg)](https://www.npmjs.com/package/@silurus/ooxml)
[![VS Code Marketplace](https://img.shields.io/visual-studio-marketplace/v/silurus.office-open-xml-viewer?label=VS%20Code&logo=visualstudiocode)](https://marketplace.visualstudio.com/items?itemName=silurus.office-open-xml-viewer)
[![VS Code installs](https://img.shields.io/visual-studio-marketplace/i/silurus.office-open-xml-viewer?label=installs)](https://marketplace.visualstudio.com/items?itemName=silurus.office-open-xml-viewer)
[![license](https://img.shields.io/npm/l/@silurus/ooxml.svg)](./LICENSE)

**[Demo (Storybook)](https://ooxml.silurus.dev)**

A browser-based viewer for Office Open XML documents that renders to an HTML Canvas element.
The parsers are written in Rust and compiled to WebAssembly; the renderers use the Canvas 2D API.
Each format also exposes a headless engine (`DocxDocument` / `XlsxWorkbook` / `PptxPresentation`) that renders into any caller-supplied canvas, so you can compose your own UI — scroll views, thumbnail grids, master-detail panes — instead of being locked into the built-in viewer. See the `Examples` section in [the Storybook demo](https://ooxml.silurus.dev).

| DOCX | XLSX | PPTX |
|:---:|:---:|:---:|
| ![docx](docs/images/docx.png) | ![xlsx](docs/images/xlsx.png) | ![pptx](docs/images/pptx.png) |

```bash
npm install @silurus/ooxml
# or
pnpm add @silurus/ooxml
```

> **Bundler note**: this package embeds `.wasm` files. With Vite add [`vite-plugin-wasm`](https://github.com/Menci/vite-plugin-wasm); with webpack use [`experiments.asyncWebAssembly`](https://webpack.js.org/configuration/experiments/).

> **Bundle size note**: npm's *Unpacked Size* figure sums ES (`.mjs`) and CJS (`.cjs`) outputs for all three formats. The size that actually lands in your app is much smaller — import only the format you need (e.g. `@silurus/ooxml/pptx`) and your bundler picks a single module format, so tree-shaking drops the other two formats entirely.

---

## Quick Start

```typescript
import { DocxViewer } from '@silurus/ooxml/docx';
import { XlsxViewer } from '@silurus/ooxml/xlsx';
import { PptxViewer } from '@silurus/ooxml/pptx';

// DOCX — caller provides the <canvas>
const canvas = document.getElementById('docx-canvas') as HTMLCanvasElement;
const docx = new DocxViewer(canvas);
await docx.load('/document.docx');
docx.nextPage();

// XLSX — viewer manages its own <canvas> + tab bar
const xlsx = new XlsxViewer(document.getElementById('xlsx-container')!);
await xlsx.load('/workbook.xlsx');

// PPTX — caller provides the <canvas>
const pptxCanvas = document.getElementById('pptx-canvas') as HTMLCanvasElement;
const pptx = new PptxViewer(pptxCanvas);
await pptx.load('/deck.pptx');
pptx.nextSlide();
```

---

<details>
<summary><strong>Architecture diagram</strong></summary>

```mermaid
flowchart TB
    subgraph build["🦀  Build-time  (Rust → WebAssembly)"]
        direction LR
        docx_rs["packages/docx/parser/src/lib.rs"]
        xlsx_rs["packages/xlsx/parser/src/lib.rs"]
        pptx_rs["packages/pptx/parser/src/lib.rs"]
        docx_rs -- wasm-pack --> docx_wasm["docx_parser.wasm"]
        xlsx_rs -- wasm-pack --> xlsx_wasm["xlsx_parser.wasm"]
        pptx_rs -- wasm-pack --> pptx_wasm["pptx_parser.wasm"]
    end

    subgraph browser["🌐  Runtime  (Browser)"]
        subgraph core_pkg["@silurus/ooxml-core  (shared primitives)"]
            CORE["renderChart · resolveFill · applyStroke\nbuildCustomPath · autoResize · shared types"]
        end
        subgraph docx_pkg["@silurus/ooxml · docx"]
            DV["DocxViewer"] --> DD["DocxDocument"]
            DD --> DW["worker.ts\n〈Web Worker — parse only〉"]
            DD --> DR["renderer.ts\n〈Canvas 2D — main thread〉"]
        end
        subgraph xlsx_pkg["@silurus/ooxml · xlsx"]
            XV["XlsxViewer"] --> XB["XlsxWorkbook"]
            XB --> XW["worker.ts\n〈Web Worker — parse only〉"]
            XB --> XR["renderer.ts\n〈Canvas 2D — main thread〉"]
        end
        subgraph pptx_pkg["@silurus/ooxml · pptx"]
            PV["PptxViewer"] --> PP["PptxPresentation"]
            PP --> PW["worker.ts\n〈Web Worker — parse only〉"]
            PP --> PR["renderer.ts\n〈Canvas 2D — main thread〉"]
        end
        DR -. uses .-> CORE
        XR -. uses .-> CORE
        PR -. uses .-> CORE
    end

    docx_wasm --> DW
    xlsx_wasm --> XW
    pptx_wasm --> PW
    DR --> canvas["&lt;canvas&gt;"]
    XR --> canvas
    PR --> canvas
```

All three formats follow the same shape: the worker parses the `.docx` / `.xlsx` / `.pptx` archive via WASM and posts a JSON model back to the main thread, where the renderer draws to the canvas. Rendering stays on the main thread so the canvas shares the document's `FontFaceSet` — an `OffscreenCanvas` in a worker has its own font registry and would silently fall back to a system font, producing subtly different text measurements (and wrap positions) from the installed theme webfonts. `@silurus/ooxml-core` holds the cross-format primitives that the three renderers all depend on: a unified chart renderer (bar / line / area / radar / waterfall), shape helpers (`resolveFill`, `applyStroke`, `buildCustomPath`, `hexToRgba`), the `autoResize` viewer utility, and the shared type definitions.

### Key files

| File | Role |
|------|------|
| `packages/docx/parser/src/lib.rs` | Rust WASM parser — DOCX ZIP → `Document` JSON |
| `packages/xlsx/parser/src/lib.rs` | Rust WASM parser — XLSX ZIP → `Workbook` JSON |
| `packages/pptx/parser/src/lib.rs` | Rust WASM parser — PPTX ZIP → `Presentation` JSON |
| `packages/docx/src/renderer.ts` | Canvas 2D rendering engine with text layout (main thread) |
| `packages/xlsx/src/renderer.ts` | Canvas 2D rendering engine with virtual scroll (main thread) |
| `packages/pptx/src/renderer.ts` | Canvas 2D rendering engine (main thread) |
| `packages/*/src/worker.ts` | Web Worker: WASM init and parsing only (one per format) |
| `packages/*/src/viewer.ts` | Public Viewer API — canvas lifecycle, navigation |
| `packages/core/src/index.ts` | Cross-format primitives — chart renderer, shape helpers, `autoResize`, shared types |

</details>

---

## Framework Examples

<details>
<summary><strong>React 19</strong></summary>

```tsx
// React 19.1 — vite-plugin-wasm required in vite.config.ts
import { useEffect, useRef, useState } from 'react';
import { PptxViewer } from '@silurus/ooxml/pptx';

export function PptxViewerComponent({ src }: { src: string }) {
  const canvasRef  = useRef<HTMLCanvasElement>(null);
  const viewerRef  = useRef<PptxViewer | null>(null);
  const [slide, setSlide] = useState({ current: 0, total: 0 });

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;

    const viewer = new PptxViewer(canvas, {
      onSlideChange: (i, total) => setSlide({ current: i, total }),
    });
    viewerRef.current = viewer;
    viewer.load(src);
  }, [src]);

  return (
    <div>
      <canvas ref={canvasRef} style={{ width: 800 }} />
      <button onClick={() => viewerRef.current?.prevSlide()}>‹ Prev</button>
      <span> {slide.current + 1} / {slide.total} </span>
      <button onClick={() => viewerRef.current?.nextSlide()}>Next ›</button>
    </div>
  );
}
```

</details>

<details>
<summary><strong>Vue 3.5</strong></summary>

```vue
<!-- Vue 3.5 — useTemplateRef is a 3.5+ feature -->
<script setup lang="ts">
import { useTemplateRef, onMounted, ref } from 'vue';
import { PptxViewer } from '@silurus/ooxml/pptx';

const props = defineProps<{ src: string }>();

const canvas  = useTemplateRef<HTMLCanvasElement>('canvas');
let viewer: PptxViewer | null = null;
const current = ref(0);
const total   = ref(0);

onMounted(async () => {
  viewer = new PptxViewer(canvas.value!, {
    onSlideChange: (i, t) => { current.value = i; total.value = t; },
  });
  await viewer.load(props.src);
});
</script>

<template>
  <div>
    <canvas ref="canvas" style="width: 800px" />
    <button @click="viewer?.prevSlide()">‹ Prev</button>
    <span> {{ current + 1 }} / {{ total }} </span>
    <button @click="viewer?.nextSlide()">Next ›</button>
  </div>
</template>
```

</details>

<details>
<summary><strong>Angular 19</strong></summary>

```typescript
// Angular 19 — standalone component with signal-based state
import {
  Component, ElementRef, viewChild,
  signal, AfterViewInit,
} from '@angular/core';
import { PptxViewer } from '@silurus/ooxml/pptx';

@Component({
  selector: 'app-pptx-viewer',
  standalone: true,
  template: `
    <div>
      <canvas #canvas style="width: 800px"></canvas>
      <button (click)="prev()">‹ Prev</button>
      <span> {{ current() + 1 }} / {{ total() }} </span>
      <button (click)="next()">Next ›</button>
    </div>
  `,
})
export class PptxViewerComponent implements AfterViewInit {
  canvasEl = viewChild.required<ElementRef<HTMLCanvasElement>>('canvas');
  current = signal(0);
  total   = signal(0);
  private viewer?: PptxViewer;

  ngAfterViewInit(): void {
    this.viewer = new PptxViewer(this.canvasEl().nativeElement, {
      onSlideChange: (i, t) => { this.current.set(i); this.total.set(t); },
    });
    this.viewer.load('/deck.pptx');
  }

  prev(): void { this.viewer?.prevSlide(); }
  next(): void { this.viewer?.nextSlide(); }
}
```

> Add `"allowSyntheticDefaultImports": true` and configure `@angular-builders/custom-webpack` (or use `esbuild` builder) with WASM support in your Angular workspace.

</details>

<details>
<summary><strong>Svelte 5</strong></summary>

```svelte
<!-- Svelte 5 — runes syntax ($props, $state) -->
<script lang="ts">
  import { onMount } from 'svelte';
  import { PptxViewer } from '@silurus/ooxml/pptx';

  let { src }: { src: string } = $props();

  let canvas: HTMLCanvasElement;
  let viewer: PptxViewer;
  let current = $state(0);
  let total   = $state(0);

  onMount(async () => {
    viewer = new PptxViewer(canvas, {
      onSlideChange: (i, t) => { current = i; total = t; },
    });
    await viewer.load(src);
  });
</script>

<div>
  <canvas bind:this={canvas} style="width: 800px"></canvas>
  <button onclick={() => viewer?.prevSlide()}>‹ Prev</button>
  <span> {current + 1} / {total} </span>
  <button onclick={() => viewer?.nextSlide()}>Next ›</button>
</div>
```

</details>

<details>
<summary><strong>SolidJS 1.9</strong></summary>

```tsx
// SolidJS 1.9
import { createSignal, onMount, onCleanup } from 'solid-js';
import { PptxViewer } from '@silurus/ooxml/pptx';

export function PptxViewerComponent(props: { src: string }) {
  let canvasEl!: HTMLCanvasElement;
  let viewer: PptxViewer | undefined;
  const [current, setCurrent] = createSignal(0);
  const [total,   setTotal  ] = createSignal(0);

  onMount(async () => {
    viewer = new PptxViewer(canvasEl, {
      onSlideChange: (i, t) => { setCurrent(i); setTotal(t); },
    });
    await viewer.load(props.src);
  });

  onCleanup(() => { /* viewer?.destroy?.() */ });

  return (
    <div>
      <canvas ref={canvasEl} style={{ width: '800px' }} />
      <button onClick={() => viewer?.prevSlide()}>‹ Prev</button>
      <span> {current() + 1} / {total()} </span>
      <button onClick={() => viewer?.nextSlide()}>Next ›</button>
    </div>
  );
}
```

</details>

<details>
<summary><strong>Qwik 2</strong></summary>

```tsx
// Qwik 2.0 — dynamic import to keep WASM out of SSR bundle
import { component$, useSignal, useVisibleTask$ } from '@builder.io/qwik';
import type { PptxViewer as PptxViewerType } from '@silurus/ooxml/pptx';

export const PptxViewerComponent = component$<{ src: string }>(({ src }) => {
  const canvasRef = useSignal<HTMLCanvasElement>();
  const current = useSignal(0);
  const total   = useSignal(0);
  let viewer: PptxViewerType | undefined;

  // useVisibleTask$ runs only in the browser, never during SSR
  useVisibleTask$(async () => {
    if (!canvasRef.value) return;
    const { PptxViewer } = await import('@silurus/ooxml/pptx');
    viewer = new PptxViewer(canvasRef.value, {
      onSlideChange: (i, t) => { current.value = i; total.value = t; },
    });
    await viewer.load(src);
  });

  return (
    <div>
      <canvas ref={canvasRef} style={{ width: '800px' }} />
      <button onClick$={() => viewer?.prevSlide()}>‹ Prev</button>
      <span> {current.value + 1} / {total.value} </span>
      <button onClick$={() => viewer?.nextSlide()}>Next ›</button>
    </div>
  );
});
```

</details>

---

## Feature Support

### Word (.docx)

| Category | Feature | Status |
|----------|---------|--------|
| **Document** | Page rendering | ✅ |
| | Page size and margins | ✅ |
| | Headers / footers (default / first / even) | ✅ |
| | Section breaks | ❌ |
| **Text** | Paragraphs | ✅ |
| | Bold, italic, underline, strikethrough | ✅ |
| | Font family, size, color | ✅ |
| | Hyperlinks | ✅ |
| | Superscript / subscript (`w:vertAlign`) | ✅ |
| **Formatting** | Paragraph alignment (left/center/right/justify) | ✅ |
| | Line spacing (auto / atLeast / exact) | ✅ |
| | Line grid (`w:docGrid`, §17.6.5) | ✅ |
| | Margin collapsing between paragraphs | ✅ |
| | Indents and tab stops | ✅ |
| | Lists (bullet and numbered) | ✅ |
| | Paragraph styles (Heading 1–9, Normal, custom) | ✅ |
| | Table style `w:pPr` cascade (§17.7.6) | ✅ |
| | keepNext / keepLines / widowControl | ✅ |
| **Elements** | Tables (with borders, fills, merges) | ✅ |
| | Images (inline and anchored, with text wrap) | ✅ |
| | Text boxes / drawing shapes | ✅ |
| **Advanced** | Footnote / endnote reference markers | ✅ |
| | Track changes / comments | ❌ |
| | Mail merge fields | ❌ Not planned |
| **Interaction** | Text selection (transparent overlay, native copy) | ✅ |

---

### Excel (.xlsx)

| Category | Feature | Status |
|----------|---------|--------|
| **Workbook** | Multiple sheets, sheet names | ✅ |
| **Cells** | Text, number, boolean, error values | ✅ |
| | Formula results (from cached `<v>`) | ✅ |
| | Dates (ECMA-376 date format codes) | ✅ |
| | Rich text (per-run formatting) | ✅ |
| **Formatting** | Bold, italic, underline, strikethrough | ✅ |
| | Font family, size, color | ✅ |
| | Cell background color | ✅ |
| | Borders | ✅ |
| | Horizontal / vertical alignment | ✅ |
| | Text wrapping | ✅ |
| | Number formats (`0.00`, `%`, `#,##0`, custom date/time) | ✅ |
| **Structure** | Merged cells | ✅ |
| | Frozen panes | ✅ |
| | Row / column sizing (custom widths and heights) | ✅ |
| | Hidden rows / columns | ✅ |
| **Elements** | Images (`<xdr:twoCellAnchor>`) | ✅ |
| | Drawing shapes / text boxes (`xdr:sp`, `xdr:txBody`) | ✅ |
| | Charts (bar, line, area, radar, scatter / bubble) | ✅ |
| | Chart markers (circle / square / diamond / triangle / x / plus / star / dot / dash, per-point `<c:dPt>` overrides) | ✅ |
| | Chart data labels (`<c:dLbl>` per-point with CELLRANGE / VALUE / SERIESNAME / CATEGORYNAME field references, position `l`/`r`/`t`/`b`/`ctr`/`outEnd`) | ✅ |
| | Chart error bars (`<c:errBars>` X/Y direction, `cust` / `fixedVal` / `stdErr` / `stdDev` / `percentage`, dashed/styled lines) | ✅ |
| | Chart manual layout (`<c:title><c:layout>` and `<c:plotArea><c:layout>`) | ✅ |
| | Sparklines (`x14:sparklineGroup` — line / column / win-loss, with markers and high/low/first/last/negative highlights) | ✅ |
| **Advanced** | Conditional formatting (`cellIs`, `colorScale`, `dataBar`, `iconSet`, `top10`, `aboveAverage`) | ✅ |
| | Slicers (static, Office 2010 extension) | ✅ |
| | Pivot tables | ❌ Not planned |
| | Data validation / comments | ❌ Not planned |
| **Interaction** | Cell selection (single / range / row / column / all) | ✅ |
| | Shift+click to extend, Ctrl+C to copy as TSV | ✅ |
| | Text selection inside cells (transparent overlay) | ✅ |
| | `onSelectionChange` callback, `getCellAt(x, y)` API | ✅ |

---

### PowerPoint (.pptx)

| Category | Feature | Status |
|----------|---------|--------|
| **Slides** | Slide rendering | ✅ |
| | Slide layout / master inheritance | ✅ |
| | Slide size (custom dimensions) | ✅ |
| | Slide background (solid, gradient, image) | ✅ |
| | Slide numbers | ✅ |
| | Notes pages | ❌ |
| | Animations / transitions | ❌ |
| **Element types** | Shapes (`sp`) | ✅ |
| | Pictures (`pic`) | ✅ |
| | Groups (`grpSp`) with nested transforms | ✅ |
| | Connectors (`cxnSp`) | ✅ |
| | Tables (`tbl` in `graphicFrame`) | ✅ |
| | Charts (bar, line, area, radar, waterfall) | ✅ |
| | Charts (pie, scatter, bubble) | ❌ |
| | SmartArt | ❌ |
| | OLE objects | ❌ |
| | Video / audio (poster + interactive playback) | ✅ |
| | Ink / handwriting (`p:contentPart`, raster fallback) | ✅ |
| **Shape geometry** | 130+ preset shapes (`prstGeom`) | ✅ |
| | Custom geometry (`custGeom`) | ✅ |
| | Rotation and flip (flipH / flipV) | ✅ |
| | 3D preset shapes | ❌ |
| **Fills** | Solid fill (`solidFill`) | ✅ |
| | Linear / radial gradient (`gradFill`) | ✅ |
| | No fill (`noFill`) | ✅ |
| | Pattern fill (`pattFill`) | ❌ |
| | Image fill on shapes (`blipFill` in `sp`) | ✅ |
| **Strokes** | Solid line color and width | ✅ |
| | Dash / dot styles | ✅ |
| | Arrow heads (`headEnd` / `tailEnd`) | ✅ |
| | Compound / double lines | ❌ |
| **Shape effects** | Drop shadow (`outerShdw`) | ✅ |
| | Inner shadow / glow / reflection | ❌ |
| | Bevel / 3D extrusion | ❌ |
| **Text — characters** | Bold, italic, underline, strikethrough | ✅ |
| | Font family, size, color | ✅ |
| | Superscript / subscript | ✅ |
| | Hyperlinks | ❌ |
| | Text shadow / outline effects | ❌ |
| **Text — paragraphs** | Horizontal alignment (left / center / right / justify) | ✅ |
| | Vertical anchor (top / center / bottom) | ✅ |
| | Line spacing (`spcPct`, `spcPts`) | ✅ |
| | Space before / after paragraph | ✅ |
| | Bullet points (character and auto-numbered) | ✅ |
| | Tab stops | ✅ |
| | Indent / margin | ✅ |
| | Vertical / RTL text | ❌ |
| **Text — body** | Text padding (insets) | ✅ |
| | normAutoFit (shrink to fit) | ✅ |
| | spAutoFit (expand box) | ✅ |
| | Word wrap / no wrap | ✅ |
| **Tables** | Cells, rows, columns | ✅ |
| | Cell merges (horizontal / vertical) | ✅ |
| | Cell borders | ✅ |
| | Cell fills (solid / gradient) | ✅ |
| | Cell diagonal lines (`lnTlToBr` / `lnBlToTr`) | ✅ |
| | Table theme styles (74 built-in PowerPoint presets) | ✅ |
| **Theme** | Scheme colors (dk1/lt1/accent1–6) | ✅ |
| | Font scheme (`+mj-lt`, `+mn-lt`) | ✅ |
| | lumMod / lumOff / alpha transforms | ✅ |
| **Interaction** | Text selection (transparent overlay, native copy) | ✅ |

---

> **A note on text selection.** Across DOCX / PPTX / XLSX, text selection is currently implemented by rendering glyphs to the canvas while overlaying a transparent DOM layer that mirrors the canvas text positions for native browser selection. This dual-layer approach is a deliberate stop-gap: once the Canvas [`drawElement` API](https://chromestatus.com/feature/6051647656558592) (proposed in [WICG/html-in-canvas](https://github.com/WICG/html-in-canvas), currently in Chromium Origin Trial) ships across browsers, the project plans to migrate to a single DOM-as-source-of-truth pipeline where the canvas mirrors the DOM directly — eliminating the duplication while keeping z-order correctness and native selection / a11y.

---

## Companion packages

- **[`packages/vscode-extension/`](packages/vscode-extension/)** — VS Code extension (`ooxml-viewer`) that registers `CustomEditorProvider`s for `.docx`, `.xlsx`, and `.pptx`. Open Office files directly in the editor with the same Canvas renderer plus selection/copy.
- **[`packages/mcp-server/`](packages/mcp-server/)** — Rust MCP server (`ooxml-mcp-server`) exposing the parsers as tools for AI agents (Claude, Copilot, Codex, etc.). Provides structured queries (`docx_get_structure`, `xlsx_get_cell_range`, `pptx_get_slide_structure`, …) so agents can inspect OOXML files without shelling out to `unzip`.

---

## Development

```bash
# Install dependencies
pnpm install

# Build all WASM parsers (requires Rust + wasm-pack)
pnpm build:wasm

# Start Storybook dev server (port 6006)
pnpm storybook

# Type-check all packages
pnpm typecheck

# Run visual regression tests
pnpm vrt

# Build the library
pnpm build
```

### WASM build (individual packages)

```bash
cd packages/docx/parser && wasm-pack build --target web && cp pkg/docx_parser_bg.wasm  pkg/docx_parser.js  ../src/wasm/
cd packages/xlsx/parser && wasm-pack build --target web && cp pkg/xlsx_parser_bg.wasm  pkg/xlsx_parser.js  ../src/wasm/
cd packages/pptx/parser && wasm-pack build --target web && cp pkg/pptx_parser_bg.wasm pkg/pptx_parser.js ../src/wasm/
```

## Security & Privacy

- **Canvas-only rendering.** Documents are decoded and drawn to an `HTMLCanvasElement`. No script, link, form, or other active content from the source file is executed or injected into the DOM.
- **ZIP decompression cap.** Each entry in the source archive is limited to 512 MiB of uncompressed output to block zip-bomb DoS.
- **No network by default.** The library does not send telemetry or analytics, and does not contact third-party services unless you ask it to. In particular, theme webfonts (and Office font metric substitutes for XLSX) are **not** loaded from Google Fonts unless you pass `useGoogleFonts: true` to the relevant `Viewer` / `load(...)` options — supported uniformly by `DocxViewer`, `PptxViewer`, and `XlsxViewer`. Enabling that option causes the end-user's browser to send an HTTP request (IP and User-Agent) to `fonts.googleapis.com`, which may have GDPR implications for your application — consider self-hosting the required fonts via `@font-face` instead.
- **XML parsing.** Uses `roxmltree`, which does not resolve external entities (XXE-safe by default).

## License

MIT

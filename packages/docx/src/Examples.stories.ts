import type { Meta, StoryObj } from '@storybook/html';
import { buildViewerUI } from './DocxViewer.stories';
import { DocxDocument } from './document';
import { DocxViewer } from './viewer';
import { DocxScrollViewer } from './scroll-viewer';
import { buildDocxTextLayer } from './text-layer';
import type { DocxTextRunInfo } from './renderer';
import { math } from '../../../src/math';
import { threeD } from '../../../src/three-d';
import { regionMap } from '../../../src/region-map';
import { chartEx } from '../../../src/chart-ex';

const fullRenderers = { math, threeD, regionMap, chartEx } as const;

type DemoArgs = { width: number };
type LayoutArgs = Record<string, never>;

const SAMPLE_URL = `${import.meta.env.BASE_URL}docx/demo/sample-1.docx`;

const meta: Meta<DemoArgs> = {
  title: 'DocxViewer/Examples',
  argTypes: {
    width: {
      control: { type: 'range', min: 400, max: 1200, step: 40 },
      description: 'Canvas render width (px) — used by the Demo story',
    },
  },
  // Match the pptx Examples width so the loading placeholder (and the rendered
  // document) are the same size across the Demo / Offscreen stories.
  args: { width: 960 },
};
export default meta;

type DemoStory = StoryObj<DemoArgs>;
type LayoutStory = StoryObj<LayoutArgs>;

export const Demo: DemoStory = {
  name: 'Demo — single viewer (demo.docx)',
  render(args) {
    const { root } = buildViewerUI(args, SAMPLE_URL);
    return root;
  },
};

export const Offscreen: DemoStory = {
  name: 'Offscreen — Web Worker rendering (demo.docx)',
  // The single-viewer Demo, rendered entirely in a Web Worker (mode: 'worker').
  // Identical UX — only the pixels are produced off the main thread.
  render(args) {
    const { root } = buildViewerUI(args, SAMPLE_URL, { mode: 'worker' });
    return root;
  },
};

function makeStatus(root: HTMLElement): HTMLDivElement {
  const s = document.createElement('div');
  s.style.cssText = 'color:#666;font-size:13px;margin-bottom:8px;min-height:18px;';
  s.textContent = 'Loading…';
  root.appendChild(s);
  return s;
}

export const ScrollView: LayoutStory = {
  name: 'ScrollView — stack all pages',
  render() {
    const root = document.createElement('div');
    root.style.cssText = 'font-family:sans-serif;padding:16px;';
    const heading = document.createElement('h3');
    heading.textContent = 'ScrollView — scroll through every page';
    heading.style.cssText = 'margin:0 0 8px;font-size:14px;';
    root.appendChild(heading);
    const status = makeStatus(root);

    const scroller = document.createElement('div');
    scroller.style.cssText =
      'max-height:720px;overflow-y:auto;border:1px solid #ccc;background:#f5f5f5;padding:12px;';
    root.appendChild(scroller);

    DocxDocument.load(SAMPLE_URL, fullRenderers)
      .then(async (doc) => {
        status.textContent = `Rendering ${doc.pageCount} pages…`;
        const widthPx = 700;

        for (let i = 0; i < doc.pageCount; i++) {
          const pageWrapper = document.createElement('div');
          pageWrapper.style.cssText =
            'position:relative;display:block;max-width:700px;margin:0 auto 12px;';

          const canvas = document.createElement('canvas');
          canvas.style.cssText =
            'display:block;width:100%;max-width:700px;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,0.2);';

          const textLayer = document.createElement('div');
          textLayer.style.cssText =
            'position:absolute;top:0;left:0;width:100%;height:100%;' +
            'overflow:hidden;pointer-events:none;user-select:text;-webkit-user-select:text;';

          pageWrapper.appendChild(canvas);
          pageWrapper.appendChild(textLayer);
          scroller.appendChild(pageWrapper);

          const runs: DocxTextRunInfo[] = [];
          await doc.renderPage(canvas, i, { width: widthPx, onTextRun: (r) => runs.push(r) });
          // Pass the page's intended CSS box (px) as the % denominators. The
          // canvas is scaled responsively (width:100%;max-width:700px), and the
          // overlay's %-placed spans track that actual rendered size.
          const cssHeight = parseFloat(canvas.style.height) || canvas.height;
          buildDocxTextLayer(textLayer, runs, widthPx, cssHeight);
        }
        status.textContent = `Loaded ${doc.pageCount} pages`;
      })
      .catch((e: Error) => {
        status.textContent = `Error: ${e.message}`;
        status.style.color = 'red';
      });

    return root;
  },
};

// The scroll viewer owns a live scroll subscription, a ResizeObserver, and a
// per-slot canvas pool, so a stale instance left mounted across a Storybook
// re-render (args change / HMR) would keep observing a detached container and
// leak canvases. Hold the last instance module-side and destroy it before
// building a fresh one.
let scrollViewerInstance: DocxScrollViewer | null = null;

export const ScrollViewer: LayoutStory = {
  name: 'DocxScrollViewer — virtualized continuous scroll',
  render() {
    scrollViewerInstance?.destroy();
    scrollViewerInstance = null;

    const root = document.createElement('div');
    root.style.cssText = 'font-family:sans-serif;padding:16px;';
    const heading = document.createElement('h3');
    heading.textContent = 'DocxScrollViewer — virtualized (Ctrl/⌘+wheel to zoom)';
    heading.style.cssText = 'margin:0 0 8px;font-size:14px;';
    root.appendChild(heading);
    const status = makeStatus(root);

    // The viewer owns a container-sized scroll surface; give it a fixed box.
    const container = document.createElement('div');
    container.style.cssText = 'height:720px;border:1px solid #ccc;background:#f5f5f5;';
    root.appendChild(container);

    const viewer = new DocxScrollViewer(container, {
      gap: 16,
      paddingTop: 24, // desk margin above the first page (defaults to gap when omitted)
      paddingBottom: 24, // desk margin below the last page
      paddingLeft: 24, // horizontal desk gutter left of the pages
      paddingRight: 24, // horizontal desk gutter right of the pages
      overscan: 1,
      enableTextSelection: true,
      background: '#f5f5f5', // light desk behind/between pages (matches the ScrollView recipe)
      onVisiblePageChange: (top, total, complete) => {
        // The ellipsis marks a provisional total: under progressive layout the
        // count grows until the whole document has been laid out.
        status.textContent = `Page ${top + 1} / ${total}${complete ? '' : '…'}`;
      },
      onError: (e) => {
        status.textContent = `Error: ${e.message}`;
        status.style.color = 'red';
      },
    });
    scrollViewerInstance = viewer;
    viewer
      .load(SAMPLE_URL)
      .then(() => {
        status.textContent = `Loaded ${viewer.pageCount} pages`;
      })
      .catch((e: Error) => {
        status.textContent = `Error: ${e.message}`;
        status.style.color = 'red';
      });

    return root;
  },
};

let progressiveViewerInstance: DocxScrollViewer | null = null;

export const ProgressiveScrollViewer: LayoutStory = {
  name: 'DocxScrollViewer — progressive first paint',
  render() {
    progressiveViewerInstance?.destroy();
    progressiveViewerInstance = null;

    const root = document.createElement('div');
    root.style.cssText = 'font-family:sans-serif;padding:16px;';
    const heading = document.createElement('h3');
    heading.textContent = 'DocxScrollViewer — progressive first paint';
    heading.style.cssText = 'margin:0 0 8px;font-size:14px;';
    root.appendChild(heading);
    const note = document.createElement('p');
    note.textContent =
      'With progressiveLayout, load() resolves on the document\u2019s opening pages and the '
      + 'rest is laid out in the background, so the page count starts small and grows. The '
      + 'difference is only visible on a long document \u2014 a short one finishes before '
      + 'there is anything to defer.';
    note.style.cssText = 'margin:0 0 8px;font-size:12px;color:#555;max-width:60em;';
    root.appendChild(note);
    const status = makeStatus(root);

    const container = document.createElement('div');
    container.style.cssText = 'height:720px;border:1px solid #ccc;background:#f5f5f5;';
    root.appendChild(container);

    const started = performance.now();
    const viewer = new DocxScrollViewer(container, {
      gap: 16,
      overscan: 1,
      enableTextSelection: true,
      background: '#f5f5f5',
      progressiveLayout: true,
      onError: (e) => {
        status.textContent = `Error: ${e.message}`;
        status.style.color = 'red';
      },
    });
    progressiveViewerInstance = viewer;
    viewer
      .load(SAMPLE_URL)
      .then(async () => {
        const painted = Math.round(performance.now() - started);
        status.textContent = `First paint after ${painted}ms with ${viewer.pageCount} page(s); laying out the rest\u2026`;
        // pageCount is the pages available so far, not the document total, until
        // layout completes.
        await viewer.waitUntilLayoutComplete();
        const settled = Math.round(performance.now() - started);
        status.textContent = `First paint ${painted}ms \u2192 full layout ${settled}ms, ${viewer.pageCount} pages`;
      })
      .catch((e: Error) => {
        status.textContent = `Error: ${e.message}`;
        status.style.color = 'red';
      });

    return root;
  },
};

export const ThumbnailGrid: LayoutStory = {
  name: 'ThumbnailGrid — overview of all pages',
  render() {
    const root = document.createElement('div');
    root.style.cssText = 'font-family:sans-serif;padding:16px;';
    const heading = document.createElement('h3');
    heading.textContent = 'ThumbnailGrid — every page at a glance';
    heading.style.cssText = 'margin:0 0 8px;font-size:14px;';
    root.appendChild(heading);
    const status = makeStatus(root);

    const grid = document.createElement('div');
    grid.style.cssText =
      'display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:16px;';
    root.appendChild(grid);

    DocxDocument.load(SAMPLE_URL, fullRenderers)
      .then(async (doc) => {
        status.textContent = `Rendering ${doc.pageCount} thumbnails…`;
        const thumbWidth = 160;
        for (let i = 0; i < doc.pageCount; i++) {
          const cell = document.createElement('div');
          cell.style.cssText = 'display:flex;flex-direction:column;align-items:center;cursor:pointer;';
          const canvas = document.createElement('canvas');
          canvas.style.cssText =
            'display:block;width:100%;max-width:160px;background:#fff;' +
            'box-shadow:0 1px 3px rgba(0,0,0,0.2);';
          const caption = document.createElement('div');
          caption.textContent = `Page ${i + 1}`;
          caption.style.cssText = 'font-size:12px;color:#444;margin-top:4px;';
          cell.append(canvas, caption);
          const idx = i;
          cell.addEventListener('click', () => {
            console.log(`[docx ThumbnailGrid] clicked page ${idx + 1}`);
          });
          grid.appendChild(cell);
          await doc.renderPage(canvas, i, { width: thumbWidth });
        }
        status.textContent = `Loaded ${doc.pageCount} pages`;
      })
      .catch((e: Error) => {
        status.textContent = `Error: ${e.message}`;
        status.style.color = 'red';
      });

    return root;
  },
};

export const MasterDetail: LayoutStory = {
  name: 'MasterDetail — thumbnails + large preview',
  render() {
    const root = document.createElement('div');
    root.style.cssText = 'font-family:sans-serif;padding:16px;';
    const heading = document.createElement('h3');
    heading.textContent = 'MasterDetail — click a thumbnail to preview';
    heading.style.cssText = 'margin:0 0 8px;font-size:14px;';
    root.appendChild(heading);
    const status = makeStatus(root);

    const layout = document.createElement('div');
    layout.style.cssText = 'display:flex;gap:16px;height:720px;';
    root.appendChild(layout);

    const thumbCol = document.createElement('div');
    thumbCol.style.cssText =
      'flex:0 0 200px;overflow-y:auto;border:1px solid #ccc;background:#f5f5f5;padding:8px;' +
      'display:flex;flex-direction:column;gap:10px;';
    const detailCol = document.createElement('div');
    detailCol.style.cssText =
      'flex:1 1 auto;border:1px solid #ccc;background:#f5f5f5;padding:12px;overflow:auto;' +
      'display:flex;align-items:flex-start;justify-content:center;';
    layout.append(thumbCol, detailCol);

    // Detail canvas + viewer with text selection
    const detailCanvas = document.createElement('canvas');
    detailCanvas.style.cssText = 'display:block;max-width:100%;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,0.2);';
    detailCol.appendChild(detailCanvas);
    const detailViewer = new DocxViewer(detailCanvas, {
      enableTextSelection: true,
      width: 600,
      ...fullRenderers,
    });

    // Load thumbnails using DocxDocument; detail viewer loads independently
    Promise.all([
      DocxDocument.load(SAMPLE_URL, fullRenderers),
      detailViewer.load(SAMPLE_URL),
    ])
      .then(async ([doc]) => {
        status.textContent = `Rendering ${doc.pageCount} thumbnails…`;
        const thumbEntries: HTMLDivElement[] = [];

        const selectPage = async (i: number) => {
          for (let k = 0; k < thumbEntries.length; k++) {
            thumbEntries[k].style.outline = k === i ? '2px solid #0366d6' : 'none';
          }
          await detailViewer.goToPage(i);
        };

        for (let i = 0; i < doc.pageCount; i++) {
          const cell = document.createElement('div');
          cell.style.cssText = 'display:flex;flex-direction:column;align-items:center;cursor:pointer;padding:4px;';
          const canvas = document.createElement('canvas');
          canvas.style.cssText =
            'display:block;width:100%;max-width:180px;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,0.2);';
          const caption = document.createElement('div');
          caption.textContent = `Page ${i + 1}`;
          caption.style.cssText = 'font-size:12px;color:#444;margin-top:4px;';
          cell.append(canvas, caption);
          const idx = i;
          cell.addEventListener('click', () => {
            selectPage(idx).catch((e: Error) => {
              status.textContent = `Render error: ${e.message}`;
            });
          });
          thumbCol.appendChild(cell);
          thumbEntries.push(cell);
          await doc.renderPage(canvas, i, { width: 180 });
        }

        // Highlight first thumbnail
        if (thumbEntries.length > 0) thumbEntries[0].style.outline = '2px solid #0366d6';
        status.textContent = `Loaded ${doc.pageCount} pages`;
      })
      .catch((e: Error) => {
        status.textContent = `Error: ${e.message}`;
        status.style.color = 'red';
      });

    return root;
  },
};

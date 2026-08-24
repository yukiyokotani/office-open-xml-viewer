import type { Meta, StoryObj } from '@storybook/html';
import { buildViewerUI } from './PptxViewer.stories';
import { PptxPresentation } from './presentation';
import { PptxViewer } from './viewer';
import { PptxScrollViewer } from './scroll-viewer';
import type { PptxTextRunInfo } from './renderer';
import { buildPptxTextLayer } from './text-layer';
import { math } from '../../../src/math';
import { threeD } from '../../../src/three-d';
import { regionMap } from '../../../src/region-map';
import { chartEx } from '../../../src/chart-ex';

const fullRenderers = { math, threeD, regionMap, chartEx } as const;

type DemoArgs = { width: number };
type LayoutArgs = Record<string, never>;

const SAMPLE_URL = `${import.meta.env.BASE_URL}pptx/demo/sample-1.pptx`;

const meta: Meta<DemoArgs> = {
  title: 'PptxViewer/Examples',
  argTypes: {
    width: {
      control: { type: 'range', min: 400, max: 1600, step: 40 },
      description: 'Canvas render width (px) — used by the Demo story',
    },
  },
  args: { width: 960 },
};
export default meta;

type DemoStory = StoryObj<DemoArgs>;
type LayoutStory = StoryObj<LayoutArgs>;

export const Demo: DemoStory = {
  name: 'Demo — single viewer (demo.pptx)',
  render(args) {
    const { root } = buildViewerUI(args, SAMPLE_URL);
    return root;
  },
};

export const Offscreen: DemoStory = {
  name: 'Offscreen — Web Worker rendering (demo.pptx)',
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
  name: 'ScrollView — stack all slides',
  render() {
    const root = document.createElement('div');
    root.style.cssText = 'font-family:sans-serif;padding:16px;';
    const heading = document.createElement('h3');
    heading.textContent = 'ScrollView — scroll through every slide';
    heading.style.cssText = 'margin:0 0 8px;font-size:14px;';
    root.appendChild(heading);
    const status = makeStatus(root);

    const scroller = document.createElement('div');
    scroller.style.cssText =
      'max-height:720px;overflow-y:auto;border:1px solid #ccc;background:#f5f5f5;padding:12px;';
    root.appendChild(scroller);

    PptxPresentation.load(SAMPLE_URL, { useGoogleFonts: true, ...fullRenderers })
      .then(async (pres) => {
        status.textContent = `Rendering ${pres.slideCount} slides…`;
        const widthPx = 800;
        const cssHeight = Math.round(pres.slideHeight * widthPx / pres.slideWidth);

        for (let i = 0; i < pres.slideCount; i++) {
          const slideWrapper = document.createElement('div');
          slideWrapper.style.cssText =
            'position:relative;display:block;max-width:800px;margin:0 auto 12px;';

          const canvas = document.createElement('canvas');
          canvas.style.cssText =
            'display:block;width:100%;max-width:800px;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,0.2);';

          const textLayer = document.createElement('div');
          textLayer.style.cssText =
            'position:absolute;top:0;left:0;width:100%;height:100%;' +
            'overflow:hidden;pointer-events:none;user-select:text;-webkit-user-select:text;';

          slideWrapper.appendChild(canvas);
          slideWrapper.appendChild(textLayer);
          scroller.appendChild(slideWrapper);

          const runs: PptxTextRunInfo[] = [];
          await pres.renderSlide(canvas, i, { width: widthPx, onTextRun: (r) => runs.push(r) });
          buildPptxTextLayer(textLayer, runs, widthPx, cssHeight);
        }
        status.textContent = `Loaded ${pres.slideCount} slides`;
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
let scrollViewerInstance: PptxScrollViewer | null = null;

export const ScrollViewer: LayoutStory = {
  name: 'PptxScrollViewer — virtualized continuous scroll',
  render() {
    scrollViewerInstance?.destroy();
    scrollViewerInstance = null;

    const root = document.createElement('div');
    root.style.cssText = 'font-family:sans-serif;padding:16px;';
    const heading = document.createElement('h3');
    heading.textContent = 'PptxScrollViewer — virtualized (Ctrl/⌘+wheel to zoom)';
    heading.style.cssText = 'margin:0 0 8px;font-size:14px;';
    root.appendChild(heading);
    const status = makeStatus(root);

    // The viewer owns a container-sized scroll surface; give it a fixed box.
    const container = document.createElement('div');
    container.style.cssText = 'height:720px;border:1px solid #ccc;background:#f5f5f5;';
    root.appendChild(container);

    const viewer = new PptxScrollViewer(container, {
      gap: 16,
      paddingTop: 24, // desk margin above the first slide (defaults to gap when omitted)
      paddingBottom: 24, // desk margin below the last slide
      paddingLeft: 24, // horizontal desk gutter left of the slides
      paddingRight: 24, // horizontal desk gutter right of the slides
      overscan: 1,
      enableTextSelection: true,
      useGoogleFonts: true,
      background: '#f5f5f5', // light desk behind/between slides (matches the ScrollView recipe)
      onVisibleSlideChange: (top, total) => {
        status.textContent = `Slide ${top + 1} / ${total}`;
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
        status.textContent = `Loaded ${viewer.slideCount} slides`;
      })
      .catch((e: Error) => {
        status.textContent = `Error: ${e.message}`;
        status.style.color = 'red';
      });

    return root;
  },
};

export const ThumbnailGrid: LayoutStory = {
  name: 'ThumbnailGrid — overview of all slides',
  render() {
    const root = document.createElement('div');
    root.style.cssText = 'font-family:sans-serif;padding:16px;';
    const heading = document.createElement('h3');
    heading.textContent = 'ThumbnailGrid — every slide at a glance';
    heading.style.cssText = 'margin:0 0 8px;font-size:14px;';
    root.appendChild(heading);
    const status = makeStatus(root);

    const grid = document.createElement('div');
    grid.style.cssText =
      'display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:16px;';
    root.appendChild(grid);

    PptxPresentation.load(SAMPLE_URL, { useGoogleFonts: true, ...fullRenderers })
      .then(async (pres) => {
        status.textContent = `Rendering ${pres.slideCount} thumbnails…`;
        const thumbWidth = 240;
        for (let i = 0; i < pres.slideCount; i++) {
          const cell = document.createElement('div');
          cell.style.cssText = 'display:flex;flex-direction:column;align-items:center;cursor:pointer;';
          const canvas = document.createElement('canvas');
          canvas.style.cssText =
            'display:block;width:100%;max-width:240px;background:#fff;' +
            'box-shadow:0 1px 3px rgba(0,0,0,0.2);';
          const caption = document.createElement('div');
          caption.textContent = `Slide ${i + 1}`;
          caption.style.cssText = 'font-size:12px;color:#444;margin-top:4px;';
          cell.append(canvas, caption);
          const idx = i;
          cell.addEventListener('click', () => {
            console.log(`[pptx ThumbnailGrid] clicked slide ${idx + 1}`);
          });
          grid.appendChild(cell);
          await pres.renderSlide(canvas, i, { width: thumbWidth });
        }
        status.textContent = `Loaded ${pres.slideCount} slides`;
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
      'flex:0 0 240px;overflow-y:auto;border:1px solid #ccc;background:#f5f5f5;padding:8px;' +
      'display:flex;flex-direction:column;gap:10px;';
    const detailCol = document.createElement('div');
    detailCol.style.cssText =
      'flex:1 1 auto;border:1px solid #ccc;background:#f5f5f5;padding:12px;overflow:auto;' +
      'display:flex;align-items:flex-start;justify-content:center;position:relative;';
    layout.append(thumbCol, detailCol);

    // Detail viewer with text selection
    const detailCanvas = document.createElement('canvas');
    detailCol.appendChild(detailCanvas);
    const detailViewer = new PptxViewer(detailCanvas, {
      width: 800,
      useGoogleFonts: true,
      enableTextSelection: true,
      ...fullRenderers,
    });

    // Load thumbnails using PptxPresentation; detail viewer loads independently
    Promise.all([
      PptxPresentation.load(SAMPLE_URL, { useGoogleFonts: true, ...fullRenderers }),
      detailViewer.load(SAMPLE_URL),
    ])
      .then(async ([pres]) => {
        status.textContent = `Rendering ${pres.slideCount} thumbnails…`;
        const thumbEntries: HTMLDivElement[] = [];

        const selectSlide = async (i: number) => {
          for (let k = 0; k < thumbEntries.length; k++) {
            thumbEntries[k].style.outline = k === i ? '2px solid #0366d6' : 'none';
          }
          await detailViewer.goToSlide(i);
        };

        for (let i = 0; i < pres.slideCount; i++) {
          const cell = document.createElement('div');
          cell.style.cssText = 'display:flex;flex-direction:column;align-items:center;cursor:pointer;padding:4px;';
          const canvas = document.createElement('canvas');
          canvas.style.cssText =
            'display:block;width:100%;max-width:220px;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,0.2);';
          const caption = document.createElement('div');
          caption.textContent = `Slide ${i + 1}`;
          caption.style.cssText = 'font-size:12px;color:#444;margin-top:4px;';
          cell.append(canvas, caption);
          const idx = i;
          cell.addEventListener('click', () => {
            selectSlide(idx).catch((e: Error) => {
              status.textContent = `Render error: ${e.message}`;
            });
          });
          thumbCol.appendChild(cell);
          thumbEntries.push(cell);
          await pres.renderSlide(canvas, i, { width: 220 });
        }

        // Highlight first thumbnail
        if (thumbEntries.length > 0) thumbEntries[0].style.outline = '2px solid #0366d6';
        status.textContent = `Loaded ${pres.slideCount} slides`;
      })
      .catch((e: Error) => {
        status.textContent = `Error: ${e.message}`;
        status.style.color = 'red';
      });

    return root;
  },
};

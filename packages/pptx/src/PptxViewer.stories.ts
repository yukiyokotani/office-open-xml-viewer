import type { Meta, StoryObj } from '@storybook/html';
import { PptxViewer } from './viewer';
import { math } from '../../../src/math';
import { threeD } from '../../../src/three-d';
import { regionMap } from '../../../src/region-map';
import { chartEx } from '../../../src/chart-ex';

type Args = {
  width: number;
  debug?: boolean;
};

const meta: Meta<Args> = {
  title: 'PptxViewer',
  excludeStories: ['buildViewerUI', 'createCanvasSpinner'],
  argTypes: {
    width: {
      control: { type: 'range', min: 400, max: 1600, step: 40 },
      description: 'Canvas render width (px)',
    },
  },
  args: { width: 960 },
};
export default meta;
type Story = StoryObj<Args>;

// ---------------------------------------------------------------------------
// Helper: build nav bar + viewer (exported for use in local-only sample stories)
// ---------------------------------------------------------------------------
export function buildViewerUI(
  args: Args,
  autoLoadUrl?: string,
  extra?: { mode?: 'main' | 'worker' }
): { root: HTMLElement; viewer: PptxViewer } {
  const root = document.createElement('div');
  root.style.cssText = 'font-family:sans-serif;padding:16px;';

  const toolbar = document.createElement('div');
  toolbar.style.cssText = 'display:flex;gap:10px;align-items:center;margin-bottom:10px;flex-wrap:wrap;';

  const prevBtn = document.createElement('button');
  prevBtn.textContent = '← Prev';
  prevBtn.disabled = true;

  const nextBtn = document.createElement('button');
  nextBtn.textContent = 'Next →';
  nextBtn.disabled = true;

  const slideInfo = document.createElement('span');
  slideInfo.style.fontSize = '14px';

  const status = document.createElement('div');
  status.style.cssText = 'color:#666;font-size:13px;margin-bottom:8px;min-height:18px;';

  toolbar.append(prevBtn, nextBtn, slideInfo);
  root.append(toolbar, status);

  const container = document.createElement('div');
  container.style.cssText =
    `position:relative;width:${args.width}px;max-width:100%;border:1px solid #ccc;background:#f0f0f0;min-height:120px;`;
  root.appendChild(container);

  const spinner = createCanvasSpinner();
  container.appendChild(spinner);

  const canvas = document.createElement('canvas');
  container.appendChild(canvas);

  const viewer = new PptxViewer(canvas, {
    width: args.width,
    useGoogleFonts: true,
    enableTextSelection: true,
    math,
    threeD,
    regionMap,
    chartEx,
    onSlideChange: (idx, total) => {
      slideInfo.textContent = `Slide ${idx + 1} / ${total}`;
      prevBtn.disabled = idx === 0;
      nextBtn.disabled = idx === total - 1;
    },
    onError: (err) => { status.textContent = `Error: ${err.message}`; },
    ...extra,
  });

  prevBtn.addEventListener('click', () => viewer.prevSlide());
  nextBtn.addEventListener('click', () => viewer.nextSlide());

  if (autoLoadUrl) {
    status.textContent = 'Loading…';
    viewer.load(autoLoadUrl)
      .then(() => {
        status.textContent = 'Loaded';
        spinner.remove();
      })
      .catch((err) => {
        status.textContent = `Failed: ${err.message}`;
        spinner.remove();
      });
  } else {
    spinner.remove();
  }

  return { root, viewer };
}

/**
 * Returns an absolutely-positioned spinner overlay. The element is a simple
 * CSS-keyframe ring centered in its parent — the parent must be positioned
 * (set `position:relative`) so the overlay anchors correctly.
 */
export function createCanvasSpinner(): HTMLElement {
  const el = document.createElement('div');
  el.setAttribute('aria-label', 'Loading');
  el.style.cssText = [
    'position:absolute',
    'top:50%', 'left:50%',
    'width:40px', 'height:40px',
    'margin:-20px 0 0 -20px',
    'border:3px solid rgba(0,0,0,0.12)',
    'border-top-color:rgba(0,0,0,0.55)',
    'border-radius:50%',
    'pointer-events:none',
    'animation:pptxSpinnerRotate 0.9s linear infinite',
  ].join(';');
  // Inject keyframes once per document.
  const keyframesId = '__pptx-spinner-keyframes';
  if (!document.getElementById(keyframesId)) {
    const style = document.createElement('style');
    style.id = keyframesId;
    style.textContent = '@keyframes pptxSpinnerRotate { to { transform: rotate(360deg); } }';
    document.head.appendChild(style);
  }
  return el;
}

// ---------------------------------------------------------------------------
// File-upload viewer (shared by main-thread and Web Worker stories)
// ---------------------------------------------------------------------------
function renderFileUpload(args: Args, mode: 'main' | 'worker'): HTMLElement {
    const root = document.createElement('div');
    root.style.cssText = 'font-family:sans-serif;padding:16px;';

    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.pptx';
    fileInput.style.marginBottom = '12px';

    const status = document.createElement('div');
    status.style.cssText = 'color:#666;font-size:13px;margin-bottom:8px;min-height:18px;';

    const toolbar = document.createElement('div');
    toolbar.style.cssText = 'display:flex;gap:10px;align-items:center;margin-bottom:10px;flex-wrap:wrap;';

    const prevBtn = document.createElement('button');
    prevBtn.textContent = '← Prev';
    prevBtn.disabled = true;

    const nextBtn = document.createElement('button');
    nextBtn.textContent = 'Next →';
    nextBtn.disabled = true;

    const slideInfo = document.createElement('span');
    slideInfo.style.fontSize = '14px';

    toolbar.append(prevBtn, nextBtn, slideInfo);

    const container = document.createElement('div');
    container.style.cssText =
      `width:${args.width}px;max-width:100%;border:1px solid #ccc;background:#f0f0f0;` +
      `display:flex;align-items:center;justify-content:center;min-height:200px;`;
    const hint = document.createElement('span');
    hint.textContent = 'Drop a .pptx here or use the chooser above';
    hint.style.color = '#aaa';
    container.appendChild(hint);

    root.append(fileInput, status, toolbar, container);

    let viewer: PptxViewer | null = null;

    async function loadBuffer(name: string, buffer: ArrayBuffer) {
      status.textContent = `Parsing ${name}…`;
      viewer?.destroy();
      container.innerHTML = '';
      const canvas = document.createElement('canvas');
      container.appendChild(canvas);
      viewer = new PptxViewer(canvas, {
        mode,
        width: args.width,
        debug: args.debug,
        enableMediaPlayback: true,
        math,
        threeD,
        regionMap,
        chartEx,
        onSlideChange: (idx, total) => {
          slideInfo.textContent = `Slide ${idx + 1} / ${total}`;
          prevBtn.disabled = idx === 0;
          nextBtn.disabled = idx === total - 1;
        },
        onError: (err) => { status.textContent = `Error: ${err.message}`; },
      });
      try {
        await viewer.load(buffer);
        status.textContent = `Loaded ${name}`;
      } catch (err) {
        status.textContent = `Failed: ${err instanceof Error ? err.message : String(err)}`;
      }
    }

    fileInput.addEventListener('change', async () => {
      const file = fileInput.files?.[0];
      if (!file) return;
      loadBuffer(file.name, await file.arrayBuffer());
    });

    root.addEventListener('dragover', (e) => e.preventDefault());
    root.addEventListener('drop', async (e) => {
      e.preventDefault();
      const file = e.dataTransfer?.files[0];
      if (file?.name.endsWith('.pptx')) {
        loadBuffer(file.name, await file.arrayBuffer());
      }
    });

    prevBtn.addEventListener('click', () => viewer?.prevSlide());
    nextBtn.addEventListener('click', () => viewer?.nextSlide());

    return root;
}

const fileUploadArgTypes = {
  debug: {
    control: 'boolean' as const,
    description: 'Print resource-usage metrics to the browser console',
  },
};

export const FileUpload: Story = {
  name: 'Load from file — main thread',
  args: { width: 960, debug: true },
  argTypes: fileUploadArgTypes,
  render: (args) => renderFileUpload(args, 'main'),
};

export const FileUploadWorker: Story = {
  name: 'Load from file — Web Worker',
  args: { width: 960, debug: true },
  argTypes: fileUploadArgTypes,
  render: (args) => renderFileUpload(args, 'worker'),
};

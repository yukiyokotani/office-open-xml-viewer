import type { Meta, StoryObj } from '@storybook/html';
import { DocxDocument } from './document';
import { DocxViewer } from './viewer';
import { math } from '../../../src/math';
import { threeD } from '../../../src/three-d';
import { regionMap } from '../../../src/region-map';
import { chartEx } from '../../../src/chart-ex';

type Args = {
  width: number;
  debug?: boolean;
};

const meta: Meta<Args> = {
  title: 'DocxViewer',
  excludeStories: ['buildViewerUI'],
  argTypes: {
    width: {
      control: { type: 'range', min: 400, max: 1200, step: 40 },
      description: 'Canvas render width (px)',
    },
  },
  args: { width: 700 },
};
export default meta;
type Story = StoryObj<Args>;

// ---------------------------------------------------------------------------
// Helper: build nav bar + viewer (exported for use in local-only sample stories)
// ---------------------------------------------------------------------------
export function buildViewerUI(
  args: Args,
  autoLoadUrl?: string,
  extra?: { mode?: 'main' | 'worker' },
): { root: HTMLElement; doc: DocxDocument | null } {
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

  const pageInfo = document.createElement('span');
  pageInfo.style.fontSize = '14px';

  const status = document.createElement('div');
  status.style.cssText = 'color:#666;font-size:13px;margin-bottom:8px;min-height:18px;';

  toolbar.append(prevBtn, nextBtn, pageInfo);
  root.append(toolbar, status);

  const container = document.createElement('div');
  container.style.cssText =
    `position:relative;width:${args.width}px;max-width:100%;border:1px solid #ccc;background:#f0f0f0;min-height:120px;`;
  root.appendChild(container);

  const spinner = createCanvasSpinner();
  container.appendChild(spinner);

  const canvas = document.createElement('canvas');
  container.appendChild(canvas);

  const viewer = new DocxViewer(canvas, {
    width: args.width,
    dpr: window.devicePixelRatio,
    enableTextSelection: true,
    useGoogleFonts: true,
    math,
    threeD,
    regionMap,
    chartEx,
    ...extra,
  });

  const updateNav = () => {
    const total = viewer.pageCount;
    pageInfo.textContent = total > 0 ? `Page ${viewer.currentPage + 1} / ${total}` : '';
    prevBtn.disabled = viewer.currentPage <= 0;
    nextBtn.disabled = viewer.currentPage >= total - 1;
  };

  prevBtn.addEventListener('click', () => { viewer.prevPage(); updateNav(); });
  nextBtn.addEventListener('click', () => { viewer.nextPage(); updateNav(); });

  if (autoLoadUrl) {
    status.textContent = `Loading ${autoLoadUrl}…`;
    viewer.load(autoLoadUrl)
      .then(() => {
        status.textContent = `Loaded — ${viewer.pageCount} page(s)`;
        updateNav();
        spinner.remove();
      })
      .catch((e: Error) => {
        status.textContent = `Error: ${e.message}`;
        status.style.color = 'red';
        spinner.remove();
      });
  } else {
    spinner.remove();
  }

  return { root, doc: null };
}

/**
 * Absolutely-positioned spinner overlay (CSS-keyframe ring) centered in its
 * parent — the parent must be positioned. Shown during the load (notably the
 * worker-mode cold start) and removed once the first page renders.
 */
function createCanvasSpinner(): HTMLElement {
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
    'animation:ooxmlSpinnerRotate 0.9s linear infinite',
  ].join(';');
  const keyframesId = '__ooxml-spinner-keyframes';
  if (!document.getElementById(keyframesId)) {
    const style = document.createElement('style');
    style.id = keyframesId;
    style.textContent = '@keyframes ooxmlSpinnerRotate { to { transform: rotate(360deg); } }';
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
    fileInput.accept = '.docx';
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

    const pageInfo = document.createElement('span');
    pageInfo.style.fontSize = '14px';

    toolbar.append(prevBtn, nextBtn, pageInfo);

    const container = document.createElement('div');
    container.style.cssText =
      `width:${args.width}px;max-width:100%;border:1px solid #ccc;background:#f0f0f0;` +
      `display:flex;align-items:center;justify-content:center;min-height:200px;`;
    const hint = document.createElement('span');
    hint.textContent = 'Drop a .docx here or use the chooser above';
    hint.style.color = '#aaa';
    container.appendChild(hint);

    root.append(fileInput, status, toolbar, container);

    let viewer: DocxViewer | null = null;

    const updateNav = () => {
      const total = viewer?.pageCount ?? 0;
      pageInfo.textContent = total > 0 ? `Page ${(viewer?.currentPage ?? 0) + 1} / ${total}` : '';
      prevBtn.disabled = (viewer?.currentPage ?? 0) <= 0;
      nextBtn.disabled = (viewer?.currentPage ?? 0) >= total - 1;
    };

    async function loadBuffer(name: string, buffer: ArrayBuffer) {
      status.textContent = `Parsing ${name}…`;
      container.innerHTML = '';
      const canvas = document.createElement('canvas');
      container.appendChild(canvas);
      viewer = new DocxViewer(canvas, {
        mode,
        width: args.width,
        dpr: window.devicePixelRatio,
        debug: args.debug,
        enableTextSelection: true,
        useGoogleFonts: true,
        math,
        threeD,
        regionMap,
        chartEx,
      });
      try {
        await viewer.load(buffer);
        status.textContent = `Loaded ${name} — ${viewer.pageCount} page(s)`;
        updateNav();
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
      if (file?.name.endsWith('.docx')) {
        loadBuffer(file.name, await file.arrayBuffer());
      }
    });

    prevBtn.addEventListener('click', () => { viewer?.prevPage(); updateNav(); });
    nextBtn.addEventListener('click', () => { viewer?.nextPage(); updateNav(); });

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
  args: { width: 700, debug: true },
  argTypes: fileUploadArgTypes,
  render: (args) => renderFileUpload(args, 'main'),
};

export const FileUploadWorker: Story = {
  name: 'Load from file — Web Worker',
  args: { width: 700, debug: true },
  argTypes: fileUploadArgTypes,
  render: (args) => renderFileUpload(args, 'worker'),
};


export const TrackedChanges: Story = {
  name: 'Tracked changes (\u00a717.13.5 markup view)',
  args: { width: 700 },
  render: (args) => {
    const root = document.createElement('div');
    root.style.cssText = 'font-family:sans-serif;padding:16px;';

    const toolbar = document.createElement('div');
    toolbar.style.cssText = 'display:flex;gap:10px;align-items:center;margin-bottom:10px;flex-wrap:wrap;';
    const markupBtn = document.createElement('button');
    const status = document.createElement('div');
    status.style.cssText = 'color:#666;font-size:13px;margin-bottom:8px;min-height:18px;';
    toolbar.append(markupBtn);
    root.append(toolbar, status);

    const container = document.createElement('div');
    container.style.cssText =
      `position:relative;width:${args.width}px;max-width:100%;border:1px solid #ccc;`
      + 'background:#f0f0f0;min-height:120px;overflow:auto;padding:8px;';
    root.appendChild(container);

    let viewer: DocxViewer | null = null;
    let markupOn = false;
    const refreshButtons = () => {
      markupBtn.textContent = `Markup view: ${markupOn ? 'on' : 'off'}`;
    };
    refreshButtons();

    async function loadFixture(): Promise<void> {
      const { generateTrackedChangesDocx } = await import('./conformance/generate.js');
      const bytes = generateTrackedChangesDocx();
      container.innerHTML = '';
      const canvas = document.createElement('canvas');
      container.appendChild(canvas);
      viewer = new DocxViewer(canvas, {
        width: args.width,
        dpr: window.devicePixelRatio,
        enableTextSelection: true,
        showTrackedChanges: markupOn,
      });
      const copy = bytes.buffer.slice(
        bytes.byteOffset,
        bytes.byteOffset + bytes.byteLength,
      ) as ArrayBuffer;
      await viewer.load(copy);
      status.textContent =
        'Synthetic redistributable fixture \u2014 toggle the markup view: deletions appear '
        + 'struck through and insertions underlined in per-author colours, with margin change bars.';
    }

    markupBtn.addEventListener('click', () => {
      markupOn = !markupOn;
      markupBtn.disabled = true;
      // The markup view is a different pagination, built in the worker under
      // mode: 'worker'. Await the switch, so the button reports the view that
      // is actually on screen rather than the one that was requested.
      void viewer?.setShowTrackedChanges(markupOn)
        .catch((error: unknown) => {
          markupOn = !markupOn;
          status.textContent = `Markup view switch failed: ${String(error)}`;
        })
        .finally(() => {
          markupBtn.disabled = false;
          refreshButtons();
        });
    });
    void loadFixture();
    return root;
  },
};

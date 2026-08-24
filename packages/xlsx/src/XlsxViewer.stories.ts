import type { Meta, StoryObj } from '@storybook/html';
import { XlsxViewer } from './viewer';
import { math } from '../../../src/math';
import { threeD } from '../../../src/three-d';
import { regionMap } from '../../../src/region-map';
import { chartEx } from '../../../src/chart-ex';

type Args = {
  scale: number;
  selectionColor: string;
  debug?: boolean;
};

const meta: Meta<Args> = {
  title: 'XlsxViewer',
  excludeStories: ['buildViewerUI'],
  parameters: { layout: 'fullscreen' },
  argTypes: {
    scale: {
      control: { type: 'range', min: 0.25, max: 2, step: 0.05 },
      description: 'Cell/header scale (1 = normal size)',
    },
    selectionColor: {
      control: { type: 'color' },
      description: 'Cell-selection accent (border + translucent fill)',
    },
  },
  args: { scale: 1, selectionColor: '#1a73e8' },
};
export default meta;
type Story = StoryObj<Args>;

// ---------------------------------------------------------------------------
// Helper: build full-viewport viewer UI (exported for use in sample stories)
// ---------------------------------------------------------------------------
export function buildViewerUI(
  args: Args,
  autoLoadUrl?: string,
  extra?: { mode?: 'main' | 'worker' },
): { root: HTMLElement; viewer: XlsxViewer } {
  const root = document.createElement('div');
  root.style.cssText = 'width:100%;height:100vh;display:flex;flex-direction:column;overflow:hidden;font-family:sans-serif;box-sizing:border-box;';

  const viewerContainer = document.createElement('div');
  viewerContainer.style.cssText = 'position:relative;flex:1;min-height:0;';
  root.appendChild(viewerContainer);

  const viewer = new XlsxViewer(viewerContainer, {
    cellScale: args.scale,
    selectionColor: args.selectionColor,
    useGoogleFonts: true,
    math,
    threeD,
    regionMap,
    chartEx,
    ...extra,
  });

  // Overlay a spinner during load (notably the worker-mode cold start); the
  // viewer builds its own DOM inside viewerContainer, so append the spinner
  // last so it sits on top.
  const spinner = createCanvasSpinner();
  viewerContainer.appendChild(spinner);

  if (autoLoadUrl) {
    viewer.load(autoLoadUrl)
      .then(() => { spinner.remove(); })
      .catch((err) => {
        spinner.remove();
        viewerContainer.textContent = `Failed: ${err.message}`;
      });
  } else {
    spinner.remove();
  }

  return { root, viewer };
}

// ---------------------------------------------------------------------------
// File upload (shared by main-thread and Web Worker stories)
// ---------------------------------------------------------------------------
function renderFileUpload(args: Args, mode: 'main' | 'worker'): HTMLElement {
    const root = document.createElement('div');
    root.style.cssText = 'width:100%;height:100vh;display:flex;flex-direction:column;overflow:hidden;font-family:sans-serif;box-sizing:border-box;';

    const toolbar = document.createElement('div');
    toolbar.style.cssText = 'display:flex;align-items:center;gap:8px;padding:4px 8px;height:32px;flex-shrink:0;';

    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.xlsx';

    const status = document.createElement('div');
    status.style.cssText = 'color:#666;font-size:12px;';

    toolbar.append(fileInput, status);
    root.appendChild(toolbar);

    const viewerContainer = document.createElement('div');
    viewerContainer.style.cssText = 'flex:1;min-height:0;';
    root.appendChild(viewerContainer);

    let viewer: XlsxViewer | null = null;
    async function loadBuffer(buf: ArrayBuffer) {
      viewer?.destroy();
      viewerContainer.innerHTML = '';
      viewer = new XlsxViewer(viewerContainer, {
        mode,
        cellScale: args.scale,
        debug: args.debug,
        useGoogleFonts: true,
        math,
        threeD,
        regionMap,
        chartEx,
        onReady: () => { status.textContent = ''; },
        onError: (err) => { status.textContent = `Error: ${err.message}`; },
      });
      await viewer.load(buf);
    }

    fileInput.addEventListener('change', async () => {
      const file = fileInput.files?.[0];
      if (!file) return;
      status.textContent = 'Parsing…';
      loadBuffer(await file.arrayBuffer());
    });

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
  args: { debug: true },
  argTypes: fileUploadArgTypes,
  render: (args) => renderFileUpload(args, 'main'),
};

export const FileUploadWorker: Story = {
  name: 'Load from file — Web Worker',
  args: { debug: true },
  argTypes: fileUploadArgTypes,
  render: (args) => renderFileUpload(args, 'worker'),
};

// ---------------------------------------------------------------------------
// Interactions: drag-to-resize columns/rows, Ctrl+wheel zoom, selection color
// ---------------------------------------------------------------------------
/**
 * Auto-loads the bundled demo sheet to exercise the view-only interactions:
 *  - **Resize**: drag a column-header or row-header border to resize it. The
 *    cursor turns into a col-/row-resize handle over the border. Resizing is
 *    held in viewer state and never written back to the file (issue #567).
 *  - **Zoom**: Ctrl/⌘ + mouse wheel (or trackpad pinch) zooms the grid.
 *  - **Selection color**: change the `selectionColor` control to recolor the
 *    selection rectangle (border + translucent fill).
 */
export const Interactions: Story = {
  name: 'Resize / zoom / selection color',
  render(args) {
    const { root } = buildViewerUI(args, '/xlsx/demo/sample-1.xlsx');
    return root;
  },
};


/**
 * Absolutely-positioned spinner overlay (CSS-keyframe ring) centered in its
 * parent — the parent must be positioned. Shown during the load (notably the
 * worker-mode cold start) and removed once the first sheet renders.
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
    'z-index:5',
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

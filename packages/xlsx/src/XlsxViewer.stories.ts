import type { Meta, StoryObj } from '@storybook/html';
import { XlsxViewer } from './viewer';
import init, { parse_xlsx } from './wasm/xlsx_parser.js';

type Args = {
  scale: number;
};

const meta: Meta<Args> = {
  title: 'XlsxViewer',
  excludeStories: ['buildViewerUI'],
  argTypes: {
    scale: {
      control: { type: 'range', min: 0.25, max: 2, step: 0.05 },
      description: 'Cell/header scale (1 = normal size)',
    },
  },
  args: { scale: 1 },
};
export default meta;
type Story = StoryObj<Args>;

// ---------------------------------------------------------------------------
// Helper: build full-viewport viewer UI (exported for use in sample stories)
// ---------------------------------------------------------------------------
export function buildViewerUI(
  args: Args,
  autoLoadUrl?: string,
): { root: HTMLElement; viewer: XlsxViewer } {
  const root = document.createElement('div');
  root.style.cssText = 'width:100%;height:100vh;display:flex;flex-direction:column;overflow:hidden;font-family:sans-serif;box-sizing:border-box;';

  const status = document.createElement('div');
  status.style.cssText = 'padding:4px 8px;color:#666;font-size:12px;height:24px;flex-shrink:0;display:flex;align-items:center;';
  root.appendChild(status);

  const viewerContainer = document.createElement('div');
  viewerContainer.style.cssText = 'flex:1;min-height:0;';
  root.appendChild(viewerContainer);

  const viewer = new XlsxViewer(viewerContainer, {
    cellScale: args.scale,
    onReady: (names) => {
      status.textContent = `Loaded — ${names.length} sheet(s)`;
    },
    onSheetChange: (_idx, name) => {
      status.textContent = `Sheet: ${name}`;
    },
    onError: (err) => { status.textContent = `Error: ${err.message}`; },
  });

  if (autoLoadUrl) {
    status.textContent = 'Loading…';
    fetch(autoLoadUrl)
      .then((r) => {
        if (!r.ok) throw new Error(`HTTP ${r.status}`);
        return r.arrayBuffer();
      })
      .then((buf) => viewer.load(buf))
      .catch((err) => { status.textContent = `Failed: ${err.message}`; });
  }

  return { root, viewer };
}

// ---------------------------------------------------------------------------
// Debug: raw JSON from WASM parser
// ---------------------------------------------------------------------------
export const DebugJson: Story = {
  name: 'Debug – raw parse JSON',
  render(_args) {
    const root = document.createElement('div');
    root.style.cssText = 'font-family:sans-serif;padding:16px;';

    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.xlsx';

    const pre = document.createElement('pre');
    pre.style.cssText =
      'font-size:11px;line-height:1.4;max-height:600px;overflow:auto;' +
      'background:#1e1e1e;color:#d4d4d4;padding:12px;border-radius:4px;margin-top:12px;';
    pre.textContent = 'Load an .xlsx to see the parsed JSON here.';

    root.append(fileInput, pre);

    let wasmReady = false;
    init().then(() => { wasmReady = true; });

    fileInput.addEventListener('change', async () => {
      const file = fileInput.files?.[0];
      if (!file || !wasmReady) return;
      try {
        const buf = await file.arrayBuffer();
        const json = parse_xlsx(new Uint8Array(buf));
        const parsed = JSON.parse(json);
        pre.textContent = JSON.stringify(parsed, null, 2);
        console.log('[xlsx debug] full JSON:', parsed);
      } catch (err) {
        pre.textContent = `Error: ${err instanceof Error ? err.message : String(err)}`;
      }
    });

    return root;
  },
};

// ---------------------------------------------------------------------------
// File upload
// ---------------------------------------------------------------------------
export const FileUpload: Story = {
  name: 'Load from file',
  render(args) {
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
        cellScale: args.scale,
        onReady: (names) => { status.textContent = `${names.length} sheet(s)`; },
        onSheetChange: (_idx, name) => { status.textContent = `Sheet: ${name}`; },
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
  },
};

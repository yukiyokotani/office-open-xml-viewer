import type { Meta, StoryObj } from '@storybook/html';
import { XlsxViewer } from './viewer';
import init, { parse_xlsx } from './wasm/xlsx_parser.js';

type Args = {
  width: number;
  height: number;
  scale: number;
};

const meta: Meta<Args> = {
  title: 'XlsxViewer',
  argTypes: {
    width: {
      control: { type: 'range', min: 400, max: 1920, step: 40 },
      description: 'Canvas render width (px)',
    },
    height: {
      control: { type: 'range', min: 200, max: 1200, step: 40 },
      description: 'Canvas render height (px)',
    },
    scale: {
      control: { type: 'range', min: 0.25, max: 1, step: 0.05 },
      description: 'Display scale (CSS transform)',
    },
  },
  args: { width: 1200, height: 600, scale: 0.75 },
};
export default meta;
type Story = StoryObj<Args>;

// ---------------------------------------------------------------------------
// Helper: build viewer UI with sheet switcher
// ---------------------------------------------------------------------------
function buildViewerUI(
  args: Args,
  autoLoadUrl?: string,
): { root: HTMLElement; viewer: XlsxViewer } {
  const scale = args.scale ?? 0.75;
  const root = document.createElement('div');
  root.style.cssText = 'font-family:sans-serif;padding:16px;';

  const status = document.createElement('div');
  status.style.cssText = 'color:#666;font-size:13px;margin-bottom:8px;min-height:18px;';
  root.appendChild(status);

  const scaleWrap = document.createElement('div');
  scaleWrap.style.cssText = `
    display: inline-block;
    transform: scale(${scale});
    transform-origin: top left;
    width: ${args.width}px;
    height: ${args.height + 30}px;
  `.replace(/\n\s+/g, ' ').trim();
  root.appendChild(scaleWrap);

  const container = document.createElement('div');
  container.style.cssText = `max-width:100%;`;
  scaleWrap.appendChild(container);

  const viewer = new XlsxViewer(container, {
    width: args.width,
    height: args.height,
    onReady: (names) => {
      status.textContent = `Loaded — ${names.length} sheet(s)`;
    },
    onSheetChange: (_idx, name) => {
      status.textContent = `Sheet: ${name}`;
    },
    onError: (err) => { status.textContent = `Error: ${err.message}`; },
  });

  // Shrink the outer wrapper so the scaled content doesn't leave whitespace
  scaleWrap.style.marginBottom = `${-(args.height + 30) * (1 - scale)}px`;

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
// Sample file stories
// ---------------------------------------------------------------------------
const SAMPLES = [
  'sample-1.xlsx',
  'sample-2.xlsx',
  'sample-3.xlsx',
  'sample-4.xlsx',
  'sample-5.xlsx',
  'sample-6.xlsx',
] as const;

export const Sample1: Story = {
  name: 'sample-1.xlsx',
  render: (args) => buildViewerUI(args, `/${SAMPLES[0]}`).root,
};

export const Sample2: Story = {
  name: 'sample-2.xlsx',
  render: (args) => buildViewerUI(args, `/${SAMPLES[1]}`).root,
};

export const Sample3: Story = {
  name: 'sample-3.xlsx',
  render: (args) => buildViewerUI(args, `/${SAMPLES[2]}`).root,
};

export const Sample4: Story = {
  name: 'sample-4.xlsx',
  render: (args) => buildViewerUI(args, `/${SAMPLES[3]}`).root,
};

export const Sample5: Story = {
  name: 'sample-5.xlsx',
  render: (args) => buildViewerUI(args, `/${SAMPLES[4]}`).root,
};

export const Sample6: Story = {
  name: 'sample-6.xlsx',
  render: (args) => buildViewerUI(args, `/${SAMPLES[5]}`).root,
};

export const SampleCF: Story = {
  name: 'sample-cf.xlsx (conditional formatting)',
  render: (args) => buildViewerUI(args, '/sample-cf.xlsx').root,
};

// ---------------------------------------------------------------------------
// File upload
// ---------------------------------------------------------------------------
export const FileUpload: Story = {
  name: 'Load from file',
  render(args) {
    const scale = args.scale ?? 0.75;
    const root = document.createElement('div');
    root.style.cssText = 'font-family:sans-serif;padding:16px;';

    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.xlsx';
    fileInput.style.marginBottom = '12px';

    const status = document.createElement('div');
    status.style.cssText = 'color:#666;font-size:13px;margin-bottom:8px;min-height:18px;';

    const scaleWrap = document.createElement('div');
    scaleWrap.style.cssText = `
      display: inline-block;
      transform: scale(${scale});
      transform-origin: top left;
      width: ${args.width}px;
      margin-bottom: ${-(args.height + 30) * (1 - scale)}px;
    `.replace(/\n\s+/g, ' ').trim();

    const container = document.createElement('div');
    container.style.cssText = `max-width:100%;min-height:200px;`;
    scaleWrap.appendChild(container);
    const hint = document.createElement('span');
    hint.textContent = 'Select an .xlsx file above';
    hint.style.cssText = 'display:block;padding:20px;color:#aaa;';
    container.appendChild(hint);

    root.append(fileInput, status, scaleWrap);

    let viewer: XlsxViewer | null = null;

    async function loadBuffer(buf: ArrayBuffer) {
      viewer?.destroy();
      container.innerHTML = '';
      viewer = new XlsxViewer(container, {
        width: args.width,
        height: args.height,
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

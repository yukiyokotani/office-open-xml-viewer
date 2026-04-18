import type { Meta, StoryObj } from '@storybook/html';
import { XlsxViewer } from './viewer';
import init, { parse_xlsx } from './wasm/xlsx_parser.js';

type Args = {
  width: number;
  height: number;
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
  },
  args: { width: 1200, height: 600 },
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
  const root = document.createElement('div');
  root.style.cssText = 'font-family:sans-serif;padding:16px;';

  const toolbar = document.createElement('div');
  toolbar.style.cssText = 'display:flex;gap:10px;align-items:center;margin-bottom:10px;flex-wrap:wrap;';

  const sheetSelect = document.createElement('select');
  sheetSelect.style.cssText = 'padding:4px 8px;border-radius:4px;border:1px solid #ccc;font-size:13px;';
  sheetSelect.disabled = true;

  const status = document.createElement('div');
  status.style.cssText = 'color:#666;font-size:13px;margin-bottom:8px;min-height:18px;';

  toolbar.append(sheetSelect);
  root.append(toolbar, status);

  const container = document.createElement('div');
  container.style.cssText =
    `width:${args.width}px;max-width:100%;border:1px solid #ccc;background:#f0f0f0;min-height:120px;`;
  root.appendChild(container);

  const viewer = new XlsxViewer(container, {
    width: args.width,
    height: args.height,
    onReady: (names) => {
      sheetSelect.innerHTML = '';
      names.forEach((name, i) => {
        const opt = document.createElement('option');
        opt.value = String(i);
        opt.textContent = name;
        sheetSelect.appendChild(opt);
      });
      sheetSelect.disabled = false;
      status.textContent = `Loaded — ${names.length} sheet(s)`;
    },
    onSheetChange: (idx, name) => {
      sheetSelect.value = String(idx);
      status.textContent = `Sheet: ${name}`;
    },
    onError: (err) => { status.textContent = `Error: ${err.message}`; },
  });

  sheetSelect.addEventListener('change', () => {
    viewer.showSheet(Number(sheetSelect.value));
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

// ---------------------------------------------------------------------------
// File upload
// ---------------------------------------------------------------------------
export const FileUpload: Story = {
  name: 'Load from file',
  render(args) {
    const root = document.createElement('div');
    root.style.cssText = 'font-family:sans-serif;padding:16px;';

    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.xlsx';
    fileInput.style.marginBottom = '12px';

    const status = document.createElement('div');
    status.style.cssText = 'color:#666;font-size:13px;margin-bottom:8px;min-height:18px;';

    const sheetSelect = document.createElement('select');
    sheetSelect.style.cssText = 'padding:4px 8px;border-radius:4px;border:1px solid #ccc;font-size:13px;display:none;';

    const toolbar = document.createElement('div');
    toolbar.style.cssText = 'display:flex;gap:10px;align-items:center;margin-bottom:10px;';
    toolbar.appendChild(sheetSelect);

    const container = document.createElement('div');
    container.style.cssText =
      `width:${args.width}px;max-width:100%;border:1px solid #ccc;background:#f0f0f0;` +
      `min-height:200px;display:flex;align-items:center;justify-content:center;`;
    const hint = document.createElement('span');
    hint.textContent = 'Select an .xlsx file above';
    hint.style.color = '#aaa';
    container.appendChild(hint);

    root.append(fileInput, status, toolbar, container);

    let viewer: XlsxViewer | null = null;

    async function loadBuffer(buf: ArrayBuffer) {
      viewer?.destroy();
      container.innerHTML = '';
      viewer = new XlsxViewer(container, {
        width: args.width,
        height: args.height,
        onReady: (names) => {
          sheetSelect.innerHTML = '';
          names.forEach((name, i) => {
            const opt = document.createElement('option');
            opt.value = String(i);
            opt.textContent = name;
            sheetSelect.appendChild(opt);
          });
          sheetSelect.style.display = 'block';
          status.textContent = `${names.length} sheet(s)`;
        },
        onSheetChange: (idx, name) => {
          sheetSelect.value = String(idx);
          status.textContent = `Sheet: ${name}`;
        },
        onError: (err) => { status.textContent = `Error: ${err.message}`; },
      });
      sheetSelect.onchange = () => viewer?.showSheet(Number(sheetSelect.value));
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

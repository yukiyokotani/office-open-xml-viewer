import type { Meta, StoryObj } from '@storybook/html';
import { XlsxViewer } from './viewer';
import type { CellRange } from './viewer';
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

// ---------------------------------------------------------------------------
// Selectable Viewer (cell selection + Ctrl+C copy)
// ---------------------------------------------------------------------------

function buildSelectableUI(args: Args, autoLoadUrl?: string): HTMLElement {
  const root = document.createElement('div');
  root.style.cssText =
    'width:100%;height:100vh;display:flex;flex-direction:column;overflow:hidden;font-family:sans-serif;box-sizing:border-box;';

  const toolbar = document.createElement('div');
  toolbar.style.cssText =
    'display:flex;align-items:center;gap:12px;padding:4px 12px;height:36px;flex-shrink:0;' +
    'background:#f8f9fa;border-bottom:1px solid #dadce0;font-size:13px;';

  const fileInput = document.createElement('input');
  fileInput.type = 'file';
  fileInput.accept = '.xlsx';
  fileInput.style.cssText = 'font-size:12px;';

  const selectionLabel = document.createElement('span');
  selectionLabel.style.cssText = 'color:#555;flex:1;';
  selectionLabel.textContent = 'Click or drag to select cells. Ctrl+C to copy.';

  const copyBtn = document.createElement('button');
  copyBtn.textContent = 'Copy selection';
  copyBtn.style.cssText =
    'padding:2px 10px;font-size:12px;cursor:pointer;border:1px solid #c8ccd0;border-radius:3px;background:#fff;';
  copyBtn.disabled = true;

  toolbar.append(fileInput, selectionLabel, copyBtn);
  root.appendChild(toolbar);

  const viewerContainer = document.createElement('div');
  viewerContainer.style.cssText = 'flex:1;min-height:0;';
  root.appendChild(viewerContainer);

  let currentViewer: XlsxViewer | null = null;

  function onSelectionChange(sel: CellRange | null): void {
    if (!sel) {
      selectionLabel.textContent = 'Click/drag cells · Click row/col headers · Click corner for all · Shift+click to extend · Ctrl+C to copy';
      copyBtn.disabled = true;
      return;
    }
    const colLabel = (n: number) => {
      let s = '';
      while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
      return s;
    };
    if (sel.mode === 'all') {
      selectionLabel.textContent = 'Selected: all cells';
    } else if (sel.mode === 'rows') {
      const r1 = Math.min(sel.anchor.row, sel.active.row);
      const r2 = Math.max(sel.anchor.row, sel.active.row);
      selectionLabel.textContent = r1 === r2 ? `Selected: row ${r1}` : `Selected: rows ${r1}–${r2}`;
    } else if (sel.mode === 'cols') {
      const c1 = Math.min(sel.anchor.col, sel.active.col);
      const c2 = Math.max(sel.anchor.col, sel.active.col);
      selectionLabel.textContent = c1 === c2 ? `Selected: col ${colLabel(c1)}` : `Selected: cols ${colLabel(c1)}–${colLabel(c2)}`;
    } else {
      const r1 = Math.min(sel.anchor.row, sel.active.row);
      const r2 = Math.max(sel.anchor.row, sel.active.row);
      const c1 = Math.min(sel.anchor.col, sel.active.col);
      const c2 = Math.max(sel.anchor.col, sel.active.col);
      const tl = `${colLabel(c1)}${r1}`;
      const br = `${colLabel(c2)}${r2}`;
      selectionLabel.textContent = r1 === r2 && c1 === c2 ? `Selected: ${tl}` : `Selected: ${tl}:${br} (${r2 - r1 + 1}×${c2 - c1 + 1})`;
    }
    copyBtn.disabled = false;
  }

  function createViewer(): XlsxViewer {
    viewerContainer.innerHTML = '';
    return new XlsxViewer(viewerContainer, {
      cellScale: args.scale,
      onReady: (names) => { selectionLabel.textContent = `Loaded — ${names.length} sheet(s). Click a cell to select.`; },
      onSheetChange: (_idx, name) => { selectionLabel.textContent = `Sheet: ${name} — Click to select.`; },
      onError: (err) => { selectionLabel.textContent = `Error: ${err.message}`; },
      onSelectionChange,
    });
  }

  copyBtn.addEventListener('click', () => {
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'c', ctrlKey: true, bubbles: true }));
    copyBtn.textContent = 'Copied!';
    setTimeout(() => { copyBtn.textContent = 'Copy selection'; }, 1200);
  });

  fileInput.addEventListener('change', async () => {
    const file = fileInput.files?.[0];
    if (!file) return;
    selectionLabel.textContent = 'Parsing…';
    currentViewer?.destroy();
    currentViewer = createViewer();
    await currentViewer.load(await file.arrayBuffer());
  });

  currentViewer = createViewer();

  if (autoLoadUrl) {
    selectionLabel.textContent = 'Loading…';
    fetch(autoLoadUrl)
      .then((r) => { if (!r.ok) throw new Error(`HTTP ${r.status}`); return r.arrayBuffer(); })
      .then((buf) => currentViewer!.load(buf))
      .catch((err) => { selectionLabel.textContent = `Failed: ${err.message}`; });
  }

  return root;
}

export const SelectableFileUpload: Story = {
  name: 'Selectable — file upload',
  render(args) { return buildSelectableUI(args); },
};

export const SelectableSample1: Story = {
  name: 'Selectable — sample-1.xlsx',
  render(args) { return buildSelectableUI(args, `${import.meta.env.BASE_URL}xlsx/demo/sample-1.xlsx`); },
};

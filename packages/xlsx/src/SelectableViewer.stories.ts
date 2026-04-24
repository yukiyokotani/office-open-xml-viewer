import type { Meta, StoryObj } from '@storybook/html';
import { XlsxViewer } from './viewer';
import type { CellRange } from './viewer';

type Args = {
  scale: number;
};

const meta: Meta<Args> = {
  title: 'XlsxViewer/SelectableViewer',
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

function buildSelectableUI(args: Args, autoLoadUrl?: string): HTMLElement {
  const root = document.createElement('div');
  root.style.cssText =
    'width:100%;height:100vh;display:flex;flex-direction:column;overflow:hidden;font-family:sans-serif;box-sizing:border-box;';

  // Toolbar
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
      selectionLabel.textContent = 'Click or drag to select cells. Ctrl+C to copy.';
      copyBtn.disabled = true;
      return;
    }
    const r1 = Math.min(sel.anchor.row, sel.active.row);
    const r2 = Math.max(sel.anchor.row, sel.active.row);
    const c1 = Math.min(sel.anchor.col, sel.active.col);
    const c2 = Math.max(sel.anchor.col, sel.active.col);
    const colLabel = (n: number) => {
      let s = '';
      while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
      return s;
    };
    const topLeft = `${colLabel(c1)}${r1}`;
    const bottomRight = `${colLabel(c2)}${r2}`;
    selectionLabel.textContent =
      r1 === r2 && c1 === c2
        ? `Selected: ${topLeft}`
        : `Selected: ${topLeft}:${bottomRight} (${r2 - r1 + 1}×${c2 - c1 + 1})`;
    copyBtn.disabled = false;
  }

  function createViewer(): XlsxViewer {
    viewerContainer.innerHTML = '';
    const viewer = new XlsxViewer(viewerContainer, {
      cellScale: args.scale,
      onReady: (names) => { selectionLabel.textContent = `Loaded — ${names.length} sheet(s). Click a cell to select.`; },
      onSheetChange: (_idx, name) => { selectionLabel.textContent = `Sheet: ${name} — Click a cell to select.`; },
      onError: (err) => { selectionLabel.textContent = `Error: ${err.message}`; },
      onSelectionChange,
    });
    return viewer;
  }

  copyBtn.addEventListener('click', () => {
    if (!currentViewer) return;
    const sel = currentViewer.selection;
    if (!sel) return;
    // Trigger the keyboard copy path by dispatching a synthetic Ctrl+C
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
    const buf = await file.arrayBuffer();
    await currentViewer.load(buf);
  });

  if (autoLoadUrl) {
    selectionLabel.textContent = 'Loading…';
    currentViewer = createViewer();
    fetch(autoLoadUrl)
      .then((r) => { if (!r.ok) throw new Error(`HTTP ${r.status}`); return r.arrayBuffer(); })
      .then((buf) => currentViewer!.load(buf))
      .catch((err) => { selectionLabel.textContent = `Failed: ${err.message}`; });
  } else {
    currentViewer = createViewer();
  }

  return root;
}

export const FileUpload: Story = {
  name: 'Cell Selection — file upload',
  render(args) {
    return buildSelectableUI(args);
  },
};

export const Sample1: Story = {
  name: 'Cell Selection — sample-1.xlsx',
  render(args) {
    return buildSelectableUI(args, '/xlsx/sample-1.xlsx');
  },
};

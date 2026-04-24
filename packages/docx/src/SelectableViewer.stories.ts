import type { Meta, StoryObj } from '@storybook/html';
import { DocxViewer } from './viewer';

type Args = {
  width: number;
};

const meta: Meta<Args> = {
  title: 'DocxViewer/SelectableViewer',
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

function buildSelectableUI(args: Args, autoLoadUrl?: string): HTMLElement {
  const root = document.createElement('div');
  root.style.cssText = 'font-family:sans-serif;padding:16px;';

  // Toolbar
  const toolbar = document.createElement('div');
  toolbar.style.cssText =
    'display:flex;gap:10px;align-items:center;margin-bottom:10px;flex-wrap:wrap;';

  const fileInput = document.createElement('input');
  fileInput.type = 'file';
  fileInput.accept = '.docx';
  fileInput.style.fontSize = '12px';

  const prevBtn = document.createElement('button');
  prevBtn.textContent = '← Prev';
  prevBtn.disabled = true;

  const nextBtn = document.createElement('button');
  nextBtn.textContent = 'Next →';
  nextBtn.disabled = true;

  const pageInfo = document.createElement('span');
  pageInfo.style.fontSize = '14px';

  const hint = document.createElement('span');
  hint.style.cssText = 'color:#555;font-size:13px;flex:1;';
  hint.textContent = 'Drag over text to select. Ctrl+C to copy.';

  toolbar.append(fileInput, prevBtn, nextBtn, pageInfo, hint);
  root.appendChild(toolbar);

  const container = document.createElement('div');
  container.style.cssText =
    `width:${args.width}px;max-width:100%;border:1px solid #ccc;background:#f0f0f0;`;
  root.appendChild(container);

  const canvas = document.createElement('canvas');
  container.appendChild(canvas);

  let viewer: DocxViewer | null = null;

  function createViewer(): DocxViewer {
    return new DocxViewer(canvas, {
      width: args.width,
      enableTextSelection: true,
    });
  }

  function updateNav(v: DocxViewer): void {
    pageInfo.textContent = `Page ${v.currentPage + 1} / ${v.pageCount}`;
    prevBtn.disabled = v.currentPage === 0;
    nextBtn.disabled = v.currentPage === v.pageCount - 1;
  }

  prevBtn.addEventListener('click', () => {
    viewer?.prevPage();
    if (viewer) updateNav(viewer);
  });
  nextBtn.addEventListener('click', () => {
    viewer?.nextPage();
    if (viewer) updateNav(viewer);
  });

  fileInput.addEventListener('change', async () => {
    const file = fileInput.files?.[0];
    if (!file) return;
    hint.textContent = 'Parsing…';
    viewer = createViewer();
    const buf = await file.arrayBuffer();
    await viewer.load(buf);
    updateNav(viewer);
    hint.textContent = 'Drag over text to select. Ctrl+C to copy.';
  });

  if (autoLoadUrl) {
    hint.textContent = 'Loading…';
    viewer = createViewer();
    fetch(autoLoadUrl)
      .then((r) => { if (!r.ok) throw new Error(`HTTP ${r.status}`); return r.arrayBuffer(); })
      .then(async (buf) => {
        await viewer!.load(buf);
        updateNav(viewer!);
        hint.textContent = 'Drag over text to select. Ctrl+C to copy.';
      })
      .catch((err) => { hint.textContent = `Failed: ${err.message}`; });
  } else {
    viewer = createViewer();
  }

  return root;
}

export const FileUpload: Story = {
  name: 'Text Selection — file upload',
  render(args) {
    return buildSelectableUI(args);
  },
};

export const Sample1: Story = {
  name: 'Text Selection — sample-1.docx',
  render(args) {
    return buildSelectableUI(args, '/docx/sample-1.docx');
  },
};

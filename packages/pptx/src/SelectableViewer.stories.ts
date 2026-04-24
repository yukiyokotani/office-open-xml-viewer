import type { Meta, StoryObj } from '@storybook/html';
import { PptxViewer } from './viewer';

type Args = {
  width: number;
};

const meta: Meta<Args> = {
  title: 'PptxViewer/SelectableViewer',
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

function buildSelectableUI(args: Args, autoLoadUrl?: string): HTMLElement {
  const root = document.createElement('div');
  root.style.cssText = 'font-family:sans-serif;padding:16px;';

  // Toolbar
  const toolbar = document.createElement('div');
  toolbar.style.cssText =
    'display:flex;gap:10px;align-items:center;margin-bottom:10px;flex-wrap:wrap;';

  const fileInput = document.createElement('input');
  fileInput.type = 'file';
  fileInput.accept = '.pptx';
  fileInput.style.fontSize = '12px';

  const prevBtn = document.createElement('button');
  prevBtn.textContent = '← Prev';
  prevBtn.disabled = true;

  const nextBtn = document.createElement('button');
  nextBtn.textContent = 'Next →';
  nextBtn.disabled = true;

  const slideInfo = document.createElement('span');
  slideInfo.style.fontSize = '14px';

  const hint = document.createElement('span');
  hint.style.cssText = 'color:#555;font-size:13px;flex:1;';
  hint.textContent = 'Drag over text to select. Ctrl+C to copy.';

  toolbar.append(fileInput, prevBtn, nextBtn, slideInfo, hint);
  root.appendChild(toolbar);

  const container = document.createElement('div');
  container.style.cssText =
    `position:relative;width:${args.width}px;max-width:100%;border:1px solid #ccc;background:#f0f0f0;`;
  root.appendChild(container);

  let viewer: PptxViewer | null = null;

  function createViewer(): PptxViewer {
    container.innerHTML = '';
    return new PptxViewer(container, {
      width: args.width,
      enableTextSelection: true,
      onSlideChange(index, total) {
        slideInfo.textContent = `Slide ${index + 1} / ${total}`;
        prevBtn.disabled = index === 0;
        nextBtn.disabled = index === total - 1;
      },
      onError(err) { hint.textContent = `Error: ${err.message}`; },
    });
  }

  prevBtn.addEventListener('click', () => viewer?.prevSlide());
  nextBtn.addEventListener('click', () => viewer?.nextSlide());

  fileInput.addEventListener('change', async () => {
    const file = fileInput.files?.[0];
    if (!file) return;
    hint.textContent = 'Parsing…';
    viewer?.destroy();
    viewer = createViewer();
    const buf = await file.arrayBuffer();
    await viewer.load(buf);
    hint.textContent = 'Drag over text to select. Ctrl+C to copy.';
  });

  if (autoLoadUrl) {
    hint.textContent = 'Loading…';
    viewer = createViewer();
    fetch(autoLoadUrl)
      .then((r) => { if (!r.ok) throw new Error(`HTTP ${r.status}`); return r.arrayBuffer(); })
      .then((buf) => viewer!.load(buf))
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
  name: 'Text Selection — sample-1.pptx',
  render(args) {
    return buildSelectableUI(args, '/pptx/sample-1.pptx');
  },
};

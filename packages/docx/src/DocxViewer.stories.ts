import type { Meta, StoryObj } from '@storybook/html';
import { DocxDocument } from './document';

type Args = {
  file: string;
  width: number;
};

const meta: Meta<Args> = {
  title: 'DocxViewer',
  argTypes: {
    file: {
      control: { type: 'select' },
      options: ['sample-1.docx', 'sample-2.docx'],
      description: 'DOCX file to display',
    },
    width: {
      control: { type: 'range', min: 400, max: 1200, step: 40 },
      description: 'Canvas render width (px)',
    },
  },
  args: { file: 'sample-1.docx', width: 700 },
};
export default meta;
type Story = StoryObj<Args>;

function buildViewer(args: Args): HTMLElement {
  const root = document.createElement('div');
  root.style.cssText = 'font-family: sans-serif; padding: 16px;';

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
  container.style.cssText = `
    width: ${args.width}px;
    max-width: 100%;
    border: 1px solid #ccc;
    background: #f0f0f0;
    min-height: 120px;
  `;
  root.appendChild(container);

  const canvas = document.createElement('canvas');
  container.appendChild(canvas);

  let doc: DocxDocument | null = null;
  let currentPage = 0;

  const updateNav = () => {
    const total = doc?.pageCount ?? 0;
    pageInfo.textContent = total > 0 ? `Page ${currentPage + 1} / ${total}` : '';
    prevBtn.disabled = currentPage <= 0;
    nextBtn.disabled = currentPage >= total - 1;
  };

  const render = () => {
    if (!doc) return;
    doc.renderPage(canvas, currentPage, { width: args.width, dpr: window.devicePixelRatio });
    canvas.style.maxWidth = '100%';
  };

  prevBtn.addEventListener('click', () => {
    if (currentPage > 0) { currentPage--; render(); updateNav(); }
  });
  nextBtn.addEventListener('click', () => {
    if (doc && currentPage < doc.pageCount - 1) { currentPage++; render(); updateNav(); }
  });

  status.textContent = `Loading ${args.file}…`;
  DocxDocument.load(`/${args.file}`)
    .then((d) => {
      doc = d;
      currentPage = 0;
      status.textContent = `Loaded — ${d.pageCount} page(s)`;
      render();
      updateNav();
    })
    .catch((e: Error) => {
      status.textContent = `Error: ${e.message}`;
      status.style.color = 'red';
    });

  return root;
}

export const Sample1: Story = {
  args: { file: 'sample-1.docx', width: 700 },
  render: (args) => buildViewer(args),
};

export const Sample2: Story = {
  args: { file: 'sample-2.docx', width: 700 },
  render: (args) => buildViewer(args),
};

import type { Meta, StoryObj } from '@storybook/html';
import { DocxDocument } from './document';
import init, { parse_docx } from './wasm/docx_parser.js';

type Args = {
  width: number;
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
    `width:${args.width}px;max-width:100%;border:1px solid #ccc;background:#f0f0f0;min-height:120px;`;
  root.appendChild(container);

  const canvas = document.createElement('canvas');
  container.appendChild(canvas);

  const ctx: { doc: DocxDocument | null; page: number } = { doc: null, page: 0 };

  const updateNav = () => {
    const total = ctx.doc?.pageCount ?? 0;
    pageInfo.textContent = total > 0 ? `Page ${ctx.page + 1} / ${total}` : '';
    prevBtn.disabled = ctx.page <= 0;
    nextBtn.disabled = ctx.page >= total - 1;
  };

  const render = async () => {
    if (!ctx.doc) return;
    await ctx.doc.renderPage(canvas, ctx.page, { width: args.width, dpr: window.devicePixelRatio });
    canvas.style.maxWidth = '100%';
  };

  prevBtn.addEventListener('click', async () => {
    if (ctx.page > 0) { ctx.page--; await render(); updateNav(); }
  });
  nextBtn.addEventListener('click', async () => {
    if (ctx.doc && ctx.page < ctx.doc.pageCount - 1) { ctx.page++; await render(); updateNav(); }
  });

  if (autoLoadUrl) {
    status.textContent = `Loading ${autoLoadUrl}…`;
    DocxDocument.load(autoLoadUrl)
      .then(async (d) => {
        ctx.doc = d;
        ctx.page = 0;
        status.textContent = `Loaded — ${d.pageCount} page(s)`;
        await render();
        updateNav();
      })
      .catch((e: Error) => {
        status.textContent = `Error: ${e.message}`;
        status.style.color = 'red';
      });
  }

  return { root, doc: ctx.doc };
}

// ---------------------------------------------------------------------------
// Debug: raw JSON from WASM parser
// ---------------------------------------------------------------------------
export const DebugJson: Story = {
  name: 'Debug – raw parse JSON',
  args: { width: 700 },
  render() {
    const root = document.createElement('div');
    root.style.cssText = 'font-family:sans-serif;padding:16px;';

    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.docx';

    const pre = document.createElement('pre');
    pre.style.cssText =
      'font-size:11px;line-height:1.4;max-height:600px;overflow:auto;' +
      'background:#1e1e1e;color:#d4d4d4;padding:12px;border-radius:4px;';
    pre.textContent = 'Load a .docx to see the parsed JSON here.';

    root.append(fileInput, pre);

    let wasmReady = false;
    init().then(() => { wasmReady = true; });

    fileInput.addEventListener('change', async () => {
      const file = fileInput.files?.[0];
      if (!file || !wasmReady) return;
      try {
        const buf = await file.arrayBuffer();
        const json = parse_docx(new Uint8Array(buf));
        const parsed = JSON.parse(json);
        // Strip out base64 image data to keep the output readable
        const sanitized = JSON.parse(JSON.stringify(parsed, (key, val) =>
          key === 'dataUrl' && typeof val === 'string' && val.startsWith('data:')
            ? `[base64 ${(val.length / 1.33 / 1024).toFixed(0)} KB]`
            : val,
        ));
        pre.textContent = JSON.stringify(sanitized, null, 2);
        console.log('[docx debug] full JSON:', parsed);
      } catch (err) {
        pre.textContent = `Error: ${err instanceof Error ? err.message : String(err)}`;
      }
    });

    return root;
  },
};

// ---------------------------------------------------------------------------
// File-upload viewer
// ---------------------------------------------------------------------------
export const FileUpload: Story = {
  name: 'Load from file',
  args: { width: 700 },
  render(args) {
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

    let docRef: DocxDocument | null = null;
    let currentPage = 0;
    let canvas: HTMLCanvasElement | null = null;

    const updateNav = () => {
      const total = docRef?.pageCount ?? 0;
      pageInfo.textContent = total > 0 ? `Page ${currentPage + 1} / ${total}` : '';
      prevBtn.disabled = currentPage <= 0;
      nextBtn.disabled = currentPage >= total - 1;
    };

    const render = async () => {
      if (!docRef || !canvas) return;
      await docRef.renderPage(canvas, currentPage, { width: args.width, dpr: window.devicePixelRatio });
      canvas.style.maxWidth = '100%';
    };

    async function loadBuffer(name: string, buffer: ArrayBuffer) {
      status.textContent = `Parsing ${name}…`;
      container.innerHTML = '';
      canvas = document.createElement('canvas');
      container.appendChild(canvas);
      try {
        docRef = await DocxDocument.load(buffer);
        currentPage = 0;
        status.textContent = `Loaded ${name} — ${docRef.pageCount} page(s)`;
        await render();
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

    prevBtn.addEventListener('click', async () => {
      if (currentPage > 0) { currentPage--; await render(); updateNav(); }
    });
    nextBtn.addEventListener('click', async () => {
      if (docRef && currentPage < docRef.pageCount - 1) { currentPage++; await render(); updateNav(); }
    });

    return root;
  },
};

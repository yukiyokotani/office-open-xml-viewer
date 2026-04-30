import type { Meta, StoryObj } from '@storybook/html';
import { DocxDocument } from './document';
import { DocxViewer } from './viewer';
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

  const viewer = new DocxViewer(canvas, {
    width: args.width,
    dpr: window.devicePixelRatio,
    enableTextSelection: true,
    useGoogleFonts: true,
  });

  const updateNav = () => {
    const total = viewer.pageCount;
    pageInfo.textContent = total > 0 ? `Page ${viewer.currentPage + 1} / ${total}` : '';
    prevBtn.disabled = viewer.currentPage <= 0;
    nextBtn.disabled = viewer.currentPage >= total - 1;
  };

  prevBtn.addEventListener('click', () => { viewer.prevPage(); updateNav(); });
  nextBtn.addEventListener('click', () => { viewer.nextPage(); updateNav(); });

  if (autoLoadUrl) {
    status.textContent = `Loading ${autoLoadUrl}…`;
    viewer.load(autoLoadUrl)
      .then(() => {
        status.textContent = `Loaded — ${viewer.pageCount} page(s)`;
        updateNav();
      })
      .catch((e: Error) => {
        status.textContent = `Error: ${e.message}`;
        status.style.color = 'red';
      });
  }

  return { root, doc: null };
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

    let viewer: DocxViewer | null = null;

    const updateNav = () => {
      const total = viewer?.pageCount ?? 0;
      pageInfo.textContent = total > 0 ? `Page ${(viewer?.currentPage ?? 0) + 1} / ${total}` : '';
      prevBtn.disabled = (viewer?.currentPage ?? 0) <= 0;
      nextBtn.disabled = (viewer?.currentPage ?? 0) >= total - 1;
    };

    async function loadBuffer(name: string, buffer: ArrayBuffer) {
      status.textContent = `Parsing ${name}…`;
      container.innerHTML = '';
      const canvas = document.createElement('canvas');
      container.appendChild(canvas);
      viewer = new DocxViewer(canvas, {
        width: args.width,
        dpr: window.devicePixelRatio,
        enableTextSelection: true,
        useGoogleFonts: true,
      });
      try {
        await viewer.load(buffer);
        status.textContent = `Loaded ${name} — ${viewer.pageCount} page(s)`;
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

    prevBtn.addEventListener('click', () => { viewer?.prevPage(); updateNav(); });
    nextBtn.addEventListener('click', () => { viewer?.nextPage(); updateNav(); });

    return root;
  },
};

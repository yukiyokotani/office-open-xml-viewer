import type { Meta, StoryObj } from '@storybook/html';
import { PptxViewer } from './viewer';
import init, { parse_pptx } from './wasm/pptx_parser.js';

type Args = {
  width: number;
};

const meta: Meta<Args> = {
  title: 'PptxViewer',
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

// ---------------------------------------------------------------------------
// Helper: build nav bar + viewer
// ---------------------------------------------------------------------------
function buildViewerUI(
  args: Args,
  autoLoadUrl?: string
): { root: HTMLElement; viewer: PptxViewer } {
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

  const slideInfo = document.createElement('span');
  slideInfo.style.fontSize = '14px';

  const status = document.createElement('div');
  status.style.cssText = 'color:#666;font-size:13px;margin-bottom:8px;min-height:18px;';

  toolbar.append(prevBtn, nextBtn, slideInfo);
  root.append(toolbar, status);

  const container = document.createElement('div');
  container.style.cssText =
    `width:${args.width}px;max-width:100%;border:1px solid #ccc;background:#f0f0f0;min-height:120px;`;
  root.appendChild(container);

  const viewer = new PptxViewer(container, {
    width: args.width,
    onReady: () => { status.textContent = 'Ready'; },
    onSlideChange: (idx, total) => {
      slideInfo.textContent = `Slide ${idx + 1} / ${total}`;
      prevBtn.disabled = idx === 0;
      nextBtn.disabled = idx === total - 1;
    },
    onError: (err) => { status.textContent = `Error: ${err.message}`; },
  });

  prevBtn.addEventListener('click', () => viewer.prevSlide());
  nextBtn.addEventListener('click', () => viewer.nextSlide());

  if (autoLoadUrl) {
    fetch(autoLoadUrl)
      .then((r) => {
        if (!r.ok) throw new Error(`HTTP ${r.status} from ${autoLoadUrl}`);
        return r.arrayBuffer();
      })
      .then((buf) => {
        status.textContent = 'Parsing…';
        return viewer.load(buf);
      })
      .then(() => { status.textContent = 'Loaded'; })
      .catch((err) => {
        status.textContent = `Failed: ${err.message}`;
      });
  }

  return { root, viewer };
}

// ---------------------------------------------------------------------------
// Auto-load from public/sample.pptx  (place a file at public/sample.pptx)
// ---------------------------------------------------------------------------
export const SampleFile: Story = {
  name: 'Sample file (public/sample.pptx)',
  args: { width: 960 },
  render(args) {
    const { root } = buildViewerUI(args, '/sample.pptx');
    return root;
  },
};

// ---------------------------------------------------------------------------
// Debug: raw JSON from WASM parser (helps diagnose blank output)
// ---------------------------------------------------------------------------
export const DebugJson: Story = {
  name: 'Debug – raw parse JSON',
  args: { width: 960 },
  render(_args) {
    const root = document.createElement('div');
    root.style.cssText = 'font-family:sans-serif;padding:16px;';

    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.pptx';

    const pre = document.createElement('pre');
    pre.style.cssText =
      'font-size:11px;line-height:1.4;max-height:600px;overflow:auto;' +
      'background:#1e1e1e;color:#d4d4d4;padding:12px;border-radius:4px;';
    pre.textContent = 'Load a .pptx to see the parsed JSON here.';

    root.append(fileInput, pre);

    let wasmReady = false;
    init().then(() => { wasmReady = true; });

    fileInput.addEventListener('change', async () => {
      const file = fileInput.files?.[0];
      if (!file || !wasmReady) return;
      try {
        const buf = await file.arrayBuffer();
        const json = parse_pptx(new Uint8Array(buf));
        const parsed = JSON.parse(json);
        // Print summary: slide count, element count per slide
        const summary = {
          slideWidth: parsed.slideWidth,
          slideHeight: parsed.slideHeight,
          slideCount: parsed.slides.length,
          slides: (parsed.slides as Array<{ elements: unknown[] }>).map((s, i) => ({
            slideIndex: i,
            elementCount: s.elements.length,
            elements: s.elements,
          })),
        };
        pre.textContent = JSON.stringify(summary, null, 2);
        console.log('[pptx debug] full JSON:', parsed);
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
  args: { width: 960 },
  render(args) {
    const root = document.createElement('div');
    root.style.cssText = 'font-family:sans-serif;padding:16px;';

    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.pptx';
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

    const slideInfo = document.createElement('span');
    slideInfo.style.fontSize = '14px';

    toolbar.append(prevBtn, nextBtn, slideInfo);

    const container = document.createElement('div');
    container.style.cssText =
      `width:${args.width}px;max-width:100%;border:1px solid #ccc;background:#f0f0f0;` +
      `display:flex;align-items:center;justify-content:center;min-height:200px;`;
    const hint = document.createElement('span');
    hint.textContent = 'Drop a .pptx here or use the chooser above';
    hint.style.color = '#aaa';
    container.appendChild(hint);

    root.append(fileInput, status, toolbar, container);

    let viewer: PptxViewer | null = null;

    async function loadBuffer(name: string, buffer: ArrayBuffer) {
      status.textContent = `Parsing ${name}…`;
      viewer?.destroy();
      container.innerHTML = '';
      viewer = new PptxViewer(container, {
        width: args.width,
        onSlideChange: (idx, total) => {
          slideInfo.textContent = `Slide ${idx + 1} / ${total}`;
          prevBtn.disabled = idx === 0;
          nextBtn.disabled = idx === total - 1;
        },
        onError: (err) => { status.textContent = `Error: ${err.message}`; },
      });
      try {
        await viewer.load(buffer);
        status.textContent = `Loaded ${name}`;
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
      if (file?.name.endsWith('.pptx')) {
        loadBuffer(file.name, await file.arrayBuffer());
      }
    });

    prevBtn.addEventListener('click', () => viewer?.prevSlide());
    nextBtn.addEventListener('click', () => viewer?.nextSlide());

    return root;
  },
};

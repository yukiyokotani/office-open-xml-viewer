import type { Meta, StoryObj } from '@storybook/html';
// Direct workspace-relative import (instead of `@silurus/ooxml-markdown`)
// so this story works without depending on whether pnpm has been re-run
// after the workspace dep was added. Vite resolves the path directly.
import { docxToMarkdown, initDocxFromBytes } from '../../markdown/src/index';
import docxWasmUrl from './wasm/docx_parser_bg.wasm?url';

type Args = Record<string, never>;

const meta: Meta<Args> = {
  title: 'DocxViewer/Markdown',
};
export default meta;

type Story = StoryObj<Args>;

let initOnce: Promise<void> | null = null;
function ensureInit(): Promise<void> {
  if (!initOnce) {
    initOnce = fetch(docxWasmUrl)
      .then((r) => r.arrayBuffer())
      .then((buf) => initDocxFromBytes(new Uint8Array(buf)));
  }
  return initOnce;
}

const DEMO_URL = `${import.meta.env.BASE_URL}docx/demo/sample-1.docx`;

export const Markdown: Story = {
  name: 'DOCX → Markdown',
  render() {
    const root = document.createElement('div');
    root.style.cssText = 'font-family:sans-serif;padding:16px;';

    const toolbar = document.createElement('div');
    toolbar.style.cssText = 'display:flex;gap:8px;align-items:center;margin-bottom:10px;flex-wrap:wrap;';
    const demoBtn = document.createElement('button');
    demoBtn.textContent = 'Load demo';
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.docx';
    const status = document.createElement('span');
    status.style.cssText = 'font-size:13px;color:#444;';
    toolbar.append(demoBtn, fileInput, status);
    root.append(toolbar);

    const stats = document.createElement('div');
    stats.style.cssText = 'font-size:12px;color:#666;margin-bottom:8px;min-height:18px;';
    root.append(stats);

    const pre = document.createElement('pre');
    pre.style.cssText =
      'font-size:12px;line-height:1.4;max-height:600px;overflow:auto;' +
      'background:#1e1e1e;color:#d4d4d4;padding:12px;border-radius:4px;white-space:pre-wrap;word-break:break-word;';
    pre.textContent = `Pick a .docx file or click "Load demo".`;
    root.append(pre);

    const run = async (buf: ArrayBuffer, label: string) => {
      status.textContent = 'Parsing…';
      try {
        await ensureInit();
        const t0 = performance.now();
        const md = docxToMarkdown(buf);
        const elapsed = performance.now() - t0;
        const inKB = (buf.byteLength / 1024).toFixed(1);
        const outKB = (new TextEncoder().encode(md).byteLength / 1024).toFixed(1);
        stats.textContent = `${label} — ${inKB} KB → ${outKB} KB markdown (${(buf.byteLength / md.length).toFixed(1)}× compression, ${elapsed.toFixed(0)} ms)`;
        pre.textContent = md;
        status.textContent = 'OK';
      } catch (err) {
        status.textContent = `Error: ${(err as Error).message}`;
      }
    };

    demoBtn.addEventListener('click', async () => {
      const r = await fetch(DEMO_URL);
      const buf = await r.arrayBuffer();
      await run(buf, 'sample-1.docx');
    });
    fileInput.addEventListener('change', async () => {
      const f = fileInput.files?.[0];
      if (!f) return;
      const buf = await f.arrayBuffer();
      await run(buf, f.name);
    });

    return root;
  },
};

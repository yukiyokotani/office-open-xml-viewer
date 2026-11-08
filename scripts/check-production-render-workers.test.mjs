import assert from 'node:assert/strict';
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { checkProductionRenderWorkers } from './check-production-render-workers.mjs';

test('production workers reject an extra emitted source worker even when no host references it', () => {
  const dist = mkdtempSync(join(tmpdir(), 'ooxml-render-workers-'));
  try {
    mkdirSync(join(dist, 'assets'));
    for (const format of ['docx', 'pptx', 'xlsx']) {
      writeFileSync(join(dist, `render-worker-host-${format}.js`),
        `export const url = new URL('./assets/render-worker-${format}.js', import.meta.url).href;`);
      writeFileSync(join(dist, 'assets', `render-worker-${format}.js`), 'self.onmessage = () => {};');
    }
    assert.match(checkProductionRenderWorkers(dist), /3 self-contained module assets/);
    writeFileSync(join(dist, 'assets', 'render-worker-source-duplicate.js'), 'self.onmessage = () => {};');
    assert.throws(() => checkProductionRenderWorkers(dist), /exactly 3 shared render-worker assets/);
  } finally {
    rmSync(dist, { recursive: true, force: true });
  }
});

test('production workers reject a worker URL without ./ (#2134)', () => {
  const dist = mkdtempSync(join(tmpdir(), 'ooxml-render-workers-'));
  try {
    mkdirSync(join(dist, 'assets'));
    for (const format of ['docx', 'pptx', 'xlsx']) {
      writeFileSync(join(dist, `render-worker-host-${format}.js`),
        `export const url = new URL('assets/render-worker-${format}.js', import.meta.url).href;`);
      writeFileSync(join(dist, 'assets', `render-worker-${format}.js`), 'self.onmessage = () => {};');
    }
    assert.throws(() => checkProductionRenderWorkers(dist), /dot-relative module worker/);
  } finally {
    rmSync(dist, { recursive: true, force: true });
  }
});

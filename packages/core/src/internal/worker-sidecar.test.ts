import { describe, expect, it } from 'vitest';
import { importWorkerSidecar } from './worker-sidecar.js';

describe('importWorkerSidecar', () => {
  it('imports an absolute module URL', async () => {
    const namespace = await importWorkerSidecar(new URL('./bundled-asset-url.ts', import.meta.url).href) as Record<string, unknown>;
    expect(typeof namespace.resolveBundledAssetUrl).toBe('function');
  });

  it('rejects bare specifiers, relative paths and other protocols before importing', async () => {
    for (const url of ['@silurus/ooxml-legacy-office', '@silurus/ooxml-legacy-office/ppt', './x.mjs', '/x.mjs', '', ' https://a.test/x.mjs',
      'data:text/javascript,export default 1', 'javascript:alert(1)', 'node:fs']) {
      await expect(importWorkerSidecar(url), url).rejects.toThrow(TypeError);
    }
  });
});

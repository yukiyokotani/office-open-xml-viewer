import { describe, expect, it } from 'vitest';
import { loadWorkerRenderers } from '@silurus/ooxml-core/worker';
import type { RenderWorkerRequest } from './worker-protocol.js';

describe('PPTX worker optional renderer wire', () => {
  it('carries and reconstructs the shared ChartEx module without cloning functions', async () => {
    const request = {
      kind: 'parse',
      id: 1,
      buffer: new ArrayBuffer(0),
      resourcePolicy: {} as never,
      renderers: {
        chartEx: {
          protocol: 'ooxml-worker-renderer-module/v1',
          builtin: 'chartEx',
        },
      },
    } satisfies Extract<RenderWorkerRequest, { kind: 'parse' }>;

    const cloned = structuredClone(request);
    expect(cloned.renderers?.chartEx).toEqual({
      protocol: 'ooxml-worker-renderer-module/v1',
      builtin: 'chartEx',
    });
    expect(typeof (await loadWorkerRenderers(cloned.renderers)).chartEx?.render).toBe('function');
  });
});

import { describe, expect, it } from 'vitest';
import {
  registerBuiltinWorkerRenderer,
  workerRendererDescriptors,
} from './renderer-module-contract.js';
import {
  loadWorkerRenderer,
  loadWorkerRenderers,
} from './renderer-module.js';

describe('worker renderer module descriptors', () => {
  it('keeps worker transport metadata outside the public renderer object', async () => {
    const threeD = registerBuiltinWorkerRenderer({ render: () => true }, 'threeD');
    const descriptor = workerRendererDescriptors({ threeD })?.threeD;
    if (!descriptor) throw new Error('registered renderer descriptor missing');

    expect(threeD).toEqual({ render: expect.any(Function) });
    expect(structuredClone(descriptor)).toEqual({
      protocol: 'ooxml-worker-renderer-module/v1',
      builtin: 'threeD',
    });
    const renderer = await loadWorkerRenderer<{ render: unknown }>(descriptor);
    expect(typeof renderer.render).toBe('function');
  });

  it('rejects incompatible protocols before importing code', async () => {
    const descriptor = {
      protocol: 'ooxml-worker-renderer-module/v2',
      builtin: 'threeD',
    } as never;

    await expect(loadWorkerRenderer(descriptor)).rejects.toThrow(
      'Unsupported worker renderer protocol',
    );
  });

  it('projects only registered first-party renderers into a cloneable wire set', () => {
    const threeD = registerBuiltinWorkerRenderer({ render: () => true }, 'threeD');
    const mainOnlyMath = {
      loadMathJax: async () => undefined,
      mathMLToSvg: async () => ({ svg: '', widthEm: 0, ascentEm: 0, descentEm: 0 }),
    };
    const descriptors = workerRendererDescriptors({
      math: mainOnlyMath,
      threeD,
    });

    expect(structuredClone(descriptors)).toEqual({
      threeD: {
        protocol: 'ooxml-worker-renderer-module/v1',
        builtin: 'threeD',
      },
    });
  });

  it('reconstructs the typed renderer set inside the worker realm', async () => {
    const sources = {
      math: registerBuiltinWorkerRenderer({}, 'math', {
        engineAssetUrl: 'https://assets.example.test/mathjax.js',
      }),
      threeD: registerBuiltinWorkerRenderer({}, 'threeD'),
      regionMap: registerBuiltinWorkerRenderer({}, 'regionMap'),
      chartEx: registerBuiltinWorkerRenderer({}, 'chartEx'),
    };

    const loaded = await loadWorkerRenderers(workerRendererDescriptors(sources));

    expect(typeof loaded.math?.mathMLToSvg).toBe('function');
    expect(typeof loaded.threeD?.render).toBe('function');
    expect(typeof loaded.regionMap?.render).toBe('function');
    expect(typeof loaded.chartEx?.render).toBe('function');
  });

  it('carries the consumer-resolved math asset URL without exposing it on the renderer', () => {
    const math = registerBuiltinWorkerRenderer({ loadMathJax: () => undefined }, 'math', {
      engineAssetUrl: 'https://cdn.example.test/mathjax-stix2.js',
    });

    expect(math).toEqual({ loadMathJax: expect.any(Function) });
    expect(structuredClone(workerRendererDescriptors({ math })?.math)).toEqual({
      protocol: 'ooxml-worker-renderer-module/v1',
      builtin: 'math',
      engineAssetUrl: 'https://cdn.example.test/mathjax-stix2.js',
    });
  });
});

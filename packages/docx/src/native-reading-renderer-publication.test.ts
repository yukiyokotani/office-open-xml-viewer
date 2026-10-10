import { afterEach, describe, expect, it } from 'vitest';
import { renderLayoutSourceToCanvas } from './renderer.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import { createLayoutServices } from './layout-runtime.js';
import { installStubCanvas, syntheticDocxModel } from './testing/synthetic-document.js';
const originalCanvas = Object.getOwnPropertyDescriptor(globalThis, 'OffscreenCanvas');
afterEach(() => {
  if (originalCanvas) Object.defineProperty(globalThis, 'OffscreenCanvas', originalCanvas);
  else Reflect.deleteProperty(globalThis, 'OffscreenCanvas');
});
describe('main reading renderer publication normalization', () => {
  it('retains the ownership hook through real source selection and normalized paint options', async () => {
    installStubCanvas();
    const model = syntheticDocxModel('plain', { paragraphs: 1, wordsPerParagraph: 1 });
    const paragraph = model.body[0];
    if (paragraph.type !== 'paragraph' || paragraph.runs[0].type !== 'text') throw new Error('invented text absent');
    Object.assign(paragraph.runs[0], { __nativeReadingWordBreaking: { rawHres: 0, rawChHres: 1 } });
    const source = layoutSourceStore(model);
    const services = createLayoutServices(source);
    const target = new OffscreenCanvas(1, 1);
    target.width = 87; target.height = 43;
    let current = true;
    const error = new Error('reading publication revoked during normalized acquisition');
    const pending = renderLayoutSourceToCanvas(source, target, 0, {
      dpr: 1, layoutServices: services,
      assertPublicationCurrent: () => { if (!current) throw error; },
    });
    current = false;
    await expect(pending).rejects.toBe(error);
    expect([target.width, target.height]).toEqual([87, 43]);
  });
});

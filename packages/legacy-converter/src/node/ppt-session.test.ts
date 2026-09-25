// The Node PPTX presentation session over the direct PPT reader: no OOXML
// resource-usage snapshot, lazy images, lifecycle cancellation and native
// ownership. The direct path never acquires the OOXML PPTX parser WASM.
import { describe, expect, it, vi } from 'vitest';
import { testPptSource } from '../test-sources.js';
import { openPptxPresentation } from './node-facade.js';
import { anchor, buildPptFixture, concat, drawing, properties, record, shapeAtom, spContainer } from './ppt-records.js';

const { loadOoxml } = vi.hoisted(() => ({
  loadOoxml: vi.fn(() => { throw new Error('OOXML WASM must not load for direct PPT'); }),
}));
vi.mock('../../../node/src/wasm-loader.js', async (importOriginal) => ({
  ...await importOriginal<typeof import('../../../node/src/wasm-loader.js')>(),
  createLazyWasmModule: () => loadOoxml,
}));

// A minimal valid PNG (2x1) stored twice as embedded BLIPs, each drawn by one picture.
const png = Uint8Array.from(Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAYAAAD0In+KAAAADklEQVR4nGP4z8AAQg0AD3oDfnfpf5cAAAAASUVORK5CYII=', 'base64',
));
const picture = (index: number) => spContainer(shapeAtom(75, 41 + index, 0xa00), anchor(576, 576 * index, 576 * index + 576, 864), properties([[0x4104, index]]));
const blip = () => record(0x6e00, 0xf01e, concat(new Uint8Array(17), png));
const input = () => buildPptFixture(drawing(picture(1), picture(2)), undefined, undefined, { entries: [blip(), blip()] });

/** Count native close/free calls on the generated glue during one operation. */
async function nativeReleases(operation: (glue: typeof import('../wasm-direct-ppt/legacy_ppt_direct.js')) => Promise<unknown>) {
  const glue = await import('../wasm-direct-ppt/legacy_ppt_direct.js');
  const free = vi.spyOn(glue.LegacyPptPresentation.prototype, 'free');
  const close = vi.spyOn(glue.LegacyPptPresentation.prototype, 'close_presentation_session');
  try {
    await operation(glue);
    return { close: close.mock.calls.length, free: free.mock.calls.length };
  } finally { free.mockRestore(); close.mockRestore(); }
}

describe('Node direct PPT session', () => {
  it('keeps resource usage absent, serves images lazily, honours lifecycle abort and releases once', async () => {
    const releases = await nativeReleases(async () => {
      const lifecycle = new AbortController();
      const session = await openPptxPresentation(input(), { modelSources: [testPptSource()], signal: lifecycle.signal });
      // Exhausting the slide iterator closes the session, so read inside it.
      for await (const slide of session) {
        const paths = slide.elements.map(element => (element.type === 'picture' ? element.imagePath : ''));
        expect(paths).toHaveLength(2);
        expect(session.resourceUsage).toBeUndefined();
        const image = await session.getImage(paths[0], 'image/png');
        expect(image).toMatchObject({ size: png.length, type: 'image/png' });
        expect(new Uint8Array(await image.arrayBuffer())).toEqual(png);
        expect(session.resourceUsage).toBeUndefined();
        lifecycle.abort();
        // An extraction after the abort fails (a part already extracted is
        // served from the session's raw-part cache without native work).
        await expect(session.getImage(paths[1], 'image/png')).rejects.toMatchObject({ name: 'AbortError' });
        break;
      }
      await session.close();
      await session.close();
    });
    // The slide cursor closes its presentation session once, then the source
    // owner releases the archive once (its own close before free); a second
    // session close() releases nothing.
    expect(releases).toEqual({ close: 2, free: 1 });
    expect(loadOoxml).not.toHaveBeenCalled();
  });

  it('releases the native archive when the lifecycle aborts during bootstrap', async () => {
    const lifecycle = new AbortController();
    const releases = await nativeReleases(async (glue) => {
      const bootstrap = glue.LegacyPptPresentation.prototype.presentation_bootstrap;
      const spy = vi.spyOn(glue.LegacyPptPresentation.prototype, 'presentation_bootstrap').mockImplementation(function (this: unknown) {
        lifecycle.abort();
        return bootstrap.call(this as InstanceType<typeof glue.LegacyPptPresentation>);
      });
      try {
        await expect(openPptxPresentation(input(), { modelSources: [testPptSource()], signal: lifecycle.signal }))
          .rejects.toMatchObject({ name: 'AbortError' });
      } finally { spy.mockRestore(); }
    });
    expect(releases).toEqual({ close: 1, free: 1 });
  });
});

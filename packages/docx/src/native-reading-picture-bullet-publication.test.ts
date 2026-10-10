import type { ImagePaintResourceDescriptor } from './layout/types.js';
import { afterEach, describe, expect, it } from 'vitest';
import { syntheticDocxModel, installStubCanvas } from './testing/synthetic-document.js';
import { inventedReadingPictureBulletNumbering } from './testing/native-reading-picture-bullet.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { renderLayoutSourceToCanvas } from './renderer.js';
import { hasNativeReadingRequests, nativeReadingNotices, readingPictureBulletKeys } from './native-reading-notice.js';
const originalCanvas = Object.getOwnPropertyDescriptor(globalThis, 'OffscreenCanvas');
afterEach(() => {
  if (originalCanvas) Object.defineProperty(globalThis, 'OffscreenCanvas', originalCanvas);
  else Reflect.deleteProperty(globalThis, 'OffscreenCanvas');
});
function targetCanvas(width: number, height: number): OffscreenCanvas {
  const target = new OffscreenCanvas(width, height);
  target.width = width; target.height = height;
  return target;
}
function fixture(reading: boolean) {
  installStubCanvas();
  const model = syntheticDocxModel('plain', { paragraphs: 2, wordsPerParagraph: 2 });
  for (const paragraph of model.body) {
    if (paragraph.type !== 'paragraph') throw new Error('invented paragraph absent');
    const numbering = inventedReadingPictureBulletNumbering();
    const { __nativeReadingPictureBullet: _readingFacts, ...authoredNumbering } = numbering;
    paragraph.numbering = reading ? numbering : authoredNumbering;
  }
  const source = layoutSourceStore(model);
  const services = createLayoutServices(source);
  return { source, services, layout: layoutDocument(source, services) };
}
describe('stored-size picture-bullet publication', () => {
  it('discloses reading markers with distinct occurrence descriptors sharing embedded media', () => {
    const { source, layout } = fixture(true);
    expect(hasNativeReadingRequests(source)).toBe(true);
    const keys = layout.pages.flatMap(page => readingPictureBulletKeys(page));
    expect(keys).toHaveLength(2);
    expect(new Set(keys).size).toBe(2);
    const descriptors = keys.map(key => source.paintResources.resolve(key, 'picture-bullet') as ImagePaintResourceDescriptor);
    expect(descriptors.map(descriptor => descriptor.partPath)).toEqual([
      'media/bullet/7', 'media/bullet/7',
    ]);
    expect(nativeReadingNotices(layout, { contour: false, wordBreaking: false, pictureBullets: true }).map(notice => notice.code)).toEqual(['PICTURE_BULLETS_SIZED_FOR_READING']);
    expect(nativeReadingNotices({ ...layout, pages: [] }, { contour: false, wordBreaking: false, pictureBullets: true })).toEqual([]);
    expect(hasNativeReadingRequests(source)).toBe(true);
  });
  it('leaves authored DOCX picture bullets outside native reading publication', () => {
    const { source, layout } = fixture(false);
    expect(hasNativeReadingRequests(source)).toBe(false);
    expect(layout.pages.flatMap(page => readingPictureBulletKeys(page))).toEqual([]);
    expect(nativeReadingNotices(layout, { contour: false, wordBreaking: false, pictureBullets: false })).toEqual([]);
  });
  it('rejects an unavailable owned picture before resizing or clearing the caller surface', async () => {
    const { source, services } = fixture(true);
    const target = targetCanvas(87, 43);
    await expect(renderLayoutSourceToCanvas(source, target, 0, { dpr: 1, layoutServices: services }))
      .rejects.toThrow('no decoded owned image');
    expect([target.width, target.height]).toEqual([87, 43]);
  });
  it('retains the publication hook through real normalization before clearing', async () => {
    const { source, services } = fixture(false);
    const target = targetCanvas(87, 43);
    const error = new Error('publication superseded during acquisition');
    let current = true;
    const pending = renderLayoutSourceToCanvas(source, target, 0, { dpr: 1, layoutServices: services,
      assertPublicationCurrent: () => { if (!current) throw error; } });
    current = false;
    await expect(pending).rejects.toBe(error);
    expect([target.width, target.height]).toEqual([87, 43]);
  });
});

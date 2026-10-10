import { afterEach, describe, expect, it, vi } from 'vitest';
import * as sourceKeys from './layout/source-key.js';
import { isLayoutSourceStore } from './layout/layout-source-store.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { projectRenderWorkerLayoutMeta } from './render-worker-metadata.js';
import { nativeReadingRequests, hasNativeReadingRequests, nativeReadingNotices,
  retainedReadingNotices, readingPictureBulletKeys } from './native-reading-notice.js';
import { inventedReadingPictureBulletNumbering } from './testing/native-reading-picture-bullet.js';
import { installStubCanvas, syntheticDocxModel } from './testing/synthetic-document.js';
import type { DocumentLayout } from './layout/types.js';

const originalCanvas = Object.getOwnPropertyDescriptor(globalThis, 'OffscreenCanvas');
afterEach(() => {
  vi.restoreAllMocks();
  if (originalCanvas) Object.defineProperty(globalThis, 'OffscreenCanvas', originalCanvas);
  else Reflect.deleteProperty(globalThis, 'OffscreenCanvas');
});
const wordOwner = { rawHres: 0, rawChHres: 1 };
const markerOwner = { rawHres: 2, rawChHres: 120 };
function model(word = true) {
  const result = syntheticDocxModel('plain', { paragraphs: 2, wordsPerParagraph: 2 });
  const first = result.body[0], second = result.body[1];
  if (first.type !== 'paragraph' || second.type !== 'paragraph' || first.runs[0].type !== 'text')
    throw new Error('invented paragraphs absent');
  if (word) Object.assign(first.runs[0], { __nativeReadingWordBreaking: wordOwner });
  second.numbering = inventedReadingPictureBulletNumbering();
  if (word) Object.assign(second.numbering, { __nativeReadingWordBreaking: markerOwner });
  return result;
}

// Invented source-bound wire: no Office measurement or private data. The first
// paragraph requests a contour while a later paragraph owns picture numbering.
function addContourAndOptionalHyphen(result: ReturnType<typeof model>) {
  const first = result.body[0];
  if (first.type !== 'paragraph' || first.runs[0].type !== 'text') throw new Error('invented text absent');
  first.runs.push({ ...first.runs[0], text: '', __optionalHyphen: true } as typeof first.runs[number]);
  const edges = { topPt: null, topStatus: 'missing', rightPt: null, rightStatus: 'missing',
    bottomPt: null, bottomStatus: 'missing', leftPt: null, leftStatus: 'missing' };
  first.runs.push({ type: 'anchorHost', __anchorOccurrenceId: 'composed-image' } as unknown as typeof first.runs[number]);
  first.runs.push({ type: 'image', anchor: true, imagePath: 'invented-composed.png', mimeType: 'image/png',
    widthPt: 10, heightPt: 10, __anchorAcquisition: {
      occurrenceId: 'composed-image', nativeReadingRelocation: 'completeScene',
      simplePosition: { enabled: false, status: 'valid', xPt: 0, xStatus: 'valid', yPt: 0, yStatus: 'valid' },
      horizontal: { relativeFrom: 'column', relativeFromStatus: 'valid', choice: { kind: 'offset', valuePt: 0 } },
      vertical: { relativeFrom: 'paragraph', relativeFromStatus: 'valid', choice: { kind: 'offset', valuePt: 0 } },
      extent: { widthPt: 10, heightPt: 10, widthStatus: 'valid', heightStatus: 'valid' },
      parentEffectExtent: edges, anchorDistances: edges, relativeSize: { horizontal: null, vertical: null },
      wrap: { kind: 'tight', authoredKinds: ['wrapTight'], side: null, distances: edges, effectExtent: null, polygon: null },
      behavior: { behindDoc: false, behindDocStatus: 'valid', relativeHeight: 0, relativeHeightStatus: 'valid',
        locked: false, lockedStatus: 'valid', allowOverlap: true, allowOverlapStatus: 'valid',
        layoutInCell: true, layoutInCellStatus: 'valid' }, group: null,
    },
  } as unknown as typeof first.runs[number]);
}

describe('composed native reading ownership and disclosures', () => {
  it('acquires every capability across paragraphs without dropping numbering owners or authored optional hyphens', () => {
    const original = model();
    addContourAndOptionalHyphen(original);
    const before = JSON.stringify(original);
    const source = layoutSourceStore(original);
    const first = source.blocks.resolve({ story: 'body', storyInstance: 'body', path: [0] });
    const second = source.blocks.resolve({ story: 'body', storyInstance: 'body', path: [1] });
    if (first.type !== 'paragraph' || second.type !== 'paragraph' || first.runs[0].type !== 'text'
      || first.runs[1].type !== 'text') throw new Error('acquired owners absent');
    expect(nativeReadingRequests(source)).toEqual({ contour: true, wordBreaking: true, pictureBullets: true });
    expect(hasNativeReadingRequests(source)).toBe(true);
    expect(first.runs[0].nativeReadingWordBreaking).toEqual(wordOwner);
    expect(first.runs[1].optionalHyphen).toBe(true);
    expect(first.runs[0]).not.toHaveProperty('__nativeReadingWordBreaking');
    expect(first.runs[1]).not.toHaveProperty('__optionalHyphen');
    expect(second.nativeReadingNumberingWordBreaking).toEqual(markerOwner);
    expect(second.nativeReadingPictureBullet?.rawPbiFlags).toBe(0xa5fd);
    expect(second.numbering).not.toHaveProperty('__nativeReadingWordBreaking');
    expect(second.numbering).not.toHaveProperty('__nativeReadingPictureBullet');
    expect(JSON.stringify(original)).toBe(before);
  });

  it('projects mixed word and picture notices from real acquisition/layout into worker metadata', () => {
    installStubCanvas();
    const source = layoutSourceStore(model());
    const layout = layoutDocument(source, createLayoutServices(source));
    const requests = nativeReadingRequests(source);
    expect(requests).toEqual({ contour: false, wordBreaking: true, pictureBullets: true });
    expect(layout.pages.flatMap(readingPictureBulletKeys).length).toBeGreaterThan(0);
    const notices = nativeReadingNotices(layout, requests);
    expect(notices.map(notice => notice.code)).toEqual([
      'WORD_BREAKING_SIMPLIFIED_FOR_READING', 'PICTURE_BULLETS_SIZED_FOR_READING',
    ]);
    const meta = projectRenderWorkerLayoutMeta(layout, source, { comments: [], revisions: [] });
    expect(meta.nativeReadingRequested).toBe(true);
    expect(retainedReadingNotices(structuredClone(meta.readingNotices))).toEqual(notices);
  });

  it('never turns a picture-only source into word-breaking reading, including a notice-free variant', () => {
    installStubCanvas();
    const source = layoutSourceStore(model(false));
    const layout = layoutDocument(source, createLayoutServices(source));
    const requests = nativeReadingRequests(source);
    expect(requests).toEqual({ contour: false, wordBreaking: false, pictureBullets: true });
    expect(nativeReadingNotices(layout, requests).map(notice => notice.code))
      .toEqual(['PICTURE_BULLETS_SIZED_FOR_READING']);
    // A retained variant with no selected marker still owns the source lifecycle.
    // This is a disclosure-boundary control, not a Word visibility/paint oracle.
    const noticeFree = { pages: [{ layers: { body: [], roots: [], paintOrder: [] } }], diagnostics: [] } as unknown as DocumentLayout;
    expect(nativeReadingNotices(noticeFree, requests)).toEqual([]);
    expect(hasNativeReadingRequests(source)).toBe(true);
  });

  it('normalizes a complete mixed disclosure and rejects reordered or dependent-only worker subsets', () => {
    // Retained-graph projection control: no parser or page-layout claim here.
    const paragraph = { kind: 'paragraph', nativeReadingRelocations: ['complete-scene'],
      nativeReadingInactivePictureData: true, nativeReadingPictureBullet: true, textBoxes: [],
      lines: [{ placements: [{ kind: 'resource', resourceKind: 'picture-bullet', resourceKey: 'owned-marker' }] }] };
    const layout = { pages: [{ layers: { body: [paragraph], roots: [{ node: paragraph }], paintOrder: [] } }],
      diagnostics: [] } as unknown as DocumentLayout;
    const notices = nativeReadingNotices(layout, { contour: true, wordBreaking: true, pictureBullets: true });
    expect(notices.map(notice => notice.code)).toEqual([
      'DRAWINGS_RELOCATED_FOR_READING', 'INACTIVE_PICTURE_DATA_RETAINED',
      'WORD_BREAKING_SIMPLIFIED_FOR_READING', 'PICTURE_BULLETS_SIZED_FOR_READING',
    ]);
    const retained = retainedReadingNotices(structuredClone(notices));
    expect(retained).toEqual(notices);
    expect(Object.isFrozen(retained)).toBe(true);
    expect(retained.every(Object.isFrozen)).toBe(true);
    expect(() => retainedReadingNotices([notices[3], notices[2]])).toThrow('Invalid reading-layout disclosure');
    expect(() => retainedReadingNotices([notices[0], notices[0]])).toThrow('Invalid reading-layout disclosure');
    expect(() => retainedReadingNotices([notices[1], notices[2]])).toThrow('Invalid reading-layout disclosure');
  });
});


// Observe the genuine sealed repository's address resolutions. The wrapper
// delegates to the actual sourceKey function; no source, resolver or layout is
// replaced and factory-owned immutability/branding remains intact.
describe('sealed native reading requests across worker publications', () => {
  it('resolves each strict source block once across repeated progressive metadata publications', () => {
    installStubCanvas();
    const source = layoutSourceStore(syntheticDocxModel('plain', { paragraphs: 3, wordsPerParagraph: 2 }));
    const layout = layoutDocument(source, createLayoutServices(source));
    expect(isLayoutSourceStore(source)).toBe(true);
    const address = sourceKeys.sourceKey;
    const resolutions = vi.spyOn(sourceKeys, 'sourceKey');
    const first = projectRenderWorkerLayoutMeta(layout, source, { comments: [], revisions: [] }, { provisional: true });
    for (let publication = 0; publication < 4; publication++) {
      expect(projectRenderWorkerLayoutMeta(layout, source, { comments: [], revisions: [] }, { provisional: true })).toEqual(first);
      expect(hasNativeReadingRequests(source)).toBe(false);
    }
    expect(first).not.toHaveProperty('nativeReadingRequested');
    expect(first).not.toHaveProperty('readingNotices');
    const counts = new Map<string, number>();
    for (const [ref] of resolutions.mock.calls) {
      const key = address(ref); counts.set(key, (counts.get(key) ?? 0) + 1);
    }
    expect([...counts]).toEqual(source.blocks.sources.map(ref => [address(ref), 1]));
  });

  it('keeps independent sealed sources and all three acquired policies distinct from mutable public models', () => {
    installStubCanvas();
    const publicModel = syntheticDocxModel('plain', { paragraphs: 2, wordsPerParagraph: 2 });
    const strict = layoutSourceStore(publicModel);
    const strictLayout = layoutDocument(strict, createLayoutServices(strict));
    const mixedModel = model(); addContourAndOptionalHyphen(mixedModel);
    // A real anchor-host segment participates in font measurement even though
    // it has no glyph. Supply the same authored font as its paragraph text.
    const host = mixedModel.body[0].type === 'paragraph'
      ? mixedModel.body[0].runs.find(run => run.type === 'anchorHost') : undefined;
    if (!host || host.type !== 'anchorHost') throw new Error('invented anchor host absent');
    Object.assign(host, { fontSize: 10, fontFamily: 'Times New Roman' });
    const mixed = layoutSourceStore(mixedModel);
    const mixedLayout = layoutDocument(mixed, createLayoutServices(mixed));
    const pictureOnly = layoutSourceStore(model(false));
    const pictureLayout = layoutDocument(pictureOnly, createLayoutServices(pictureOnly));
    const resolutions = vi.spyOn(sourceKeys, 'sourceKey');
    projectRenderWorkerLayoutMeta(strictLayout, strict, { comments: [], revisions: [] });
    expect(resolutions.mock.calls.length).toBe(strict.blocks.sources.length);
    resolutions.mockClear();
    const firstMixed = projectRenderWorkerLayoutMeta(mixedLayout, mixed, { comments: [], revisions: [] });
    expect(nativeReadingRequests(mixed)).toEqual({ contour: true, wordBreaking: true, pictureBullets: true });
    expect(firstMixed.readingNotices?.map(notice => notice.code)).toEqual([
      'DRAWINGS_RELOCATED_FOR_READING', 'WORD_BREAKING_SIMPLIFIED_FOR_READING', 'PICTURE_BULLETS_SIZED_FOR_READING',
    ]);
    expect(projectRenderWorkerLayoutMeta(mixedLayout, mixed, { comments: [], revisions: [] })).toEqual(firstMixed);
    expect(resolutions.mock.calls.length).toBe(mixed.blocks.sources.length);
    resolutions.mockClear();
    const pictureMeta = projectRenderWorkerLayoutMeta(pictureLayout, pictureOnly, { comments: [], revisions: [] });
    expect(nativeReadingRequests(pictureOnly)).toEqual({ contour: false, wordBreaking: false, pictureBullets: true });
    expect(pictureMeta.readingNotices?.map(notice => notice.code)).toEqual(['PICTURE_BULLETS_SIZED_FOR_READING']);
    expect(resolutions.mock.calls.length).toBe(pictureOnly.blocks.sources.length);
    resolutions.mockRestore();
    const paragraph = publicModel.body[0];
    if (paragraph.type !== 'paragraph' || paragraph.runs[0].type !== 'text') throw new Error('invented owner absent');
    Object.assign(paragraph.runs[0], { __nativeReadingWordBreaking: wordOwner });
    expect(nativeReadingRequests(strict)).toEqual({ contour: false, wordBreaking: false, pictureBullets: false });
    const fresh = layoutSourceStore(structuredClone(publicModel));
    expect(nativeReadingRequests(fresh)).toEqual({ contour: false, wordBreaking: true, pictureBullets: false });
  });

  it('does not retain a partially scanned genuine source when an addressing failure aborts acquisition', () => {
    const source = layoutSourceStore(syntheticDocxModel('plain', { paragraphs: 3, wordsPerParagraph: 2 }));
    expect(isLayoutSourceStore(source)).toBe(true);
    const address = sourceKeys.sourceKey;
    let visits = 0;
    const failure = new Error('invented addressing failure');
    const resolutions = vi.spyOn(sourceKeys, 'sourceKey').mockImplementation(ref => {
      visits++;
      if (visits === 2) throw failure;
      return address(ref);
    });
    let observed: unknown;
    try { nativeReadingRequests(source); } catch (error) { observed = error; }
    expect(observed === failure).toBe(true);
    resolutions.mockImplementation(address); resolutions.mockClear();
    expect(nativeReadingRequests(source)).toEqual({ contour: false, wordBreaking: false, pictureBullets: false });
    expect(resolutions.mock.calls.length).toBe(source.blocks.sources.length);
    expect(nativeReadingRequests(source)).toEqual({ contour: false, wordBreaking: false, pictureBullets: false });
    expect(resolutions.mock.calls.length).toBe(source.blocks.sources.length);
  });
});

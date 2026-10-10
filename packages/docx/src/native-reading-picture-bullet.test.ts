import { describe, expect, it } from 'vitest';
import { acquireNativeReadingPictureBullet } from './native-reading-picture-bullet.js';
import { paragraphAcquisitionInput } from './parser-model.js';
import { inventedReadingPictureBulletNumbering } from './testing/native-reading-picture-bullet.js';
import { syntheticDocxModel } from './testing/synthetic-document.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import { hasNativeReadingRequests, retainedReadingNotices } from './native-reading-notice.js';

const source = { story: 'body' as const, storyInstance: 'body', path: [0] };
describe('native picture-bullet owned reading acquisition', () => {
  it('retains exact raw source facts in a new immutable owner', () => {
    const numbering = inventedReadingPictureBulletNumbering();
    const facts = acquireNativeReadingPictureBullet(numbering)!;
    expect(facts).toEqual(numbering.__nativeReadingPictureBullet);
    expect(facts).not.toBe(numbering.__nativeReadingPictureBullet);
    expect(Object.isFrozen(facts.clientAnchor)).toBe(true);
    Object.assign(numbering.__nativeReadingPictureBullet!, { rawPbiFlags: 3 });
    expect(facts.rawPbiFlags).toBe(0xa5fd);
  });
  it('uses the selected opaque resource identity without interpreting its namespace', () => {
    const numbering = inventedReadingPictureBulletNumbering();
    numbering.picBulletImagePath = 'embedded-marker:opaque-owner';
    Object.assign(numbering.__nativeReadingPictureBullet!, { resourceKey: numbering.picBulletImagePath });
    const facts = acquireNativeReadingPictureBullet(numbering)!;
    expect(facts).toMatchObject({ resourceKey: numbering.picBulletImagePath, picfOffset: 7 });
    expect(Object.isFrozen(facts)).toBe(true);
  });
  it.each(['missing', 'empty', 'detached'])('rejects a %s selected resource identity', variant => {
    const numbering = inventedReadingPictureBulletNumbering();
    const { resourceKey: _selected, ...withoutIdentity } = numbering.__nativeReadingPictureBullet! as
      typeof numbering.__nativeReadingPictureBullet & { resourceKey?: unknown };
    Object.assign(numbering, { __nativeReadingPictureBullet: variant === 'missing' ? withoutIdentity
      : { ...withoutIdentity, resourceKey: variant === 'empty' ? '' : 'another-embedded-marker' } });
    expect(() => acquireNativeReadingPictureBullet(numbering)).toThrow(TypeError);
  });
  it('does not label authored DOCX picture bullets or disabled source owners', () => {
    const { __nativeReadingPictureBullet: _readingFacts, ...numbering } = inventedReadingPictureBulletNumbering();
    expect(acquireNativeReadingPictureBullet(numbering)).toBeUndefined();
    expect(acquireNativeReadingPictureBullet(null)).toBeUndefined();
  });
  it.each([
    { rawPbiFlags: 2 }, { rawShapeFlags: 0x10 }, { shape: 100 }, { resourceKey: 'detached-marker' },
    { flagsOrigin: { kind: 'style', id: 1 } }, { indexOrigin: { kind: 'decoded' } },
    { pibFlags: { key: 0x0105, value: 0x40 } }, { pibFlags: { key: 0x4106, value: 0 } },
    { pibFlags: { key: 0x8106, value: 0 } }, { pibFlags: { key: 0xc106, value: 0 } }, { clientAnchor: [-1] },
    { clientAnchor: null, clientAnchorOptions: 0 }, { goalTwips: [0, 720] },
  ])('rejects a detached or ambiguous raw owner %j', patch => {
    const numbering = inventedReadingPictureBulletNumbering();
    Object.assign(numbering.__nativeReadingPictureBullet!, patch);
    expect(() => acquireNativeReadingPictureBullet(numbering)).toThrow(TypeError);
  });
  it.each([3, 4, 5, 6, 7, 8, 11, 12, 15, 16, 0x40, 0xffffffff])('rejects undefined or incompatible MSOBLIPFLAGS %j', value => {
    const numbering = inventedReadingPictureBulletNumbering();
    Object.assign(numbering.__nativeReadingPictureBullet!, { pibFlags: { key: 0x106, value } });
    expect(() => acquireNativeReadingPictureBullet(numbering)).toThrow('MSOBLIPFLAGS');
  });
  it.each([1, 2, 9, 10, 13, 14])('keeps valid named/link mode outside the bounded passive consumer %j', value => {
    const numbering = inventedReadingPictureBulletNumbering();
    Object.assign(numbering.__nativeReadingPictureBullet!, { pibFlags: { key: 0x106, value } });
    expect(() => acquireNativeReadingPictureBullet(numbering)).toThrow('no stored-size reading consumer');
  });
  it('retains the absence of an encoded flags operand instead of synthesizing Comment0', () => {
    const numbering = inventedReadingPictureBulletNumbering();
    Object.assign(numbering.__nativeReadingPictureBullet!, { pibFlags: null });
    expect(acquireNativeReadingPictureBullet(numbering)?.pibFlags).toBeNull();
  });
  it('moves private raw facts off semantic numbering without changing literals, box, transform or authored fonts', () => {
    const model = syntheticDocxModel('plain', { paragraphs: 1, wordsPerParagraph: 2 });
    const paragraph = model.body[0];
    if (paragraph.type !== 'paragraph') throw new Error('invented paragraph absent');
    const rawNumbering = inventedReadingPictureBulletNumbering();
    paragraph.numbering = rawNumbering;
    const input = paragraphAcquisitionInput(paragraph, source);
    expect(input.nativeReadingPictureBullet?.rawPbiFlags).toBe(0xa5fd);
    expect(input.numbering).not.toHaveProperty('__nativeReadingPictureBullet');
    expect(input.numbering).not.toHaveProperty('fontFacts');
    expect(input.numbering?.picBulletTransform).toEqual(paragraph.numbering.picBulletTransform);
    expect([input.numbering?.picBulletWidthPt, input.numbering?.picBulletHeightPt]).toEqual([36, 72]);
    expect(input.numberingMarkerShapeInput?.fontSizePt).toBe(24);
    expect(input.runs).toHaveLength(paragraph.runs.length);
    expect(input.runs[0]).toMatchObject(paragraph.runs[0]);
    expect(hasNativeReadingRequests(layoutSourceStore(model))).toBe(true);
    const { __nativeReadingPictureBullet: _readingFacts, ...authoredNumbering } = rawNumbering;
    const authoredModel = {
      ...model,
      body: model.body.map(element => element === paragraph
        ? { ...paragraph, numbering: authoredNumbering } : element),
    };
    expect(hasNativeReadingRequests(layoutSourceStore(authoredModel))).toBe(false);
    expect(hasNativeReadingRequests(layoutSourceStore(model))).toBe(true);
  });
  it('rejects unknown or altered public disclosure vocabulary', () => {
    expect(() => retainedReadingNotices([{ code: 'PICTURE_BULLETS_SIZED_FOR_READING', message: 'claims exact Word size' }])).toThrow();
    expect(retainedReadingNotices(undefined)).toEqual([]);
  });
});

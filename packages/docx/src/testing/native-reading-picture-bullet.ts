import type { InternalNumberingInfo } from '../parser-model.js';
import type { NativeReadingPictureBullet } from '../native-reading-picture-bullet.js';

/** Invented control. No private source, Word measurement or fitted scale. */
export function inventedReadingPictureBulletFacts(): NativeReadingPictureBullet {
  return {
    resourceKey: 'legacy-doc/bullet/7', rawPbiFlags: 0xa5fd,
    flagsOrigin: { kind: 'listLevel', instance: 0, list: 0, level: 0 },
    indexOrigin: { kind: 'piece', fc: 1026, prm: 1 },
    relativeCp: 0, picfOffset: 7, shape: 75, rawShapeFlags: 0xc0,
    goalTwips: [1440, 720], scalePerMille: [500, 2000],
    pibFlags: { key: 0x106, value: 0 }, clientAnchor: [7, 6, 5, 128], clientAnchorOptions: 0,
  };
}
export function inventedReadingPictureBulletNumbering(): InternalNumberingInfo {
  return {
    numId: 1, level: 0, format: 'bullet', text: '•', indentLeft: 48, tab: 12, suff: 'tab',
    fontFamily: 'Symbol', fontFacts: { fontFamily: 'Symbol', fontSize: 24 },
    picBulletImagePath: 'legacy-doc/bullet/7', picBulletMimeType: 'image/png',
    picBulletWidthPt: 36, picBulletHeightPt: 72,
    picBulletTransform: { srcRect: { l: 0.125, t: 0.1, r: 0.2, b: 0.25 }, rotation: 90, flipH: true, flipV: true },
    __nativeReadingPictureBullet: inventedReadingPictureBulletFacts(),
  };
}

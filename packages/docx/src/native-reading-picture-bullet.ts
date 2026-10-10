import type { NumberingInfo } from './types.js';

/** Internal immutable provenance. The public semantic NumberingInfo keeps only
 * its already resolved image box; these facts never imply Word AUTO metrics. */
export type NativePictureBulletOrigin =
  | Readonly<{ kind: 'direct'; fc: number }>
  | Readonly<{ kind: 'piece'; fc: number; prm: number }>
  | Readonly<{ kind: 'listLevel'; instance: number; list: number; level: number }>;
export interface NativeReadingPictureBullet {
  readonly resourceKey: string;
  readonly rawPbiFlags: number;
  readonly flagsOrigin: NativePictureBulletOrigin;
  readonly indexOrigin: NativePictureBulletOrigin;
  readonly relativeCp: number;
  readonly picfOffset: number;
  readonly shape: number;
  readonly rawShapeFlags: number;
  readonly goalTwips: readonly [number, number];
  readonly scalePerMille: readonly [number, number];
  readonly pibFlags: Readonly<{ key: number; value: number }> | null;
  readonly clientAnchor: readonly number[] | null;
  readonly clientAnchorOptions: number | null;
}
export interface NativeReadingPictureBulletWire {
  readonly __nativeReadingPictureBullet?: NativeReadingPictureBullet | null;
}
function record(value: unknown, fields: readonly string[]): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError('Invalid native picture-bullet owner');
  const result = value as Record<string, unknown>;
  if (Object.keys(result).length !== fields.length || fields.some(field => !Object.hasOwn(result, field)))
    throw new TypeError('Invalid native picture-bullet owner fields');
  return result;
}
function integer(value: unknown, minimum: number, maximum: number): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < minimum || value > maximum)
    throw new TypeError('Invalid native picture-bullet owner integer');
  return value;
}
function origin(value: unknown): NativePictureBulletOrigin {
  if (!value || typeof value !== 'object') throw new TypeError('Invalid native picture-bullet origin');
  const kind = (value as { kind?: unknown }).kind;
  if (kind === 'direct') {
    const r = record(value, ['kind', 'fc']);
    return Object.freeze({ kind, fc: integer(r.fc, 0, 0xffffffff) });
  }
  if (kind === 'piece') {
    const r = record(value, ['kind', 'fc', 'prm']);
    return Object.freeze({ kind, fc: integer(r.fc, 0, 0xffffffff), prm: integer(r.prm, 0, 65535) });
  }
  if (kind === 'listLevel') {
    const r = record(value, ['kind', 'instance', 'list', 'level']);
    return Object.freeze({ kind, instance: integer(r.instance, 0, Number.MAX_SAFE_INTEGER),
      list: integer(r.list, 0, Number.MAX_SAFE_INTEGER), level: integer(r.level, 0, 8) });
  }
  throw new TypeError('Native picture-bullet origin is not a legal direct/list owner');
}
function pair(value: unknown, maximum: number): readonly [number, number] {
  if (!Array.isArray(value) || value.length !== 2) throw new TypeError('Invalid native picture-bullet extent owner');
  const result: [number, number] = [integer(value[0], 1, maximum), integer(value[1], 1, maximum)];
  return Object.freeze(result);
}
export function acquireNativeReadingPictureBullet(
  numbering: (NumberingInfo & NativeReadingPictureBulletWire) | null,
): NativeReadingPictureBullet | undefined {
  const wire = numbering?.__nativeReadingPictureBullet;
  if (wire == null) return undefined;
  const r = record(wire, ['resourceKey', 'rawPbiFlags', 'flagsOrigin', 'indexOrigin', 'relativeCp', 'picfOffset',
    'shape', 'rawShapeFlags', 'goalTwips', 'scalePerMille', 'pibFlags', 'clientAnchor', 'clientAnchorOptions']);
  const rawPbiFlags = integer(r.rawPbiFlags, 0, 65535);
  const picfOffset = integer(r.picfOffset, 0, Number.MAX_SAFE_INTEGER);
  const rawShapeFlags = integer(r.rawShapeFlags, 0, 0xffffffff);
  if ((rawPbiFlags & 1) === 0 || r.shape !== 75 || (rawShapeFlags & 0x11d) !== 0
    || typeof r.resourceKey !== 'string' || r.resourceKey.length === 0
    || numbering?.picBulletImagePath !== r.resourceKey
    || typeof numbering.picBulletWidthPt !== 'number' || !Number.isFinite(numbering.picBulletWidthPt) || numbering.picBulletWidthPt <= 0
    || typeof numbering.picBulletHeightPt !== 'number' || !Number.isFinite(numbering.picBulletHeightPt) || numbering.picBulletHeightPt <= 0
    || !numbering.picBulletMimeType) throw new TypeError('Native reading bullet lost its selected passive carrier');
  let pibFlags: NativeReadingPictureBullet['pibFlags'] = null;
  if (r.pibFlags !== null) {
    const property = record(r.pibFlags, ['key', 'value']);
    const key = integer(property.key, 0, 65535);
    if (key !== 0x106) throw new TypeError('Invalid native picture-bullet pibFlags');
    const value = integer(property.value, 0, 0xffffffff);
    // MS-ODRAW 2.4.8: exact enum/dependency domain, then the bounded
    // embedded Comment subset. Never mask undefined bits into a valid mode.
    if ((value & ~0x0f) !== 0 || (value & 3) === 3
      || ((value & 4) !== 0 && (value & 8) === 0)
      || ((value & 8) !== 0 && (value & 3) === 0))
      throw new TypeError('Invalid native picture-bullet MSOBLIPFLAGS');
    if (value !== 0) throw new TypeError('Named or linked native picture bullet has no stored-size reading consumer');
    pibFlags = Object.freeze({ key, value });
  }
  let clientAnchor: readonly number[] | null = null;
  let clientAnchorOptions: number | null = null;
  if (r.clientAnchor !== null) {
    if (!Array.isArray(r.clientAnchor) || r.clientAnchor.length > 256 * 1024 * 1024)
      throw new TypeError('Invalid native picture-bullet ClientAnchor');
    clientAnchor = Object.freeze(r.clientAnchor.map(byte => integer(byte, 0, 255)));
    clientAnchorOptions = integer(r.clientAnchorOptions, 0, 65535);
  } else if (r.clientAnchorOptions !== null) throw new TypeError('Native picture-bullet anchor options lost their owner');
  return Object.freeze({ resourceKey: r.resourceKey as string, rawPbiFlags, flagsOrigin: origin(r.flagsOrigin), indexOrigin: origin(r.indexOrigin),
    relativeCp: integer(r.relativeCp, 0, 0x7fffffff), picfOffset, shape: 75, rawShapeFlags,
    goalTwips: pair(r.goalTwips, 32767), scalePerMille: pair(r.scalePerMille, 65535),
    pibFlags, clientAnchor, clientAnchorOptions });
}

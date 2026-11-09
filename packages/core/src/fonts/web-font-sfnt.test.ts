import { readFileSync } from 'node:fs';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { instanceOpenTypeLineMetrics, parseOpenTypeResourceMetrics } from './open-type-metrics.js';
import { metricSfnt, nativeBrotliAvailable } from './web-font-sfnt.js';

// Synthetic fixtures (scripts/generate-fixture-font.py).
const fixture = (name: string) => {
  const bytes = readFileSync(new URL(`../../../../tests/fixtures/fonts/${name}`, import.meta.url));
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
};

const parse = async (name: string) => {
  const sfnt = await metricSfnt(fixture(name));
  return sfnt instanceof Uint8Array ? parseOpenTypeResourceMetrics(sfnt, undefined, { variations: true }) : sfnt;
};

afterEach(() => { vi.unstubAllGlobals(); });

describe('metric tables of web font files', () => {
  it('reads the same metric tables from TTF, WOFF and WOFF2', async () => {
    const ttf = await parse('ooxml-fixture-gap.ttf');
    expect(ttf).toMatchObject({ unitsPerEm: 1000, hheaAscent: 800, hheaDescent: -200, hheaLineGap: 300, farEastCodePage: false });
    expect(await parse('ooxml-fixture-gap.woff')).toEqual(ttf);
    expect(nativeBrotliAvailable()).toBe(true);
    expect(await parse('ooxml-fixture-gap.woff2')).toEqual(ttf);
  });

  it('reports WOFF2 as unsupported where the engine cannot inflate Brotli', async () => {
    const Native = globalThis.DecompressionStream;
    vi.stubGlobal('DecompressionStream', class extends Native {
      constructor(format: CompressionFormat) {
        if ((format as string) === 'brotli') throw new TypeError('unsupported format');
        super(format);
      }
    });
    expect(nativeBrotliAvailable()).toBe(false);
    expect(await metricSfnt(fixture('ooxml-fixture-gap.woff2'))).toBe('unsupported');
    // WOFF 1.0 only needs deflate.
    expect(await metricSfnt(fixture('ooxml-fixture-gap.woff'))).toBeInstanceOf(Uint8Array);
  });

  it('rejects truncated or unknown files', async () => {
    const woff2 = new Uint8Array(fixture('ooxml-fixture-gap.woff2'));
    expect(await metricSfnt(woff2.slice(0, woff2.byteLength - 40).buffer)).toBeNull();
    const woff = new Uint8Array(fixture('ooxml-fixture-gap.woff'));
    expect(await metricSfnt(woff.slice(0, 60).buffer)).toBeNull();
    expect(await metricSfnt(new Uint8Array(64).fill(7).buffer)).toBeNull();
  });
});

describe('variable font line metrics (MVAR)', () => {
  it('declines a variable font unless the caller instances it', () => {
    const bytes = new Uint8Array(fixture('ooxml-fixture-gap-var.ttf'));
    expect(parseOpenTypeResourceMetrics(bytes)).toBeNull();
    expect(parseOpenTypeResourceMetrics(bytes, undefined, { variations: true })?.variation?.axes)
      .toEqual([{ tag: 'wght', min: 100, default: 400, max: 900 }]);
  });

  it('applies the MVAR deltas of the instance (default instance without axis values)', () => {
    // Masters (wght 100/400/900): ascender 800/800/900, lineGap 300/300/500.
    const metrics = parseOpenTypeResourceMetrics(new Uint8Array(fixture('ooxml-fixture-gap-var.ttf')), undefined, { variations: true })!;
    const at = (wght?: number) => {
      const instance = instanceOpenTypeLineMetrics(metrics, wght === undefined ? {} : { wght })!;
      return [instance.hheaAscent, instance.hheaLineGap, instance.typoAscent, instance.typoLineGap, instance.winAscent]
        .map((value) => Math.round(value! * 1000) / 1000);
    };
    expect(at()).toEqual([800, 300, 800, 300, 800]);
    expect(at(400)).toEqual([800, 300, 800, 300, 800]);
    // The fixture's MVAR varies hasc, hlgp and hcla (its typo values do not
    // vary): [hhea ascender, hhea lineGap, typo ascender, typo lineGap, usWinAscent].
    expect(at(900)).toEqual([900, 500, 800, 300, 900]);
    // Linear between the Regular (400) and Black (900) masters.
    expect(at(700)).toEqual([860, 420, 800, 300, 860]);
    // Below the default the Thin master equals Regular.
    expect(at(100)).toEqual([800, 300, 800, 300, 800]);
    expect(instanceOpenTypeLineMetrics(metrics, { wght: 700 })!.variation).toBeUndefined();
  });
});

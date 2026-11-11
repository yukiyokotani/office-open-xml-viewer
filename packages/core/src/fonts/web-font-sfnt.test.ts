import { readFileSync } from 'node:fs';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { instanceOpenTypeLineMetrics, parseOpenTypeResourceMetrics } from './open-type-metrics.js';
import { MAX_INFLATED_BYTES, metricSfnt, nativeBrotliAvailable } from './web-font-sfnt.js';

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

  /** The WOFF 1.0 fixture with its table directory edited. */
  const editedWoff = (edit: (view: DataView, records: number[]) => void) => {
    const bytes = new Uint8Array(fixture('ooxml-fixture-gap.woff')).slice();
    const view = new DataView(bytes.buffer);
    const records = Array.from({ length: view.getUint16(12) }, (_, index) => 44 + index * 20);
    edit(view, records);
    return bytes.buffer;
  };
  const tagAt = (view: DataView, at: number) => String.fromCharCode(...[0, 1, 2, 3].map((i) => view.getUint8(at + i)));

  it('rejects a directory that repeats a tag', async () => {
    const repeated = editedWoff((view, records) => {
      const [first, second] = records as [number, number];
      view.setUint32(second, view.getUint32(first));
    });
    expect(await metricSfnt(repeated)).toBeNull();
  });

  it('holds one inflation budget per file, checked before allocating', async () => {
    // Two compressed kept tables, each within the budget alone, together
    // beyond it: rejected from the declared sizes, before any buffer exists.
    const over = editedWoff((view, records) => {
      const compressed = records.filter((at) => ['cmap', 'head', 'hhea', 'OS/2'].includes(tagAt(view, at))
        && view.getUint32(at + 8) < view.getUint32(at + 12));
      expect(compressed.length).toBeGreaterThanOrEqual(2);
      for (const at of compressed.slice(0, 2)) view.setUint32(at + 12, Math.floor(MAX_INFLATED_BYTES / 2) + 1);
    });
    expect(await metricSfnt(over)).toBeNull();
  });

  it('counts the assembled sfnt itself in the budget', async () => {
    // One kept table declared just under the budget: alone it would fit, but
    // the sfnt it is assembled into (header, records, the other tables) does
    // not. Rejected from the declared sizes, before the buffer exists.
    const over = editedWoff((view, records) => {
      const compressed = records.find((at) => tagAt(view, at) === 'cmap' && view.getUint32(at + 8) < view.getUint32(at + 12))
        ?? records.find((at) => ['head', 'hhea', 'OS/2'].includes(tagAt(view, at)) && view.getUint32(at + 8) < view.getUint32(at + 12));
      expect(compressed).toBeDefined();
      view.setUint32(compressed! + 12, MAX_INFLATED_BYTES - 8);
    });
    expect(await metricSfnt(over)).toBeNull();
  });

  it('checks Brotli support before allocating any WOFF2 table', async () => {
    const Native = globalThis.DecompressionStream;
    vi.stubGlobal('DecompressionStream', class extends Native {
      constructor(format: CompressionFormat) {
        if ((format as string) === 'brotli') throw new TypeError('unsupported format');
        super(format);
      }
    });
    // The directory pass only reads sizes; the support check answers before
    // any table buffer is sized (web-font-sfnt.ts woff2).
    expect(await metricSfnt(fixture('ooxml-fixture-gap.woff2'))).toBe('unsupported');
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
    // MVAR value tags (OpenType MVAR): hasc/hlgp vary OS/2 sTypoAscender /
    // sTypoLineGap and hcla usWinAscent; hhea has no tag and keeps its
    // default values: [hhea ascender, hhea lineGap, typo ascender, typo lineGap, usWinAscent].
    expect(at(900)).toEqual([800, 300, 900, 500, 900]);
    // Linear between the Regular (400) and Black (900) masters.
    expect(at(700)).toEqual([800, 300, 860, 420, 860]);
    // Below the default the Thin master equals Regular.
    expect(at(100)).toEqual([800, 300, 800, 300, 800]);
    expect(instanceOpenTypeLineMetrics(metrics, { wght: 700 })!.variation).toBeUndefined();
  });
});

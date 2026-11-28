import { afterEach, describe, expect, it, vi } from 'vitest';
import { nativeRouteClusterOwned } from './native-route-line-metrics.js';
import { PROBE_EM, SENTINELS, SENTINEL_UPM } from './route-sentinels.js';

const scale = PROBE_EM / SENTINEL_UPM;

/** A backend whose contexts draw a sentinel by its own geometry and any other
 * font by one fixed box (so a route pair is equal and the control pair differs). */
function installBackend(): { created: () => number } {
  let created = 0;
  class FakeOffscreenCanvas {
    getContext() {
      created += 1;
      const context = {
        font: '10px serif',
        measureText() {
          const sentinel = SENTINELS.find((candidate) => context.font.includes(candidate.family));
          return sentinel
            ? { width: sentinel.advance * scale, actualBoundingBoxLeft: 0, actualBoundingBoxRight: sentinel.advance * scale,
              actualBoundingBoxAscent: sentinel.inkUp * scale, actualBoundingBoxDescent: sentinel.inkDown * scale }
            : { width: 5, actualBoundingBoxLeft: 0, actualBoundingBoxRight: 5, actualBoundingBoxAscent: 1, actualBoundingBoxDescent: 0 };
        },
      };
      return context;
    }
  }
  vi.stubGlobal('OffscreenCanvas', FakeOffscreenCanvas);
  return { created: () => created };
}

describe('the realm ownership probe', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('retries after a creation that found no backend', () => {
    const fonts = ['10px Route', '10px Route'] as const;
    vi.stubGlobal('OffscreenCanvas', undefined);
    vi.stubGlobal('document', undefined);
    expect(nativeRouteClusterOwned(fonts, 'a')).toBeUndefined();
    const backend = installBackend();
    expect(nativeRouteClusterOwned(fonts, 'a')).toBe(true);
    expect(backend.created()).toBeGreaterThan(0);
  });
});

import { describe, expect, it } from 'vitest';
import type { DocumentLayout, LayoutServices } from './types.js';
import { LayoutVariantStore } from './variant-store.js';

function services(textFingerprint: string): LayoutServices {
  return {
    text: { fingerprint: textFingerprint } as LayoutServices['text'],
    images: { fingerprint: 'images:test' } as LayoutServices['images'],
    math: { fingerprint: 'math:test' } as LayoutServices['math'],
  };
}

function emptyLayout(): DocumentLayout {
  return { pages: [], diagnostics: [] };
}

describe('LayoutVariantStore', () => {
  it('memoizes immutable layouts by the actual service and option fingerprints', () => {
    const builds: number[] = [];
    const store = new LayoutVariantStore(
      services('text:a'),
      { currentDateMs: 100 },
      (options) => {
        builds.push(options.currentDateMs);
        return emptyLayout();
      },
    );

    const defaultLayout = store.defaultLayout;
    expect(store.layoutFor({ currentDateMs: 100 })).toBe(defaultLayout);
    expect(store.layoutFor({ currentDateMs: 101 })).not.toBe(defaultLayout);
    expect(store.layoutFor({ currentDateMs: 101 })).toBe(store.layoutFor({ currentDateMs: 101 }));
    expect(builds).toEqual([100, 101]);
    expect(Object.isFrozen(defaultLayout)).toBe(true);
  });

  it('does not let a non-default variant replace load-time default metadata ownership', () => {
    const store = new LayoutVariantStore(
      services('text:a'),
      { currentDateMs: 100 },
      (options) => ({
        pages: [],
        diagnostics: [{
          code: 'UNSUPPORTED_FEATURE',
          severity: 'warning',
          message: String(options.currentDateMs),
        }],
      }),
    );

    const before = store.defaultLayout;
    store.layoutFor({ currentDateMs: 200 });

    expect(store.defaultLayout).toBe(before);
    expect(store.defaultLayout.diagnostics[0]?.message).toBe('100');
  });
});

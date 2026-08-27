import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const read = (path: string): string => readFileSync(new URL(path, import.meta.url), 'utf8');

const viewerSources = [
  './components/BuiltInCommentViewer.astro',
  './components/CommentListNavigationDemo.astro',
  './components/LiveShowcase.astro',
  './components/ReviewDemo.astro',
  './lib/demo-snippets.ts',
  './lib/demos.ts',
  './pages/selection-context.astro',
  './pages/try.astro',
] as const;

describe('official-site viewer BFCache lifecycle', () => {
  for (const path of viewerSources) {
    it(`does not destroy persisted viewers in ${path}`, () => {
      const source = read(path);
      const pagehideListeners = source.match(/addEventListener\('pagehide'/g) ?? [];
      const persistedGuards = source.match(/if \(event\.persisted\) return;/g) ?? [];
      expect(pagehideListeners.length).toBeGreaterThan(0);
      expect(persistedGuards).toHaveLength(pagehideListeners.length);
    });
  }
});

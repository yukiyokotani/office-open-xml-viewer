import assert from 'node:assert/strict';
import test from 'node:test';
import { dotRelativeImportMetaUrls } from '../vite.config.ts';

test('bundled sibling URLs become dot-relative (#2134)', () => {
  assert.equal(dotRelativeImportMetaUrls('new URL("assets/render-worker-a.js", import.meta.url).href'),
    'new URL("./assets/render-worker-a.js", import.meta.url).href');
  assert.equal(dotRelativeImportMetaUrls("new URL('x_bg.wasm',import.meta.url)"),
    "new URL('./x_bg.wasm',import.meta.url)");
});

test('dot-relative, parent, root-absolute and scheme URLs are kept', () => {
  for (const code of [
    'new URL("./a.mjs", import.meta.url)',
    'new URL("../a.mjs", import.meta.url)',
    'new URL("/a.mjs", import.meta.url)',
    'new URL("https://cdn.example/a.mjs", import.meta.url)',
    'new URL("data:text/javascript,0", import.meta.url)',
    'new URL("a.mjs", base)',
  ]) assert.equal(dotRelativeImportMetaUrls(code), code);
});

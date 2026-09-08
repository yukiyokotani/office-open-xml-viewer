import assert from 'node:assert/strict';
import test from 'node:test';
import { wasmAssetUrl } from '../vite.config.ts';

const bytes = new Uint8Array([1, 2, 3]);

for (const [path, expectedName] of [
  ['/tmp/wasm-direct-doc/legacy_office_converter_bg.wasm', 'legacy_doc_direct_bg.wasm'],
  ['/tmp/wasm-direct-ppt/legacy_office_converter_bg.wasm', 'legacy_ppt_direct_bg.wasm'],
  ['/tmp/wasm-direct-xls/legacy_office_converter_bg.wasm', 'legacy_xls_direct_bg.wasm'],
  ['/tmp/wasm/docx_parser_bg.wasm', 'docx_parser_bg.wasm'],
]) {
  test(`emits ${path} as ${expectedName}`, async () => {
    const reads = [];
    const emitted = [];
    const plugin = wasmAssetUrl(async (filePath) => {
      reads.push(filePath);
      return bytes;
    });
    const result = await plugin.load.call({
      emitFile(asset) {
        emitted.push(asset);
        return 'wasm-ref';
      },
    }, `${path}?url`);

    assert.deepEqual(reads, [path]);
    assert.deepEqual(emitted, [{ type: 'asset', name: expectedName, source: bytes }]);
    assert.equal(result, 'export default import.meta.ROLLUP_FILE_URL_wasm-ref;');
  });
}

test('ignores imports without the URL suffix', async () => {
  let read = false;
  const plugin = wasmAssetUrl(async () => {
    read = true;
    return bytes;
  });
  const result = await plugin.load.call({ emitFile() { throw new Error('unexpected emit'); } }, '/tmp/module.js');
  assert.equal(result, null);
  assert.equal(read, false);
});

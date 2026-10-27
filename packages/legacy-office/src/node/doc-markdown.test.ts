// Binary Word markdown export: the direct DOC archive carries the
// `to_markdown` model source capability the DOCX worker checks before
// Markdown conversion, and projects the same GitHub-flavoured markdown as
// the equivalent DOCX (shared docx-model renderer).
import { expect, it } from 'vitest';
import { buildDocFixture, concat, little16 } from '../test-fixtures.js';
import { TEST_SOURCE_URLS } from '../test-sources.js';
import { openModelSource } from '../legacy-doc-source-module.js';

const open = (bytes: Uint8Array) => openModelSource(bytes, { wasmUrl: TEST_SOURCE_URLS.doc.wasmUrl });

it('exposes to_markdown and projects paragraphs and bold like the DOCX path', async () => {
  const plain = await open(buildDocFixture({ text: 'Alpha\rBeta\r' }));
  try {
    expect(typeof plain.archive.to_markdown).toBe('function');
    expect(plain.archive.to_markdown!()).toBe('Alpha\n\nBeta\n\n');
  } finally { plain.close(); }

  // sprmCFBold on the paragraph's CHPX.
  const bold = await open(buildDocFixture({
    text: 'Bold\r',
    characterProperties: concat(little16(0x0835), new Uint8Array([1])),
  }));
  try {
    expect(bold.archive.to_markdown!()).toBe('**Bold**\n\n');
  } finally { bold.close(); }
});

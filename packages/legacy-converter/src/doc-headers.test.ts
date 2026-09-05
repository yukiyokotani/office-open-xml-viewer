import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import initDocx, { DocxArchive } from '../../docx/src/wasm/docx_parser.js';
import { createLegacyOfficeWasmConverter } from './index.js';
import { buildDocFixture, concat, little16 } from './test-fixtures.js';

const converter = createLegacyOfficeWasmConverter({ wasm: await readFile(new URL('./wasm/legacy_office_converter_bg.wasm', import.meta.url)) });
await initDocx({ module_or_path: await readFile(new URL('../../docx/src/wasm/docx_parser_bg.wasm', import.meta.url)) });
async function convert(bytes: Uint8Array) {
  const result = await converter.convert({ bytes, from: 'doc', to: 'docx', signal: new AbortController().signal, maxOutputBytes: 16 * 1024 * 1024 });
  const archive = new DocxArchive(new Uint8Array(result.bytes));
  try {
    const contentTypes = new TextDecoder().decode(archive.extract_image('[Content_Types].xml'));
    const paths = ['[Content_Types].xml', 'word/_rels/document.xml.rels', ...Array.from(contentTypes.matchAll(/PartName="\/([^"]+)"/g), m => m[1])];
    return Object.fromEntries(paths.map(path => [path, new TextDecoder().decode(archive.extract_image(path))]));
  } finally { archive.free(); }
}

describe('binary Word header/footer stories', () => {
  it('restores all six variants after footnotes, with physical piece formatting and no guard paragraphs', async () => {
    const parts = await convert(buildDocFixture({
      text: 'Body\r', footnotes: 'NOT A HEADER\r',
      headers: ['EH\r', 'OH😀\r', 'EF\r', 'OF\r', 'FH\r', 'FF\r'],
      characterProperties: concat(little16(0x0835), new Uint8Array([1])),
      defaultTabTwips: 720, facingPages: true,
      sectionProperties: concat(little16(0x300a), new Uint8Array([1])),
    }));
    const doc = parts['word/document.xml'];
    expect(doc).toContain('Body');
    expect(doc).not.toContain('NOT A HEADER');
    for (const [i, [kind, variant, text]] of [
      ['header', 'even', 'EH'], ['header', 'default', 'OH😀'],
      ['footer', 'even', 'EF'], ['footer', 'default', 'OF'],
      ['header', 'first', 'FH'], ['footer', 'first', 'FF'],
    ].entries()) {
      const part = parts[`word/${kind}${i + 1}.xml`];
      expect(part).toContain(text);
      expect(part).toContain('<w:b w:val="1"/>');
      expect(part.match(/<w:p>/g)).toHaveLength(1);
      expect(part).not.toContain('sectPr');
      expect(doc).toContain(`w:type="${variant}" r:id="rIdHf${i + 1}"`);
      expect(parts['word/_rels/document.xml.rels']).toContain(`Target="${kind}${i + 1}.xml"`);
      expect(parts['[Content_Types].xml']).toContain(`/word/${kind}${i + 1}.xml`);
    }
    expect(parts['word/settings.xml']).toContain('<w:evenAndOddHeaders/>');
    expect(doc).toContain('<w:titlePg w:val="1"/>');
  });

  it('keeps an explicit blank while absent variants do not create a part', async () => {
    const parts = await convert(buildDocFixture({ text: 'Body\r', headers: ['', '\r', '', 'Footer\r', '', ''] }));
    expect(parts['word/header2.xml'].match(/<w:p>/g)).toHaveLength(1);
    expect(parts['word/document.xml'].match(/Reference /g)).toHaveLength(2);
    expect(Object.keys(parts).filter(p => /word\/(header|footer)\d/.test(p))).toHaveLength(2);
  });

  it.each([false, true])('restores passive page fields while respecting lock=%s and stripping active instructions', async lockedHeaderFields => {
    const text = 'Page \u0013 PAGE \\* MERGEFORMAT \u001499\u0015 of \u0013NUMPAGES\u001499\u0015; \u0013INCLUDETEXT "file://secret"\u0014cached\u0015\r';
    const parts = await convert(buildDocFixture({ text: 'Body\r', headers: ['', '', '', text, '', ''], lockedHeaderFields }));
    const footer = parts['word/footer4.xml'];
    expect(footer).not.toContain('INCLUDETEXT');
    expect(footer).not.toContain('file://secret');
    expect(footer).toContain('cached');
    if (lockedHeaderFields) expect(footer).not.toContain('instrText');
    else {
      expect(footer).toContain(' PAGE \\* MERGEFORMAT ');
      expect(footer).toContain('NUMPAGES');
      expect(footer.match(/fldCharType="begin"/g)).toHaveLength(2);
      expect(footer.match(/fldCharType="end"/g)).toHaveLength(2);
    }
  });
});

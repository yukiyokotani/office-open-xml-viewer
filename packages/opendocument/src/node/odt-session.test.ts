// OpenDocument text through the Node DOCX session: the claimed package is
// admitted into the DOCX model, or rejected as a whole with a typed,
// content-free reason.
import { describe, expect, it } from 'vitest';
import { OoxmlError } from '@silurus/ooxml-core';
import { buildOdtFixture } from '../odt-fixture.js';
import { testOdtSource } from '../test-sources.js';
import { materializeDocxDocument } from './node-facade.js';

type Run = { type: string; text?: string; fontSize?: number; bold?: boolean; fontFamily?: string };
type Paragraph = { type: string; runs: Run[]; spaceBefore: number; spaceAfter: number; lineSpacing?: { value: number; rule: string } };
type Model = { section: { pageWidth: number; marginLeft: number }; body: Paragraph[] };

const open = async (bytes: Uint8Array) =>
  (await materializeDocxDocument(bytes, { odt: testOdtSource() })) as unknown as Model;

describe('ODT through the DOCX Node session', () => {
  it('maps paragraphs, spans, collapsed white space and page geometry', async () => {
    const model = await open(buildOdtFixture({
      styles: '<style:style style:name="Body" style:family="paragraph"><style:paragraph-properties fo:margin-top="6pt" fo:margin-bottom="10pt"/></style:style>'
        + '<style:style style:name="Tall" style:family="paragraph"><style:paragraph-properties fo:margin-top="6pt" fo:margin-bottom="10pt"/></style:style>',
      automaticStyles: '<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold" fo:font-size="14pt"/></style:style>',
      body: '<text:p text:style-name="Body">  Hello\n   <text:span text:style-name="T1">bold</text:span><text:s text:c="2"/>end  </text:p>'
        + '<text:p text:style-name="Tall">second</text:p>',
    }));
    expect(model.section.pageWidth).toBeCloseTo(595.28, 1);
    expect(model.section.marginLeft).toBeCloseTo(56.69, 1);
    const [first, second] = model.body;
    expect(first?.runs.map(r => r.text)).toEqual(['Hello ', 'bold', '  end']);
    expect(first?.runs[1]).toMatchObject({ bold: true, fontSize: 14, fontFamily: 'Liberation Serif' });
    // Adjacent margins add (10 + 6) and are carried by the next paragraph.
    expect(first).toMatchObject({ spaceBefore: 6, spaceAfter: 0 });
    expect(second).toMatchObject({ spaceBefore: 16, spaceAfter: 0 });
  });

  it('rejects content outside the admitted subset with a typed reason', async () => {
    const list = buildOdtFixture({ body: '<text:list><text:list-item><text:p>item</text:p></text:list-item></text:list>' });
    await expect(open(list)).rejects.toMatchObject({
      code: 'unsupported-content', details: { format: 'odt', reason: 'source-subset' },
    });
    const header = buildOdtFixture({ body: '<text:p>x</text:p>', masterContent: '<style:header><text:p>h</text:p></style:header>' });
    await expect(open(header)).rejects.toMatchObject({ code: 'unsupported-content' });
    const encrypted = buildOdtFixture({
      body: '<text:p>x</text:p>',
      manifestExtra: '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"><manifest:encryption-data/></manifest:file-entry>',
    });
    await expect(open(encrypted)).rejects.toBeInstanceOf(OoxmlError);
    await expect(open(encrypted)).rejects.toMatchObject({ code: 'unsupported-encryption' });
  });
});

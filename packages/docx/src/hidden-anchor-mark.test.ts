import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { measureContext } from './test-support/tab-fitting.test-support.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import { textRunsForPage } from './text-run-projection.js';

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
  + 'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const encode = (text: string) => new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${text}`);

const NUMBERING = `<w:numbering ${NS}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>`
  + '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>'
  + '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>';
const WP = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
  + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
  + 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

/** What a page shows after a vanished mark (with paragraph properties
 * `properties`) holding `content`: the top of the following text, every
 * text piece painted, the paragraph decoration painted and the drawings. */
function observe(content: string, properties = '', inCell = false) {
  const paragraphs = `<w:p><w:pPr>${properties}<w:rPr><w:vanish/></w:rPr></w:pPr>${content}</w:p><w:p><w:r><w:t>After</w:t></w:r></w:p>`;
  const body = inCell
    ? '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
      + `<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${paragraphs}</w:tc></w:tr></w:tbl><w:p/>`
    : paragraphs;
  const bytes = storeZip(new Map([
    ['[Content_Types].xml', encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>')],
    ['_rels/.rels', encode(`<Relationships xmlns="${PKG}"><Relationship Id="rId1" `
      + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/_rels/document.xml.rels', encode(`<Relationships xmlns="${PKG}"><Relationship Id="rIdN" `
      + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
      + '<Relationship Id="rIdL" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
      + 'Target="https://example.invalid/picture.png" TargetMode="External"/></Relationships>')],
    ['word/numbering.xml', encode(NUMBERING)],
    ['word/document.xml', encode(`<w:document ${NS} ${WP}><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>`
      + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr></w:body></w:document>')],
  ]));
  const archive = new DocxArchive(bytes);
  let parsed;
  try {
    parsed = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse())));
  } finally { archive.free(); }
  const model = parsed.document;
  const layout = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext(5) }), { currentDateMs: 0 });
  // Every paragraph node of the page with what it paints of its own.
  const painted: string[] = [];
  const seen = new Set<unknown>();
  const visit = (node: unknown): void => {
    if (!node || typeof node !== 'object' || seen.has(node)) return;
    seen.add(node);
    if (Array.isArray(node)) { node.forEach(visit); return; }
    const record = node as Record<string, unknown>;
    if (record.kind === 'paragraph' && Array.isArray(record.lines)) {
      for (const line of record.lines as { placements: { kind: string; text?: string; leader?: string }[] }[]) {
        for (const placement of line.placements) {
          if (placement.kind === 'text') painted.push(`text:${placement.text}`);
          else if (placement.kind === 'resource') painted.push('resource');
          else if (placement.kind === 'tab' && placement.leader && placement.leader !== 'none') painted.push('leader');
        }
      }
      if ((record.borders as unknown[] | undefined)?.length) painted.push('border');
      if (record.shading) painted.push('shading');
      if (record.lineNumbers) painted.push('line-number');
    }
    for (const [key, entry] of Object.entries(record)) if (key !== 'paintOrder') visit(entry);
  };
  visit(layout.pages[0]);
  return {
    top: textRunsForPage(layout, 0, { scale: 1 }).find((run) => run.text === 'After')!.y,
    painted,
    drawings: layout.pages.flatMap((page) => page.layers.paintOrder.filter((entry) => entry.kind === 'drawing')).length,
  };
}

const VML_RECT = '<w:r><w:pict><v:rect style="position:absolute;margin-left:100pt;margin-top:300pt;width:20pt;height:20pt;'
  + 'mso-position-horizontal-relative:page;mso-position-vertical-relative:page" fillcolor="red" stroked="f"/></w:pict></w:r>';

describe('a vanished paragraph mark that anchors only a floating VML shape', () => {
  it('joins the next paragraph as the empty vanished mark does, and still paints the shape', () => {
    const empty = observe('');
    const anchored = observe(VML_RECT);
    expect(anchored.top).toBe(empty.top);
    expect(anchored.drawings).toBe(empty.drawings + 1);
  });

  it.each([['the body', false], ['a table cell', true]])('paints no list marker, border, shading or line number of its own in %s', (_where, inCell) => {
    const decoration = '<w:pBdr><w:top w:val="single" w:sz="8" w:space="1" w:color="FF0000"/></w:pBdr>'
      + '<w:shd w:val="clear" w:fill="FFFF00"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>';
    const empty = observe('', decoration, inCell);
    const anchored = observe(VML_RECT, decoration, inCell);
    expect(empty.painted).toEqual(['text:After']);
    expect(anchored.painted).toEqual(empty.painted);
    expect(anchored.top).toBe(empty.top);
    expect(anchored.drawings).toBe(empty.drawings + 1);
  });
});

describe('a vanished paragraph mark that anchors only a floating picture whose image is unavailable', () => {
  it.each([['the body', false], ['a table cell', true]])('joins the next paragraph as the empty vanished mark does in %s', (_where, inCell) => {
    const linked = '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" '
      + 'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
      + '<wp:positionH relativeFrom="page"><wp:posOffset>1270000</wp:posOffset></wp:positionH>'
      + '<wp:positionV relativeFrom="page"><wp:posOffset>3810000</wp:posOffset></wp:positionV>'
      + '<wp:extent cx="254000" cy="254000"/><wp:wrapNone/><wp:docPr id="1" name="Linked"/>'
      + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
      + '<pic:nvPicPr><pic:cNvPr id="1" name="Linked"/><pic:cNvPicPr/></pic:nvPicPr>'
      + '<pic:blipFill><a:blip r:link="rIdL"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
      + '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="254000" cy="254000"/></a:xfrm>'
      + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
      + '</wp:anchor></w:drawing></w:r>';
    const empty = observe('', '', inCell);
    const anchored = observe(linked, '', inCell);
    expect(anchored.top).toBe(empty.top);
    expect(anchored.painted).toEqual(empty.painted);
  });
});

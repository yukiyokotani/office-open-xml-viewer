import assert from 'node:assert/strict';
import { mkdtempSync, mkdirSync, realpathSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { crc32, deflateRawSync } from 'node:zlib';
import { buildFastTierManifest, loadFastTierSelection } from './corpus-features.mjs';

/** Tiny deflate ZIP writer for synthetic packages. */
function zip(parts) {
  const locals = [];
  const central = [];
  let offset = 0;
  for (const [name, text] of Object.entries(parts)) {
    const raw = Buffer.from(text);
    const data = deflateRawSync(raw);
    const nameBytes = Buffer.from(name);
    const header = Buffer.alloc(30);
    header.writeUInt32LE(0x04034b50, 0);
    header.writeUInt16LE(8, 8);
    header.writeUInt32LE(crc32(raw), 14);
    header.writeUInt32LE(data.length, 18);
    header.writeUInt32LE(raw.length, 22);
    header.writeUInt16LE(nameBytes.length, 26);
    const entry = Buffer.alloc(46);
    entry.writeUInt32LE(0x02014b50, 0);
    entry.writeUInt16LE(8, 10);
    entry.writeUInt32LE(crc32(raw), 16);
    entry.writeUInt32LE(data.length, 20);
    entry.writeUInt32LE(raw.length, 24);
    entry.writeUInt16LE(nameBytes.length, 28);
    entry.writeUInt32LE(offset, 42);
    locals.push(header, nameBytes, data);
    central.push(entry, nameBytes);
    offset += header.length + nameBytes.length + data.length;
  }
  const directory = Buffer.concat(central);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(central.length / 2, 8);
  end.writeUInt16LE(central.length / 2, 10);
  end.writeUInt32LE(directory.length, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, directory, end]);
}

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const X = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const rels = (...items) => `<Relationships xmlns="${PKG}">${items.map(([id, type, target]) =>
  `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`).join('')}</Relationships>`;

function pptx() {
  const slide = (prefix, extra = '') => `<${prefix}:sld xmlns:${prefix}="${P}" xmlns:a="${A}"><${prefix}:cSld><${prefix}:spTree/></${prefix}:cSld>${extra}</${prefix}:sld>`;
  const layout = (body) => `<p:sldLayout xmlns:p="${P}" xmlns:a="${A}"><p:cSld>${body}</p:cSld></p:sldLayout>`;
  return zip({
    '_rels/.rels': rels(['r1', 'officeDocument', 'ppt/presentation.xml']),
    // The p14 section list repeats slide ids inside extLst; it must not add slots.
    'ppt/presentation.xml': `<p:presentation xmlns:p="${P}" xmlns:r="${REL}" xmlns:p14="urn:p14"><p:sldIdLst>`
      + '<p:sldId id="256" r:id="s1"/><p:sldId id="257" r:id="s2"/><p:sldId id="258" r:id="s3"/></p:sldIdLst>'
      + '<p:extLst><p:ext><p14:sectionLst><p14:section><p14:sldIdLst><p14:sldId id="256"/><p14:sldId id="257"/>'
      + '</p14:sldIdLst></p14:section></p14:sectionLst></p:ext></p:extLst></p:presentation>',
    'ppt/_rels/presentation.xml.rels': rels(
      ['s1', 'slide', 'slides/slide1.xml'], ['s2', 'slide', 'slides/slide2.xml'], ['s3', 'slide', 'slides/slide3.xml']),
    // Slides 1 and 2 differ only by namespace prefix; slide 3 lacks their
    // transition and is distinguished only by its layout's background.
    'ppt/slides/slide1.xml': slide('p', '<p:transition/>'),
    'ppt/slides/slide2.xml': slide('pp', '<pp:transition/>'),
    'ppt/slides/slide3.xml': slide('p'),
    'ppt/slides/_rels/slide1.xml.rels': rels(['l', 'slideLayout', '../slideLayouts/slideLayout1.xml']),
    'ppt/slides/_rels/slide2.xml.rels': rels(['l', 'slideLayout', '../slideLayouts/slideLayout1.xml']),
    'ppt/slides/_rels/slide3.xml.rels': rels(['l', 'slideLayout', '../slideLayouts/slideLayout2.xml']),
    'ppt/slideLayouts/slideLayout1.xml': layout('<p:spTree/>'),
    'ppt/slideLayouts/slideLayout2.xml': layout('<p:bg><p:bgPr><a:gradFill/></p:bgPr></p:bg><p:spTree/>'),
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': rels(['m', 'slideMaster', '../slideMasters/slideMaster1.xml']),
    'ppt/slideLayouts/_rels/slideLayout2.xml.rels': rels(['m', 'slideMaster', '../slideMasters/slideMaster1.xml']),
    'ppt/slideMasters/slideMaster1.xml': `<p:sldMaster xmlns:p="${P}"><p:cSld><p:spTree/></p:cSld></p:sldMaster>`,
    // A master lists every layout; reaching it must not pull sibling layouts into a slide.
    'ppt/slideMasters/_rels/slideMaster1.xml.rels': rels(
      ['a', 'slideLayout', '../slideLayouts/slideLayout1.xml'], ['b', 'slideLayout', '../slideLayouts/slideLayout2.xml']),
  });
}

function xlsx() {
  const sheet = (extra, style) => `<worksheet xmlns="${X}"><sheetData><row r="1"><c r="A1" s="${style}"><v>1</v></c></row></sheetData>${extra}</worksheet>`;
  return zip({
    '_rels/.rels': rels(['r1', 'officeDocument', 'xl/workbook.xml']),
    'xl/workbook.xml': `<workbook xmlns="${X}" xmlns:r="${REL}"><sheets><sheet name="A" sheetId="1" r:id="a"/><sheet name="B" sheetId="2" r:id="b"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels(
      ['a', 'worksheet', 'worksheets/sheet1.xml'], ['b', 'worksheet', 'worksheets/sheet2.xml'], ['s', 'styles', 'styles.xml']),
    // Sheet 2's only distinguishing trait is the aligned cell style it references.
    'xl/worksheets/sheet1.xml': sheet('<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>', 0),
    'xl/worksheets/sheet2.xml': sheet('', 1),
    'xl/styles.xml': `<styleSheet xmlns="${X}"><fonts><font/></fonts><cellXfs count="2"><xf fontId="0"/>`
      + '<xf fontId="0" applyAlignment="1"><alignment wrapText="1"/></xf></cellXfs></styleSheet>',
  });
}

test('fast tier keeps only slides/sheets that add structural features', () => {
  const pptxManifest = buildFastTierManifest('pptx', [{ name: 'pptx/deck.pptx', bytes: pptx() }]);
  assert.deepEqual(pptxManifest.files.map(({ unitCount, units }) => ({ unitCount, units })), [
    { unitCount: 3, units: [0, 2] },
  ]);
  assert.equal(pptxManifest.summary.unitsSelected, 2);

  const xlsxManifest = buildFastTierManifest('xlsx', [{ name: 'xlsx/book.xlsx', bytes: xlsx() }]);
  assert.deepEqual(xlsxManifest.files[0].units, [0, 1]);
});

test('fast tier selection fails closed on a stale or incomplete manifest', () => {
  const root = mkdtempSync(join(realpathSync(tmpdir()), 'ooxml-fast-tier-'));
  try {
    const manifestPath = join(root, 'vrt-fast-tier.json');
    const load = (files) => loadFastTierSelection({ format: 'pptx', files, manifestPath, directory: root });
    mkdirSync(join(root, 'pptx'));
    writeFileSync(join(root, 'pptx/deck.pptx'), pptx());
    assert.throws(() => load(['pptx/deck.pptx']), /missing pptx fast-tier selection manifest/);

    const manifest = buildFastTierManifest('pptx', [{ name: 'pptx/deck.pptx', bytes: pptx() }]);
    writeFileSync(manifestPath, JSON.stringify(manifest));
    assert.deepEqual(load(['pptx/deck.pptx']).files.get('pptx/deck.pptx').units, [0, 2]);

    writeFileSync(join(root, 'pptx/new.pptx'), pptx());
    assert.throws(() => load(['pptx/deck.pptx', 'pptx/new.pptx']), /new\.pptx is not in the manifest/);
    writeFileSync(join(root, 'pptx/deck.pptx'), xlsx());
    assert.throws(() => load(['pptx/deck.pptx']), /deck\.pptx changed since selection/);
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

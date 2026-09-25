/**
 * Inputs that are not Office Open XML packages, one per admission rule in
 * `ooxml_common::opc` (ECMA-376 Part 2): not a readable ZIP (arbitrary bytes and
 * an empty buffer), a ZIP without the `[Content_Types].xml` Media Types stream,
 * and an OPC package without the format's main part. Every load path must
 * reject each of them with `OoxmlError('not-ooxml')`.
 */
export type NonOoxmlFormat = 'docx' | 'xlsx' | 'pptx';

const MAIN_PART: Record<NonOoxmlFormat, string> = {
  docx: 'word/document.xml',
  xlsx: 'xl/workbook.xml',
  pptx: 'ppt/presentation.xml',
};

const CONTENT_TYPES =
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
  + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
  + '</Types>';

const ROOT_RELS =
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>';

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

const u16 = (value: number): number[] => [value & 0xff, (value >>> 8) & 0xff];
const u32 = (value: number): number[] => [
  value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff, (value >>> 24) & 0xff,
];

/** Minimal stored (uncompressed) ZIP of UTF-8 entries. */
function storedZip(files: ReadonlyArray<readonly [string, string]>): Uint8Array<ArrayBuffer> {
  const encoder = new TextEncoder();
  const local: number[] = [];
  const central: number[] = [];
  for (const [name, content] of files) {
    const nameBytes = [...encoder.encode(name)];
    const data = encoder.encode(content);
    const checksum = crc32(data);
    const offset = local.length;
    local.push(
      ...u32(0x04034b50), ...u16(20), ...u16(0), ...u16(0), ...u16(0), ...u16(0),
      ...u32(checksum), ...u32(data.length), ...u32(data.length),
      ...u16(nameBytes.length), ...u16(0), ...nameBytes, ...data,
    );
    central.push(
      ...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(0), ...u16(0), ...u16(0), ...u16(0),
      ...u32(checksum), ...u32(data.length), ...u32(data.length),
      ...u16(nameBytes.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0),
      ...u32(0), ...u32(offset), ...nameBytes,
    );
  }
  const end = [
    ...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(files.length), ...u16(files.length),
    ...u32(central.length), ...u32(local.length), ...u16(0),
  ];
  return Uint8Array.from([...local, ...central, ...end]);
}

/** Named non-OOXML inputs for `format`. Each call returns fresh buffers. */
export function nonOoxmlInputs(
  format: NonOoxmlFormat,
): ReadonlyArray<readonly [name: string, bytes: Uint8Array<ArrayBuffer>]> {
  const main = MAIN_PART[format];
  return [
    ['garbage bytes', Uint8Array.from([1, 2, 3])],
    ['empty buffer', new Uint8Array(0)],
    ['ZIP without [Content_Types].xml', storedZip([[main, '<root/>']])],
    [
      `OPC package without ${main}`,
      storedZip([['[Content_Types].xml', CONTENT_TYPES], ['_rels/.rels', ROOT_RELS]]),
    ],
  ];
}

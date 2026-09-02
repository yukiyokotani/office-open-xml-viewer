import { crc32 } from 'node:zlib';

function u16(value: number): number[] {
  return [value & 0xff, (value >>> 8) & 0xff];
}

function u32(value: number): number[] {
  return [
    value & 0xff,
    (value >>> 8) & 0xff,
    (value >>> 16) & 0xff,
    (value >>> 24) & 0xff,
  ];
}

/** Minimal stored ZIP used by parser-boundary tests. Inputs are authored test
 * XML, never private Office samples or generated golden files. */
export function storedZip(files: Readonly<Record<string, string>>): Uint8Array {
  const encoder = new TextEncoder();
  const localChunks: number[] = [];
  const central: number[] = [];
  let offset = 0;
  for (const [name, content] of Object.entries(files)) {
    const nameBytes = [...encoder.encode(name)];
    const data = [...encoder.encode(content)];
    const checksum = crc32(Uint8Array.from(data)) >>> 0;
    const local = [
      ...u32(0x04034b50), ...u16(20), ...u16(0), ...u16(0), ...u16(0), ...u16(0),
      ...u32(checksum), ...u32(data.length), ...u32(data.length),
      ...u16(nameBytes.length), ...u16(0), ...nameBytes, ...data,
    ];
    central.push(
      ...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(0), ...u16(0), ...u16(0), ...u16(0),
      ...u32(checksum), ...u32(data.length), ...u32(data.length),
      ...u16(nameBytes.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0),
      ...u32(0), ...u32(offset), ...nameBytes,
    );
    localChunks.push(...local);
    offset += local.length;
  }
  const end = [
    ...u32(0x06054b50), ...u16(0), ...u16(0),
    ...u16(Object.keys(files).length), ...u16(Object.keys(files).length),
    ...u32(central.length), ...u32(offset), ...u16(0),
  ];
  return Uint8Array.from([...localChunks, ...central, ...end]);
}

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

export function minimalDocx(documentXml: string): Uint8Array {
  return storedZip({
    '[Content_Types].xml': CONTENT_TYPES,
    '_rels/.rels': ROOT_RELS,
    'word/document.xml': documentXml,
  });
}

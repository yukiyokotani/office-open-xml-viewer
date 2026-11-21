// Deterministic ZIP writer for the public specimen corpus.
//
// Output bytes depend only on the entry names and contents:
//   - entries are sorted by the byte order of their UTF-8 names
//     ([Content_Types].xml therefore comes first: '[' < '_' < 'a'..'z');
//   - every entry is STORED (method 0). Deflate output varies between zlib
//     builds, so compressing would make the committed bytes depend on the Node
//     version that rebuilt them. The parts are a few kilobytes of XML;
//   - every timestamp is the DOS epoch, 1980-01-01 00:00:00;
//   - no extra fields, comments, data descriptors or external attributes.
// OPC (ECMA-376 Part 2, Annex C) permits stored items.

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

export function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

const DOS_TIME = 0x0000; // 00:00:00
const DOS_DATE = 0x0021; // 1980-01-01: ((1980-1980) << 9) | (1 << 5) | 1

/** @param {Record<string, string | Uint8Array>} files part name → content */
export function zipStored(files) {
  const entries = Object.keys(files)
    .map((name) => {
      const value = files[name];
      return {
        name: Buffer.from(name, 'utf8'),
        data: typeof value === 'string' ? Buffer.from(value, 'utf8') : Buffer.from(value),
      };
    })
    .sort((a, b) => Buffer.compare(a.name, b.name));
  const locals = [];
  const centrals = [];
  let offset = 0;
  for (const { name, data } of entries) {
    const crc = crc32(data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4); // version needed: 2.0
    local.writeUInt16LE(0, 6); // flags
    local.writeUInt16LE(0, 8); // method: stored
    local.writeUInt16LE(DOS_TIME, 10);
    local.writeUInt16LE(DOS_DATE, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(name.length, 26);
    local.writeUInt16LE(0, 28);
    locals.push(local, name, data);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4); // version made by: 2.0, MS-DOS
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0, 8);
    central.writeUInt16LE(0, 10);
    central.writeUInt16LE(DOS_TIME, 12);
    central.writeUInt16LE(DOS_DATE, 14);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(data.length, 20);
    central.writeUInt32LE(data.length, 24);
    central.writeUInt16LE(name.length, 28);
    central.writeUInt32LE(offset, 42);
    centrals.push(central, name);
    offset += 30 + name.length + data.length;
  }
  const centralSize = centrals.reduce((sum, b) => sum + b.length, 0);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(centralSize, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, ...centrals, end]);
}

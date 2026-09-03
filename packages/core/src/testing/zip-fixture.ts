/** Build a deterministic, uncompressed ZIP package for unit tests. */
export function buildStoredZip(entries: Readonly<Record<string, string | Uint8Array>>): Uint8Array {
  return buildZipFixture(entries);
}

export interface ZipFixturePayload {
  /** ZIP compression method (0 = stored, 8 = raw DEFLATE). */
  readonly method: 0 | 8;
  readonly bytes: Uint8Array;
}

/** Build a deterministic ZIP while allowing tests to inject compressed payloads. */
export function buildZipFixture(
  entries: Readonly<Record<string, string | Uint8Array>>,
  encode: (name: string, bytes: Uint8Array) => ZipFixturePayload = (_name, bytes) => ({
    method: 0,
    bytes,
  }),
): Uint8Array {
  const encoded = Object.entries(entries).map(([name, value]) => ({
    name: new TextEncoder().encode(name),
    original: typeof value === 'string' ? new TextEncoder().encode(value) : value,
    payload: encode(name, typeof value === 'string' ? new TextEncoder().encode(value) : value),
  }));
  const localParts: Uint8Array[] = [];
  const centralParts: Uint8Array[] = [];
  let localOffset = 0;

  for (const entry of encoded) {
    const checksum = crc32(entry.original);
    const local = concatBytes([
      u32(0x04034b50),
      u16(20),
      u16(0x0800),
      u16(entry.payload.method),
      u16(0),
      u16(0),
      u32(checksum),
      u32(entry.payload.bytes.length),
      u32(entry.original.length),
      u16(entry.name.length),
      u16(0),
      entry.name,
      entry.payload.bytes,
    ]);
    localParts.push(local);

    centralParts.push(concatBytes([
      u32(0x02014b50),
      u16(20),
      u16(20),
      u16(0x0800),
      u16(entry.payload.method),
      u16(0),
      u16(0),
      u32(checksum),
      u32(entry.payload.bytes.length),
      u32(entry.original.length),
      u16(entry.name.length),
      u16(0),
      u16(0),
      u16(0),
      u16(0),
      u32(0),
      u32(localOffset),
      entry.name,
    ]));
    localOffset += local.length;
  }

  const central = concatBytes(centralParts);
  const end = concatBytes([
    u32(0x06054b50),
    u16(0),
    u16(0),
    u16(encoded.length),
    u16(encoded.length),
    u32(central.length),
    u32(localOffset),
    u16(0),
  ]);
  return concatBytes([...localParts, central, end]);
}

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) {
      crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function u16(value: number): Uint8Array {
  return Uint8Array.of(value & 0xff, (value >>> 8) & 0xff);
}

function u32(value: number): Uint8Array {
  return Uint8Array.of(
    value & 0xff,
    (value >>> 8) & 0xff,
    (value >>> 16) & 0xff,
    (value >>> 24) & 0xff,
  );
}

function concatBytes(parts: readonly Uint8Array[]): Uint8Array {
  const output = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
  let offset = 0;
  for (const part of parts) {
    output.set(part, offset);
    offset += part.length;
  }
  return output;
}

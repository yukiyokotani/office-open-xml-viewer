// Synthetic, redistributable Office binary fixtures for converter integration tests.
export function buildDocFixture(options: { text?: string; paragraphProperties?: Uint8Array; characterProperties?: Uint8Array; data?: Uint8Array; defaultTabTwips?: number; sectionProperties?: Uint8Array | readonly Uint8Array[]; sectionEnds?: readonly number[]; floatingAnchors?: Uint8Array; drawingGroupData?: Uint8Array; headers?: readonly string[]; footnotes?: string; facingPages?: boolean; lockedHeaderFields?: boolean } = {}): Uint8Array {
  const text = options.text ?? 'Hello 日本語\rSecond paragraph';
  const sectionEnds = options.sectionEnds ?? [text.length];
  if (options.headers && options.headers.length !== sectionEnds.length * 6) throw new Error('Expected six header/footer variants per section');
  const footnotes = options.footnotes ?? '';
  // Six empty separator stories, six per-section stories, and a final guard.
  const headerStories = options.headers?.map(t => t ? `${t}\r` : '') ?? [];
  const headerText = options.headers ? `${headerStories.join('')}\r` : '';
  const allText = text + footnotes + headerText;
  const units = Array.from({ length: allText.length }, (_, index) => allText.charCodeAt(index));
  const textOffset = 0x400;
  const word = new Uint8Array(textOffset + units.length * 2);
  const view = new DataView(word.buffer);
  view.setUint16(0, 0xa5ec, true);
  view.setUint16(2, 0x00c1, true);
  view.setUint32(0x4c, text.length, true);
  view.setUint32(0x50, footnotes.length, true);
  view.setUint32(0x54, headerText.length, true);
  view.setUint32(0x1a2, 0, true);
  units.forEach((unit, index) => view.setUint16(textOffset + index * 2, unit, true));
  const pieceProperties = concat(options.paragraphProperties ?? new Uint8Array(), options.characterProperties ?? new Uint8Array());
  const table = concat(
    ...(pieceProperties.length ? [new Uint8Array([1]), little16(pieceProperties.length), pieceProperties] : []),
    new Uint8Array([0x02]),
    little32(16),
    little32(0),
    little32(units.length),
    little16(0),
    little32(textOffset),
    little16(pieceProperties.length ? 1 : 0),
  );
  view.setUint32(0x1a6, table.length, true);
  const dop = options.defaultTabTwips === undefined ? new Uint8Array() : new Uint8Array(500);
  if (dop.length) {
    dop[0] = options.facingPages ? 1 : 0;
    // MS-DOC 2.7.2: DopBase.dxaTab at byte 10 of the Dop97 prefix.
    new DataView(dop.buffer).setUint16(10, options.defaultTabTwips as number, true);
    view.setUint32(0x192, table.length, true);
    view.setUint32(0x196, dop.length, true);
  }
  let sectionTable: Uint8Array = new Uint8Array();
  if (options.sectionProperties || options.headers || options.sectionEnds) {
    const supplied = options.sectionProperties ?? new Uint8Array();
    const shared = supplied instanceof Uint8Array;
    const properties = shared ? [supplied] : supplied;
    if (!shared && properties.length !== sectionEnds.length) throw new Error('Expected one Sepx per section');
    let offset = 0x300;
    const offsets = properties.map(bytes => {
      if (offset + 2 + bytes.length > textOffset) throw new Error('Synthetic Sepx overlaps text');
      const start = offset;
      view.setUint16(start, bytes.length, true);
      word.set(bytes, start + 2);
      offset += 2 + bytes.length;
      return start;
    });
    // PlcfSed: CP boundaries followed by 12-byte Seds, optionally sharing Sepx.
    sectionTable = concat(little32(0), ...sectionEnds.map(little32), ...sectionEnds.map((_, i) => concat(little16(0), little32(offsets[shared ? 0 : i]), new Uint8Array(6))));
    view.setUint32(0xca, table.length + dop.length, true);
    view.setUint32(0xce, sectionTable.length, true);
  }
  const floating = options.floatingAnchors ?? new Uint8Array();
  const drawing = options.drawingGroupData ?? new Uint8Array();
  if (floating.length) {
    view.setUint32(0x1da, table.length + dop.length + sectionTable.length, true);
    view.setUint32(0x1de, floating.length, true);
  }
  if (drawing.length) {
    view.setUint32(0x22a, table.length + dop.length + sectionTable.length + floating.length, true);
    view.setUint32(0x22e, drawing.length, true);
  }
  let headerTable: Uint8Array = new Uint8Array();
  if (options.headers) {
    const cps = [0, 0, 0, 0, 0, 0, 0];
    for (const story of headerStories) cps.push((cps.at(-1) as number) + story.length);
    cps.push(0xffffffff); // Undefined final PlcfHdd CP must be ignored.
    headerTable = concat(...cps.map(little32));
    view.setUint32(0xf2, table.length + dop.length + sectionTable.length + floating.length + drawing.length, true);
    view.setUint32(0xf6, headerTable.length, true);
  }
  const fieldCps: number[] = [], fieldRecords: Uint8Array[] = [];
  for (let cp = 0; cp < headerText.length; cp++) {
    const ch = headerText.charCodeAt(cp);
    if (ch >= 0x13 && ch <= 0x15) {
      fieldCps.push(cp);
      fieldRecords.push(new Uint8Array([ch, ch === 0x15 ? 0x80 | (options.lockedHeaderFields ? 0x10 : 0) : 0]));
    }
  }
  const fieldTable = fieldCps.length ? concat(...fieldCps.map(little32), little32(headerText.length + 2), ...fieldRecords) : new Uint8Array();
  if (fieldTable.length) {
    view.setUint32(0x122, table.length + dop.length + sectionTable.length + floating.length + drawing.length + headerTable.length, true);
    view.setUint32(0x126, fieldTable.length, true);
  }
  return buildCfb([
    ['WordDocument', word],
    ['0Table', concat(table, dop, sectionTable, floating, drawing, headerTable, fieldTable)],
    ...(options.data ? [['Data', options.data] as const] : []),
  ]);
}

export function buildXlsFixture(options: { sharedString?: Uint8Array; styleRecords?: Uint8Array } = {}): Uint8Array {
  const bof = (kind: number) => biffRecord(0x0809, concat(
    little16(0x0600), little16(kind), little16(0), little16(0),
  ));
  const sheetName = utf16le('表計算');
  const boundSheet = concat(
    little32(0), new Uint8Array([0, 0, 3, 1]), sheetName,
  );
  const string = utf16le('日本語');
  const sst = concat(
    little32(1), little32(1), options.sharedString ?? concat(little16(3), new Uint8Array([1]), string),
  );
  const globals = concat(
    bof(0x0005),
    biffRecord(0x0085, boundSheet),
    options.styleRecords ?? new Uint8Array(),
    biffRecord(0x00fc, sst),
    biffRecord(0x000a, new Uint8Array()),
  );
  const number = new Uint8Array(14);
  new DataView(number.buffer).setFloat64(6, 42.5, true);
  const label = concat(little16(1), little16(1), little16(0), little32(0));
  const sheet = concat(
    bof(0x0010),
    biffRecord(0x0203, number),
    biffRecord(0x00fd, label),
    biffRecord(0x000a, new Uint8Array()),
  );
  new DataView(boundSheet.buffer, boundSheet.byteOffset, boundSheet.byteLength)
    .setUint32(0, globals.length, true);
  return buildCfb([['Workbook', concat(
    bof(0x0005),
    biffRecord(0x0085, boundSheet),
    options.styleRecords ?? new Uint8Array(),
    biffRecord(0x00fc, sst),
    biffRecord(0x000a, new Uint8Array()),
    sheet,
  )]]);
}

export function buildPptFixture(slidePayload?: Uint8Array, outlinePayload: Uint8Array = new Uint8Array(), masterPayload?: Uint8Array, media?: { entries: Uint8Array[]; pictures?: Uint8Array }): Uint8Array {
  const record = (version: number, kind: number, payload: Uint8Array) => concat(
    little16(version), little16(kind), little32(payload.length), payload,
  );
  const documentAtom = concat(little32(5760), little32(4320), new Uint8Array(32));
  const slideReference = concat(little32(2), new Uint8Array(16));
  const document = record(0x000f, 1000, concat(
    record(1, 1001, documentAtom),
    record(15, 4080, concat(record(0, 1011, slideReference), outlinePayload)),
    ...(masterPayload ? [record(0x1f, 4080, record(0, 1011, concat(little32(3), new Uint8Array(8), little32(100), new Uint8Array(4))))] : []),
    ...(media ? [record(15, 1035, record(15, 0xf000, record((media.entries.length << 4) | 15, 0xf001, concat(...media.entries))))] : []),
  ));
  const slide = record(0x000f, 1006, slidePayload ?? record(0, 4000, utf16le('Legacy 日本語 slide')));
  const master = masterPayload ? record(15, 1016, masterPayload) : new Uint8Array();
  const directoryOffset = document.length + slide.length + master.length;
  const directory = record(0, 0x1772, concat(little32(masterPayload ? 0x00300001 : 0x00200001), little32(0), little32(document.length), ...(masterPayload ? [little32(document.length + slide.length)] : [])));
  const currentEdit = directoryOffset + directory.length;
  const userEdit = record(0, 0x0ff5, concat(new Uint8Array(12), little32(directoryOffset), little32(1), new Uint8Array(8)));
  const currentUserPayload = concat(
    little32(0x14),
    little32(0xe391c05f),
    little32(currentEdit),
    little16(0),
    little16(0x03f4),
    new Uint8Array([3, 0, 0, 0]),
    little32(8),
  );
  const currentUser = record(0, 0x0ff6, currentUserPayload);
  return buildCfb([
    ['PowerPoint Document', concat(document, slide, master, directory, userEdit)],
    ['Current User', currentUser],
    ...(media?.pictures ? [['Pictures', media.pictures] as const] : []),
  ]);
}

function biffRecord(kind: number, payload: Uint8Array): Uint8Array {
  return concat(little16(kind), little16(payload.length), payload);
}

export function little16(value: number): Uint8Array {
  const bytes = new Uint8Array(2);
  new DataView(bytes.buffer).setUint16(0, value, true);
  return bytes;
}

export function little32(value: number): Uint8Array {
  const bytes = new Uint8Array(4);
  new DataView(bytes.buffer).setUint32(0, value, true);
  return bytes;
}

export function utf16le(value: string): Uint8Array {
  const bytes = new Uint8Array(value.length * 2);
  const view = new DataView(bytes.buffer);
  for (let index = 0; index < value.length; index++) {
    view.setUint16(index * 2, value.charCodeAt(index), true);
  }
  return bytes;
}

export function concat(...parts: Uint8Array[]): Uint8Array {
  const output = new Uint8Array(parts.reduce((total, part) => total + part.length, 0));
  let offset = 0;
  for (const part of parts) {
    output.set(part, offset);
    offset += part.length;
  }
  return output;
}

function buildCfb(streams: ReadonlyArray<readonly [string, Uint8Array]>): Uint8Array {
  const sectorSize = 512;
  const padded = streams.map(([name, bytes]) => ({
    name,
    bytes,
    declared: Math.max(bytes.length, 4096),
    sectors: Math.ceil(Math.max(bytes.length, 4096) / sectorSize),
  }));
  const dataSectors = padded.reduce((total, entry) => total + entry.sectors, 0);
  const directorySectors = Math.max(1, Math.ceil((streams.length + 1) * 128 / sectorSize));
  let fatSectors = 1;
  while (dataSectors + directorySectors + fatSectors > fatSectors * 128) fatSectors++;
  const output = new Uint8Array(512 + (dataSectors + directorySectors + fatSectors) * sectorSize);
  const view = new DataView(output.buffer);
  output.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  view.setUint16(24, 0x003e, true);
  view.setUint16(26, 3, true);
  view.setUint16(28, 0xfffe, true);
  view.setUint16(30, 9, true);
  view.setUint16(32, 6, true);
  view.setUint32(44, fatSectors, true);
  view.setUint32(48, dataSectors, true);
  view.setUint32(56, 4096, true);
  view.setUint32(60, 0xfffffffe, true);
  view.setUint32(68, 0xfffffffe, true);
  for (let index = 0; index < 109; index++) {
    view.setUint32(76 + index * 4, index < fatSectors
      ? dataSectors + directorySectors + index
      : 0xffffffff, true);
  }

  let sector = 0;
  const starts: Array<{ start: number; size: number }> = [];
  for (const entry of padded) {
    starts.push({ start: sector, size: entry.declared });
    output.set(entry.bytes, 512 + sector * sectorSize);
    sector += entry.sectors;
  }
  const directoryStart = sector;
  const directoryOffset = 512 + directoryStart * sectorSize;
  writeDirectoryEntry(output.subarray(directoryOffset, directoryOffset + 128), 'Root Entry', 5, 0xfffffffe, 0);
  padded.forEach((entry, index) => {
    const offset = directoryOffset + (index + 1) * 128;
    writeDirectoryEntry(output.subarray(offset, offset + 128), entry.name, 2, starts[index].start, starts[index].size);
  });
  sector += directorySectors;
  const fatStart = sector;
  const fat = new Uint32Array(fatSectors * 128).fill(0xffffffff);
  let cursor = 0;
  for (const entry of padded) {
    for (let index = 0; index < entry.sectors; index++) {
      fat[cursor + index] = index + 1 === entry.sectors ? 0xfffffffe : cursor + index + 1;
    }
    cursor += entry.sectors;
  }
  for (let index = 0; index < directorySectors; index++) {
    fat[directoryStart + index] = index + 1 === directorySectors
      ? 0xfffffffe
      : directoryStart + index + 1;
  }
  for (let index = 0; index < fatSectors; index++) {
    fat[fatStart + index] = 0xfffffffd;
    const offset = 512 + (fatStart + index) * sectorSize;
    for (let entry = 0; entry < 128; entry++) {
      view.setUint32(offset + entry * 4, fat[index * 128 + entry], true);
    }
  }
  return output;
}

function writeDirectoryEntry(
  target: Uint8Array,
  name: string,
  objectType: number,
  startSector: number,
  size: number,
): void {
  const view = new DataView(target.buffer, target.byteOffset, target.byteLength);
  for (let index = 0; index < name.length; index++) {
    view.setUint16(index * 2, name.charCodeAt(index), true);
  }
  view.setUint16(name.length * 2, 0, true);
  view.setUint16(64, (name.length + 1) * 2, true);
  target[66] = objectType;
  target[67] = 1;
  view.setUint32(68, 0xffffffff, true);
  view.setUint32(72, 0xffffffff, true);
  view.setUint32(76, 0xffffffff, true);
  view.setUint32(116, startSector, true);
  view.setUint32(120, size, true);
}

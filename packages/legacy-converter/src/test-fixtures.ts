// Synthetic, redistributable Office binary fixtures for converter integration tests.
export interface BinaryNoteFixture { cp: number; text: string; automatic?: boolean }
export function buildDocFixture(options: { text?: string; paragraphProperties?: Uint8Array; characterProperties?: Uint8Array; formattingRuns?: readonly { end: number; properties: Uint8Array }[]; paragraphMarks?: readonly { end: number; properties: Uint8Array }[]; numbering?: { definitions: Uint8Array; definitionHeaderBytes: number; overrides: Uint8Array }; fonts?: readonly { name: string; charset: number }[]; data?: Uint8Array; defaultTabTwips?: number; sectionProperties?: Uint8Array | readonly Uint8Array[]; sectionEnds?: readonly number[]; floatingAnchors?: Uint8Array; drawingGroupData?: Uint8Array; headers?: readonly string[]; footnotes?: readonly BinaryNoteFixture[]; endnotes?: readonly BinaryNoteFixture[]; comments?: string; facingPages?: boolean; lockedHeaderFields?: boolean } = {}): Uint8Array {
  const text = options.text ?? 'Hello 日本語\rSecond paragraph';
  const sectionEnds = options.sectionEnds ?? [text.length];
  if (options.headers && options.headers.length !== sectionEnds.length * 6) throw new Error('Expected six header/footer variants per section');
  const noteText = (notes: readonly BinaryNoteFixture[] | undefined) => notes?.length ? notes.map(n => n.text).join('') + '\r' : '';
  const footnotes = noteText(options.footnotes), endnotes = noteText(options.endnotes);
  const comments = options.comments ?? '';
  const hasNotes = Boolean(options.footnotes?.length || options.endnotes?.length);
  // Notes need Word's standard separator stories (MS-DOC 2.3.3): separator and
  // continuation separator paragraphs, and an empty continuation notice.
  const separators = [...(options.footnotes?.length ? ['\u0003\r\r', '\u0004\r\r', ''] : ['', '', '']), ...(options.endnotes?.length ? ['\u0003\r\r', '\u0004\r\r', ''] : ['', '', ''])];
  const headers = options.headers ?? (hasNotes ? Array<string>(sectionEnds.length * 6).fill('') : undefined);
  // Six separator stories, six per-section stories, and a final guard.
  const headerStories = headers?.map(t => t ? `${t}\r` : '') ?? [];
  const headerText = headers ? `${separators.join('')}${headerStories.join('')}\r` : '';
  const allText = text + footnotes + headerText + comments + endnotes;
  const units = Array.from({ length: allText.length }, (_, index) => allText.charCodeAt(index));
  const textOffset = 0x400;
  const word = new Uint8Array(textOffset + units.length * 2);
  const view = new DataView(word.buffer);
  view.setUint16(0, 0xa5ec, true);
  view.setUint16(2, 0x00c1, true);
  // MS-DOC 2.5.1: the canonical Word 97 FIB counts csw = 0x000E,
  // cslw = 0x0016 and cbRgFcLcb = 0x005D precede the fixed field offsets the
  // readers use (cswNew at byte 898 stays 0, so nFib 0x00C1 applies).
  view.setUint16(32, 0x000e, true);
  view.setUint16(62, 0x0016, true);
  view.setUint16(152, 0x005d, true);
  view.setUint32(0x4c, text.length, true);
  view.setUint32(0x50, footnotes.length, true);
  view.setUint32(0x54, headerText.length, true);
  view.setUint32(0x5c, comments.length, true);
  view.setUint32(0x60, endnotes.length, true);
  view.setUint32(0x1a2, 0, true);
  units.forEach((unit, index) => view.setUint16(textOffset + index * 2, unit, true));
  const pieceProperties = concat(options.paragraphProperties ?? new Uint8Array(), options.characterProperties ?? new Uint8Array());
  let table = concat(
    ...(pieceProperties.length ? [new Uint8Array([1]), little16(pieceProperties.length), pieceProperties] : []),
    new Uint8Array([0x02]),
    little32(16),
    little32(0),
    little32(units.length),
    little16(0),
    little32(textOffset),
    little16(pieceProperties.length ? 1 : 0),
  );
  if (options.formattingRuns) {
    const runs = options.formattingRuns;
    const starts = [0, ...runs.slice(0, -1).map(r => r.end)];
    if (!runs.length || runs.at(-1)?.end !== units.length || runs.some((r, i) => r.end <= starts[i])) throw new Error('Invalid synthetic formatting piece boundaries');
    table = concat(
      ...runs.map(r => { const properties = concat(pieceProperties, r.properties); return concat(new Uint8Array([1]), little16(properties.length), properties); }),
      new Uint8Array([2]), little32(4 + runs.length * 12),
      little32(0), ...runs.map(r => little32(r.end)),
      ...runs.map((_, i) => concat(little16(0), little32(textOffset + starts[i] * 2), little16(i * 2 + 1))),
    );
  }
  view.setUint32(0x1a6, table.length, true);
  const dop = options.defaultTabTwips === undefined && !hasNotes ? new Uint8Array() : new Uint8Array(500);
  if (dop.length) {
    const dopView = new DataView(dop.buffer);
    dop[0] = options.facingPages ? 1 : 0;
    // MS-DOC 2.7.2: DopBase.dxaTab at byte 10 of the Dop97 prefix.
    dopView.setUint16(10, options.defaultTabTwips ?? 720, true);
    if (hasNotes) {
      // MS-DOC 2.7.2/2.7.4: bottom-of-page footnotes (fpc 1), continuous
      // numbering from 1, end-of-document endnotes (epc 3), Arabic footnote
      // and lowercase-Roman endnote numbers.
      dop[0] |= 1 << 5;
      dopView.setUint16(2, 1 << 2, true);
      dopView.setUint16(52, 1 << 2, true);
      dopView.setUint16(54, 3, true);
      dopView.setUint16(492, 0, true);
      dopView.setUint16(494, 2, true);
    }
    view.setUint32(0x192, table.length, true);
    view.setUint32(0x196, dop.length, true);
  }
  let sectionTable: Uint8Array = new Uint8Array();
  const sepxs: Uint8Array[] = [];
  {
    const supplied = options.sectionProperties ?? new Uint8Array();
    const shared = supplied instanceof Uint8Array;
    // Letter page, 1-inch margins, one column: the resolved geometry the
    // direct reader requires. Supplied sprms follow and override it.
    const geometry = concat(...([[0x9023, 1440], [0x9024, 1440], [0xb021, 1440], [0xb022, 1440], [0xb017, 720], [0xb018, 720], [0xb01f, 12240], [0xb020, 15840], [0x500b, 0], [0x900c, 720]] as const).map(([code, value]) => concat(little16(code), little16(value))));
    const properties = (shared ? [supplied] : supplied).map(bytes => concat(geometry, bytes));
    if (!shared && properties.length !== sectionEnds.length) throw new Error('Expected one Sepx per section');
    // Sepxs follow the text in the WordDocument stream (the FIB ends at 0x382).
    let offset = word.length;
    const offsets = properties.map(bytes => {
      const start = offset;
      sepxs.push(little16(bytes.length), bytes);
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
  if (headers) {
    const cps = [0];
    for (const story of separators) cps.push((cps.at(-1) as number) + story.length);
    for (const story of headerStories) cps.push((cps.at(-1) as number) + story.length);
    cps.push(0xffffffff); // Undefined final PlcfHdd CP must be ignored.
    headerTable = concat(...cps.map(little32));
    view.setUint32(0xf2, table.length + dop.length + sectionTable.length + floating.length + drawing.length, true);
    view.setUint32(0xf6, headerTable.length, true);
  }
  // MS-DOC 2.8.25 Plcfld per story (main, header, footnote, endnote): begin
  // records carry the flt of the instruction keyword (2.9.90), end records
  // grffldEnd fHasSep/fNested (2.9.110) and, for header fields, fLocked.
  const storyFieldTable = (story: string, locked: boolean): Uint8Array => {
    const cps: number[] = [], records: Uint8Array[] = [];
    const open: { begin: number; separated: boolean; parent: boolean }[] = [];
    for (let cp = 0; cp < story.length; cp++) {
      const ch = story.charCodeAt(cp);
      if (ch < 0x13 || ch > 0x15) continue;
      cps.push(cp);
      if (ch === 0x13) {
        const tail = story.slice(cp + 1);
        const keyword = tail.slice(0, tail.search(/[\u0013-\u0015]/)).trim().split(/\s+/)[0]?.toUpperCase() ?? '';
        records.push(new Uint8Array([ch, FIELD_TYPES[keyword] ?? 0]));
        open.push({ begin: cp, separated: false, parent: open.length > 0 });
      } else if (ch === 0x14) {
        records.push(new Uint8Array([ch, 0xff]));
        const top = open.at(-1);
        if (top) top.separated = true;
      } else {
        const field = open.pop();
        records.push(new Uint8Array([ch, (field?.separated ? 0x80 : 0) | (field?.parent ? 0x40 : 0) | (locked ? 0x10 : 0)]));
      }
    }
    return cps.length ? concat(...cps.map(little32), little32(story.length + 2), ...records) : new Uint8Array();
  };
  const fieldTables = ([[0x11a, text, false], [0x122, headerText, options.lockedHeaderFields === true], [0x12a, footnotes, false], [0x21a, endnotes, false]] as const)
    .map(([fibOffset, story, locked]) => [fibOffset, storyFieldTable(story, locked)] as const);
  let fieldOffset = table.length + dop.length + sectionTable.length + floating.length + drawing.length + headerTable.length;
  for (const [fibOffset, bytes] of fieldTables) {
    if (!bytes.length) continue;
    view.setUint32(fibOffset, fieldOffset, true);
    view.setUint32(fibOffset + 4, bytes.length, true);
    fieldOffset += bytes.length;
  }
  const fieldTable = concat(...fieldTables.map(([, bytes]) => bytes));
  const noteTables: Uint8Array[] = [];
  let noteOffset = table.length + dop.length + sectionTable.length + floating.length + drawing.length + headerTable.length + fieldTable.length;
  for (const [notes, offset] of [[options.footnotes, 0xaa], [options.endnotes, 0x20a]] as const) {
    if (!notes?.length) continue;
    const references = concat(...notes.map(n => little32(n.cp)), little32(0xffffffff), ...notes.map(n => little16(n.automatic === false ? 0 : 1)));
    const boundaries = [0];
    for (const note of notes) boundaries.push((boundaries.at(-1) as number) + note.text.length);
    const ranges = concat(...boundaries.map(little32), little32(0xffffffff));
    view.setUint32(offset, noteOffset, true); view.setUint32(offset + 4, references.length, true);
    noteOffset += references.length;
    view.setUint32(offset + 8, noteOffset, true); view.setUint32(offset + 12, ranges.length, true);
    noteOffset += ranges.length;
    noteTables.push(references, ranges);
  }
  const fonts = options.fonts?.length ? options.fonts : [{ name: 'Times New Roman', charset: 0 }];
  if (fonts.length > 0x7ff0) throw new Error('Too many synthetic fonts');
  const fontTable = concat(little16(fonts.length), little16(0), ...fonts.map(font => {
    if (!font.name || font.name.includes('\u0000')) throw new Error('Invalid synthetic font name');
    if (!Number.isInteger(font.charset) || font.charset < 0 || font.charset > 255) throw new Error('Invalid synthetic font charset');
    if (39 + (font.name.length + 1) * 2 > 255) throw new Error('Synthetic font name is too long');
    const name = concat(...Array.from({ length: font.name.length }, (_, index) => little16(font.name.charCodeAt(index))), little16(0));
    const record = new Uint8Array(39 + name.length);
    record[3] = font.charset;
    record.set(name, 39);
    return concat(new Uint8Array([record.length]), record);
  }));
  view.setUint32(0x112, noteOffset, true);
  view.setUint32(0x116, fontTable.length, true);
  noteOffset += fontTable.length;
  const numbering: Uint8Array[] = [];
  if (options.numbering) {
    const { definitions, definitionHeaderBytes, overrides } = options.numbering;
    view.setUint32(0x2e2, noteOffset, true);
    view.setUint32(0x2e6, definitionHeaderBytes, true); // Appended LVLs are not included.
    view.setUint32(0x2ea, noteOffset + definitions.length, true);
    view.setUint32(0x2ee, overrides.length, true);
    numbering.push(definitions, overrides);
    noteOffset += definitions.length + overrides.length;
  }
  // MS-DOC 2.9.271 STSH: an STSHI (cstd 15, cbSTDBaseInFile 10) with only the
  // Normal paragraph style (sti 0) defined, so every style reference resolves.
  const stshi = new Uint8Array(18);
  new DataView(stshi.buffer).setUint16(0, 15, true);
  new DataView(stshi.buffer).setUint16(2, 10, true);
  const normalStyle = new Uint8Array(14);
  new DataView(normalStyle.buffer).setUint16(2, 0xfff1, true);
  const stylesheet = concat(little16(18), stshi, little16(normalStyle.length), normalStyle, ...Array.from({ length: 14 }, () => little16(0)));
  view.setUint32(0xa2, noteOffset, true);
  view.setUint32(0xa6, stylesheet.length, true);
  noteOffset += stylesheet.length;
  // MS-DOC 2.8.5/2.8.6 PlcBteChpx/PlcBtePapx: one FKP page each covering the
  // whole text with no direct properties (piece properties still apply).
  const pages: Uint8Array[] = [];
  const binTables: Uint8Array[] = [];
  const textEnd = textOffset + units.length * 2;
  const bodyLength = word.length + sepxs.reduce((total, part) => total + part.length, 0);
  let wordLength = bodyLength + ((512 - bodyLength % 512) % 512);
  for (const fibOffset of [0xfa, 0x102]) {
    const page = fibOffset === 0x102 && options.paragraphMarks ? papxPage(options.paragraphMarks, textOffset, units.length) : new Uint8Array(512);
    if (!(fibOffset === 0x102 && options.paragraphMarks)) {
      new DataView(page.buffer).setUint32(0, textOffset, true);
      new DataView(page.buffer).setUint32(4, textEnd, true);
      page[511] = 1;
    }
    const bte = concat(little32(textOffset), little32(textEnd), little32(wordLength / 512));
    pages.push(page);
    wordLength += 512;
    view.setUint32(fibOffset, noteOffset, true);
    view.setUint32(fibOffset + 4, bte.length, true);
    noteOffset += bte.length;
    binTables.push(bte);
  }
  const wordStream = concat(word, ...sepxs, new Uint8Array((512 - bodyLength % 512) % 512), ...pages);
  return buildCfb([
    ['WordDocument', wordStream],
    ['0Table', concat(table, dop, sectionTable, floating, drawing, headerTable, fieldTable, ...noteTables, fontTable, ...numbering, stylesheet, ...binTables)],
    ...(options.data ? [['Data', options.data] as const] : []),
  ]);
}

/**
 * MS-DOC 2.9.175 PapxFkp: one run per paragraph mark, ending after the CP
 * `end` (exclusive); each run's GrpPrlAndIstd is `properties` (istd first).
 * A trailing run without properties covers any remaining text.
 */
function papxPage(marks: readonly { end: number; properties: Uint8Array }[], textOffset: number, length: number): Uint8Array {
  const runs = marks.at(-1)?.end === length ? marks : [...marks, { end: length, properties: new Uint8Array(2) }];
  if (runs.some((r, i) => r.end <= (i ? runs[i - 1].end : 0))) throw new Error('Invalid synthetic paragraph marks');
  const page = new Uint8Array(512);
  const view = new DataView(page.buffer);
  view.setUint32(0, textOffset, true);
  runs.forEach((r, i) => view.setUint32((i + 1) * 4, textOffset + r.end * 2, true));
  const bx = (runs.length + 1) * 4;
  let offset = 256;
  runs.forEach((r, i) => {
    const header = r.properties.length % 2 === 0 ? 2 : 1;
    if (offset + header + r.properties.length > 511) throw new Error('Synthetic PAPX FKP overflow');
    page[bx + i * 13] = offset / 2;
    page[offset + header - 1] = Math.ceil(r.properties.length / 2);
    page.set(r.properties, offset + header);
    offset = (offset + header + r.properties.length + 1) & ~1;
  });
  if (bx + runs.length * 13 > 256) throw new Error('Too many synthetic paragraph marks');
  page[511] = runs.length;
  return page;
}

// MS-DOC 2.9.90 flt values of the field keywords the synthetic fixtures use.
const FIELD_TYPES: Readonly<Record<string, number>> = { PAGE: 0x21, NUMPAGES: 0x1a, DATE: 0x1f, TIME: 0x20, DDE: 0x2d, INCLUDETEXT: 0x44 };

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
    // Every stream is a child of the root storage (a right-sibling chain),
    // as parent-scoped stream lookup requires.
    if (index + 1 < padded.length) view.setUint32(offset + 72, index + 2, true);
  });
  if (padded.length > 0) view.setUint32(directoryOffset + 76, 1, true);
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

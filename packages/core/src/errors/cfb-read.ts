/**
 * Compound File Binary (CFB / OLE2) stream reader — [MS-CFB].
 *
 * The sibling {@link import('./cfb-sniff').sniffCfb} only *classifies* a
 * container by enumerating directory-entry names; it deliberately never returns
 * stream contents. Decrypting a password-protected OOXML package needs the actual
 * bytes of two streams — `EncryptionInfo` (a small stream, so it lives in the
 * root entry's *mini stream*) and `EncryptedPackage` (a large stream in the
 * regular FAT). This module reads a named stream in full, implementing the
 * parts of [MS-CFB] the sniffer skipped:
 *
 *   - §2.5.1 DIFAT: the in-header DIFAT holds the first 109 FAT-sector
 *     locations; further FAT sectors are chained through DIFAT sectors. Both
 *     are followed here (the sniffer only needed the first 109).
 *   - §2.4 mini FAT / mini stream: streams smaller than the header's
 *     mini-stream cutoff (@0x38, normally 4096) are stored inside the root
 *     storage entry's *mini stream* as 64-byte mini sectors, chained through
 *     the mini FAT (first mini-FAT sector @0x3C). The mini stream itself is an
 *     ordinary FAT-chained stream whose start sector + size come from the root
 *     directory entry.
 *   - §2.6.1 directory-entry stream size (@0x78, LE64) and starting sector
 *     (@0x74).
 *
 * Same robustness contract as the sniffer: every read is bounds-checked, every
 * chain walk is bounded + cycle-guarded, and the function returns `null` for
 * any structural problem rather than throwing, hanging, or reading out of
 * range. It never allocates more than the declared stream size.
 */

const CFB_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

/** Any sector value >= MAXREGSECT (FREESECT / ENDOFCHAIN / FATSECT / DIFSECT)
 *  terminates a chain — see cfb-sniff.ts for the enumeration. */
const MAXREGSECT = 0xfffffffa;
const ENDOFCHAIN = 0xfffffffe;

const HEADER_SIZE = 512;
const DIR_ENTRY_SIZE = 128;

/** Hard caps so a hostile / cyclic structure can never make a walk unbounded.
 *  Real documents stay far below these. A 512 MiB package at 512-byte sectors
 *  is ~1M sectors, so the FAT-chain cap is generous but finite. */
const MAX_SECTOR_CHAIN = 4_000_000;
const MAX_MINI_CHAIN = 8_000_000;
const MAX_DIR_ENTRIES = 65_536;
const MAX_DIFAT_SECTORS = 1_000_000;

/** Parsed CFB header fields needed to walk streams. */
export interface CfbFatIndexHeader {
  sectorSize: number;
  numFatSectors: number;
  firstDifatSector: number;
  numDifatSectors: number;
}

interface CfbHeader extends CfbFatIndexHeader {
  miniSectorSize: number;
  miniStreamCutoff: number;
  firstDirSector: number;
  firstMiniFatSector: number;
}

/**
 * Read the full contents of a named stream from a CFB container.
 *
 * @returns the stream bytes (exactly its declared size), or `null` if the input
 *   is not a CFB, the stream is not found, or any structure is out of range /
 *   corrupt. Never throws.
 */
export function readCfbStream(bytes: Uint8Array, streamName: string): Uint8Array | null {
  if (bytes.length < HEADER_SIZE) return null;
  for (let i = 0; i < CFB_SIGNATURE.length; i++) {
    if (bytes[i] !== CFB_SIGNATURE[i]) return null;
  }

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const header = readHeader(view);
  if (header === null) return null;

  // Build the FAT-sector index (in-header DIFAT + DIFAT-sector extension).
  const fatSectors = collectCfbFatSectors(view, bytes.length, header);
  if (fatSectors === null) return null;

  // Locate the directory entry for the requested stream, plus the root entry
  // (needed for the mini stream). Both come from walking the directory chain.
  const dir = findDirectoryEntries(view, bytes.length, header, fatSectors, streamName);
  if (dir === null || dir.target === null) return null;

  const { target, root } = dir;

  // A stream smaller than the cutoff lives in the mini stream; otherwise in the
  // regular FAT. (Size 0 has no chain — return empty either way.)
  if (target.size === 0) return new Uint8Array(0);
  if (target.size < header.miniStreamCutoff) {
    if (root === null) return null; // no root => no mini stream
    return readMiniStream(view, bytes.length, header, fatSectors, root, target);
  }
  return readFatStream(view, bytes.length, header, fatSectors, target.startSector, target.size);
}

/** Parse the header fields relevant to stream reading. Rejects non-conformant
 *  sector shifts (see cfb-sniff.ts for why {9, 12} only). */
function readHeader(view: DataView): CfbHeader | null {
  const sectorShift = view.getUint16(0x1e, true);
  if (sectorShift !== 9 && sectorShift !== 12) return null;
  const miniSectorShift = view.getUint16(0x20, true);
  // [MS-CFB] §2.2 fixes the mini sector shift at 0x0006 (64-byte mini sectors).
  if (miniSectorShift !== 6) return null;
  const sectorSize = 1 << sectorShift;
  const miniSectorSize = 1 << miniSectorShift;
  const miniStreamCutoff = view.getUint32(0x38, true);
  return {
    sectorSize,
    numFatSectors: view.getUint32(0x2c, true),
    miniSectorSize,
    miniStreamCutoff,
    firstDirSector: view.getUint32(0x30, true),
    firstMiniFatSector: view.getUint32(0x3c, true),
    firstDifatSector: view.getUint32(0x44, true),
    numDifatSectors: view.getUint32(0x48, true),
  };
}

/** File byte offset of logical sector N (§2.2): (N + 1) * sectorSize. */
function sectorOffset(sector: number, sectorSize: number): number {
  return (sector + 1) * sectorSize;
}

function isRegularSector(sector: number): boolean {
  return sector >= 0 && sector <= MAXREGSECT;
}

/**
 * Collect the ordered list of FAT-sector locations: the 109 in-header DIFAT
 * entries (@0x4C) followed by any entries stored in chained DIFAT sectors
 * (§2.5.1). Each DIFAT sector holds (sectorSize/4 - 1) FAT locations plus a
 * trailing pointer to the next DIFAT sector.
 */
export function collectCfbFatSectors(
  view: DataView,
  totalLen: number,
  header: CfbFatIndexHeader,
): number[] | null {
  const { sectorSize } = header;
  const fatSectors: number[] = [];
  const seenFatSectors = new Set<number>();
  const physicalSectorCount = Math.max(0, Math.floor(totalLen / sectorSize) - 1);
  if (
    header.numFatSectors > physicalSectorCount
    || header.numDifatSectors > physicalSectorCount
    || header.numDifatSectors > MAX_DIFAT_SECTORS
  ) return null;
  const addFatSector = (location: number): boolean => {
    if (!isRegularSector(location)) return true;
    if (fatSectors.length >= header.numFatSectors) return true;
    if (location >= physicalSectorCount || seenFatSectors.has(location)) return false;
    seenFatSectors.add(location);
    fatSectors.push(location);
    return true;
  };

  // In-header DIFAT: 109 entries at 0x4C.
  for (let i = 0; i < 109; i++) {
    const loc = view.getUint32(0x4c + i * 4, true);
    if (!addFatSector(loc)) return null;
  }

  // DIFAT-sector extension chain.
  const entriesPerDifat = sectorSize / 4 - 1; // last slot is the next-DIFAT pointer
  let difatSector = header.firstDifatSector;
  const visited = new Set<number>();
  let steps = 0;
  while (isRegularSector(difatSector) && steps < header.numDifatSectors) {
    steps++;
    if (visited.has(difatSector)) break; // cycle guard
    visited.add(difatSector);

    const off = sectorOffset(difatSector, sectorSize);
    if (off < 0 || off + sectorSize > totalLen) return null;
    for (let i = 0; i < entriesPerDifat; i++) {
      const loc = view.getUint32(off + i * 4, true);
      if (!addFatSector(loc)) return null;
    }
    // Next DIFAT sector pointer is the last 4-byte slot.
    difatSector = view.getUint32(off + entriesPerDifat * 4, true);
  }
  if (
    steps !== header.numDifatSectors
    || isRegularSector(difatSector)
    || fatSectors.length !== header.numFatSectors
  ) return null;

  return fatSectors;
}

/** Read the next FAT entry for `sector` using the precomputed FAT-sector list. */
export function nextCfbFatSector(
  view: DataView,
  totalLen: number,
  sectorSize: number,
  fatSectors: number[],
  sector: number,
): number | null {
  const fatEntriesPerSector = sectorSize / 4;
  const fatIndex = Math.floor(sector / fatEntriesPerSector);
  const within = sector % fatEntriesPerSector;
  if (fatIndex >= fatSectors.length) return null;
  const fatSector = fatSectors[fatIndex];
  if (!isRegularSector(fatSector)) return null;
  const entryOff = sectorOffset(fatSector, sectorSize) + within * 4;
  if (entryOff < 0 || entryOff + 4 > totalLen) return null;
  return view.getUint32(entryOff, true);
}

interface DirEntry {
  startSector: number;
  size: number;
}

/**
 * Walk the directory chain and return the entry matching `streamName` plus the
 * root storage entry (object type 5). Directory entry name is UTF-16LE @0x00,
 * byte length @0x40; object type @0x42; starting sector @0x74; size (LE64)
 * @0x78.
 */
function findDirectoryEntries(
  view: DataView,
  totalLen: number,
  header: CfbHeader,
  fatSectors: number[],
  streamName: string,
): { target: DirEntry | null; root: DirEntry | null } | null {
  const { sectorSize } = header;
  const entriesPerSector = Math.floor(sectorSize / DIR_ENTRY_SIZE);
  if (entriesPerSector < 1) return null;

  let target: DirEntry | null = null;
  let root: DirEntry | null = null;

  const visited = new Set<number>();
  let sector = header.firstDirSector;
  let steps = 0;
  let scanned = 0;

  while (isRegularSector(sector)) {
    if (steps++ > MAX_SECTOR_CHAIN) return null;
    if (visited.has(sector)) break; // cycle guard
    visited.add(sector);

    const base = sectorOffset(sector, sectorSize);
    if (base < 0 || base + sectorSize > totalLen) return null;

    for (let i = 0; i < entriesPerSector; i++) {
      if (scanned++ > MAX_DIR_ENTRIES) return { target, root };
      const entryOff = base + i * DIR_ENTRY_SIZE;
      const objType = view.getUint8(entryOff + 0x42);
      if (objType === 0) continue; // unused entry
      const startSector = view.getUint32(entryOff + 0x74, true);
      // §2.6.1: for a v3 file the high 32 bits of the size MUST be zero; we read
      // the low 32 bits, which caps a single stream at 4 GiB — ample here.
      const size = view.getUint32(entryOff + 0x78, true);
      if (objType === 5) {
        root = { startSector, size }; // root storage: size = mini-stream length
        continue;
      }
      const name = readEntryName(view, entryOff);
      if (name === streamName) target = { startSector, size };
    }

    const next = nextCfbFatSector(view, totalLen, sectorSize, fatSectors, sector);
    if (next === null) break;
    sector = next;
  }

  return { target, root };
}

/** UTF-16LE name @0x00, byte length (incl. NUL) @0x40. */
function readEntryName(view: DataView, entryOff: number): string {
  const nameLenBytes = view.getUint16(entryOff + 0x40, true);
  if (nameLenBytes < 2 || nameLenBytes > 64) return '';
  const units = nameLenBytes / 2 - 1;
  let s = '';
  for (let i = 0; i < units; i++) {
    const code = view.getUint16(entryOff + i * 2, true);
    if (code === 0) break;
    s += String.fromCharCode(code);
  }
  return s;
}

/** Read a regular (FAT-chained) stream of `size` bytes starting at `start`. */
function readFatStream(
  view: DataView,
  totalLen: number,
  header: CfbHeader,
  fatSectors: number[],
  start: number,
  size: number,
): Uint8Array | null {
  const { sectorSize } = header;
  const out = new Uint8Array(size);
  let written = 0;
  let sector = start;
  const visited = new Set<number>();
  let steps = 0;

  while (isRegularSector(sector) && written < size) {
    if (steps++ > MAX_SECTOR_CHAIN) return null;
    if (visited.has(sector)) return null; // cycle => corrupt
    visited.add(sector);

    const base = sectorOffset(sector, sectorSize);
    if (base < 0 || base + sectorSize > totalLen) return null;
    const take = Math.min(sectorSize, size - written);
    out.set(new Uint8Array(view.buffer, view.byteOffset + base, take), written);
    written += take;

    const next = nextCfbFatSector(view, totalLen, sectorSize, fatSectors, sector);
    if (next === null) return null;
    sector = next;
  }

  return written === size ? out : null;
}

/**
 * Read a mini stream: the target's sectors are 64-byte *mini* sectors located
 * inside the root's mini stream (an ordinary FAT-chained stream) and chained
 * through the mini FAT. We first materialise the root mini stream, then walk
 * the mini-FAT chain to gather the target's mini sectors.
 */
function readMiniStream(
  view: DataView,
  totalLen: number,
  header: CfbHeader,
  fatSectors: number[],
  root: DirEntry,
  target: DirEntry,
): Uint8Array | null {
  const { sectorSize, miniSectorSize } = header;

  // Materialise the whole root mini stream from the regular FAT.
  const miniStream = readFatStream(view, totalLen, header, fatSectors, root.startSector, root.size);
  if (miniStream === null) return null;

  const out = new Uint8Array(target.size);
  let written = 0;
  let miniSector = target.startSector;
  const visited = new Set<number>();
  let steps = 0;
  const miniFatEntriesPerSector = sectorSize / 4;

  while (isRegularSector(miniSector) && written < target.size) {
    if (steps++ > MAX_MINI_CHAIN) return null;
    if (visited.has(miniSector)) return null; // cycle => corrupt
    visited.add(miniSector);

    const off = miniSector * miniSectorSize;
    if (off < 0 || off + miniSectorSize > miniStream.length) return null;
    const take = Math.min(miniSectorSize, target.size - written);
    out.set(miniStream.subarray(off, off + take), written);
    written += take;

    // Next mini sector via the mini FAT (chained through the regular FAT).
    const next = nextMiniFatEntry(
      view,
      totalLen,
      header,
      fatSectors,
      miniFatEntriesPerSector,
      miniSector,
    );
    if (next === null) return null;
    miniSector = next;
  }

  return written === target.size ? out : null;
}

/** Resolve the next mini sector by indexing the mini-FAT chain (which is itself
 *  a regular FAT-chained stream starting at header.firstMiniFatSector). */
function nextMiniFatEntry(
  view: DataView,
  totalLen: number,
  header: CfbHeader,
  fatSectors: number[],
  entriesPerSector: number,
  miniSector: number,
): number | null {
  const { sectorSize } = header;
  const miniFatSectorIndex = Math.floor(miniSector / entriesPerSector);
  const within = miniSector % entriesPerSector;

  // Walk the mini-FAT chain to the sector holding this entry.
  let fatSector = header.firstMiniFatSector;
  const visited = new Set<number>();
  for (let i = 0; i < miniFatSectorIndex; i++) {
    if (!isRegularSector(fatSector)) return null;
    if (visited.has(fatSector)) return null;
    visited.add(fatSector);
    const next = nextCfbFatSector(view, totalLen, sectorSize, fatSectors, fatSector);
    if (next === null) return null;
    fatSector = next;
  }
  if (!isRegularSector(fatSector)) return null;

  const entryOff = sectorOffset(fatSector, sectorSize) + within * 4;
  if (entryOff < 0 || entryOff + 4 > totalLen) return null;
  const v = view.getUint32(entryOff, true);
  return v === ENDOFCHAIN ? ENDOFCHAIN : v;
}

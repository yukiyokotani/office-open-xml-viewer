/**
 * Compound File Binary (CFB / OLE2) container sniffer — [MS-CFB].
 *
 * Password-protected Office documents and legacy binary formats (.doc / .xls /
 * .ppt) are not ZIP archives; they are CFB containers whose first 8 bytes are
 * the CFB signature `D0 CF 11 E0 A1 B1 1A E1`. Handing those to the ZIP-based
 * parser produces an opaque "zip open error". This pure function recognises the
 * container on the main thread so the `load()` factories can throw a typed
 * {@link import('./ooxml-error').OoxmlError} instead.
 *
 * Scope: only enough of [MS-CFB] to enumerate the directory-entry names —
 *
 *   - §2.2 header: signature, sector shift (2^SectorShift @ 0x1E), first
 *     directory sector location (@ 0x30), and DIFAT locations.
 *   - §2.3 FAT / §2.5.1 DIFAT: walk the directory-stream sector chain via the
 *     FAT, including DIFAT-sector extensions used by larger compound files. The
 *     mini FAT is irrelevant because the directory is a regular FAT stream.
 *   - §2.6 directory entries: 128 bytes each, name is UTF-16LE @ 0x00..0x40
 *     with the byte length @ 0x40.
 *
 * Every read is bounds-checked and the FAT walk is bounded + cycle-guarded, so
 * arbitrary / hostile bytes can only make it return early — never throw, hang,
 * or read out of range.
 */

import {
  collectCfbFatSectors,
  nextCfbFatSector,
  type CfbFatIndexHeader,
} from './cfb-read';

/** CFB header signature (§2.2). */
const CFB_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

// Special FAT sector values (§2.2 / §2.3): FREESECT (0xFFFFFFFF), ENDOFCHAIN
// (0xFFFFFFFE), FATSECT (0xFFFFFFFD), DIFSECT (0xFFFFFFFC). All four are FAT-
// internal markers that sit above MAXREGSECT, and the walk treats any value
// >= MAXREGSECT as "stop" — so none of them needs an individually named
// constant; a single upper-bound comparison in `isRegularSector` covers all.
const MAXREGSECT = 0xfffffffa;

const HEADER_SIZE = 512;
const DIR_ENTRY_SIZE = 128;
/** Max directory entries to scan — a hard cap so a huge / cyclic chain cannot
 *  make enumeration unbounded. Real documents have a handful. */
const MAX_DIR_ENTRIES = 4096;
/** Max sectors to walk in any FAT chain — bounds the loop independently of the
 *  visited-set cycle guard. */
const MAX_CHAIN_SECTORS = 8192;

export type CfbKind = 'encrypted' | 'legacy-binary-format' | 'cfb-unknown';
export type LegacyCfbFormat = 'doc' | 'xls' | 'ppt';

/** Directory-entry names that mark a legacy binary Office document. Compared
 *  case-sensitively — [MS-CFB] entry names are case-preserving and these are
 *  the canonical spellings the producers write. */
const LEGACY_STREAM_NAMES = new Set([
  'WordDocument', // .doc (§[MS-DOC])
  'Workbook', // .xls (§[MS-XLS], current)
  'Book', // .xls (older BIFF)
  'PowerPoint Document', // .ppt (§[MS-PPT])
]);

/** The stream that marks a password-protected / encrypted OOXML package
 *  (§[MS-OFFCRYPTO] §2.3.4). */
const ENCRYPTION_INFO_NAME = 'EncryptionInfo';

/**
 * Recognise a CFB container from its leading bytes and classify it.
 *
 * @returns
 *   - `'encrypted'` — a CFB with an `EncryptionInfo` stream (password-protected
 *     OOXML, or an encrypted legacy binary).
 *   - `'legacy-binary-format'` — a CFB with a `WordDocument` / `Workbook` /
 *     `Book` / `PowerPoint Document` stream (a raw .doc / .xls / .ppt).
 *   - `'cfb-unknown'` — a CFB whose directory could not be classified (no known
 *     stream, or a corrupt / out-of-range structure).
 *   - `null` — not a CFB at all (e.g. a ZIP-based .docx / .pptx / .xlsx).
 */
export function sniffCfb(bytes: Uint8Array): CfbKind | null {
  return inspectCfb(bytes)?.kind ?? null;
}

/**
 * Return an unambiguous legacy family when the CFB directory contains markers
 * for exactly one Office binary format. Multiple markers can legitimately
 * occur for embedded OLE objects, so ambiguous containers return `null` and
 * remain the converter's responsibility to validate.
 */
export function sniffLegacyOfficeFormat(bytes: Uint8Array): LegacyCfbFormat | null {
  const inspection = inspectCfb(bytes);
  if (inspection?.kind !== 'legacy-binary-format') return null;
  const formats = new Set<LegacyCfbFormat>();
  if (inspection.names.has('WordDocument')) formats.add('doc');
  if (inspection.names.has('Workbook') || inspection.names.has('Book')) formats.add('xls');
  if (inspection.names.has('PowerPoint Document')) formats.add('ppt');
  return formats.size === 1 ? [...formats][0] as LegacyCfbFormat : null;
}

function inspectCfb(bytes: Uint8Array): Readonly<{ kind: CfbKind; names: ReadonlySet<string> }> | null {
  // Not a CFB unless the full header signature matches.
  if (bytes.length < HEADER_SIZE) {
    // Still confirm the signature so a merely-short CFB is distinguishable from
    // a non-CFB: but without a full header we cannot walk anything, so a short
    // non-CFB is null and a short "CFB" is also null (we cannot prove it).
    return null;
  }
  for (let i = 0; i < CFB_SIGNATURE.length; i++) {
    if (bytes[i] !== CFB_SIGNATURE[i]) return null;
  }

  // From here we know it IS a CFB. Any structural problem => 'cfb-unknown'
  // (never null — null means "not a CFB", which we have already ruled out).
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

  const sectorShift = view.getUint16(0x1e, true);
  // [MS-CFB] §2.2 mandates exactly 0x0009 (512 B) for major version 3 and
  // 0x000C (4096 B) for major version 4 — no other value is a conformant CFB.
  // A wider "sane range" is not merely permissive: with shift 7-8 (sectorSize
  // 128/256), fileOffsetOfSector(N) = (N + 1) * sectorSize can land inside the
  // 512-byte header region instead of past it, so the FAT/directory walk would
  // silently misinterpret header bytes as sector data instead of failing
  // closed. Reject anything else as 'cfb-unknown'.
  if (sectorShift !== 9 && sectorShift !== 12) {
    return { kind: 'cfb-unknown', names: new Set() };
  }
  const sectorSize = 1 << sectorShift;

  const firstDirSector = view.getUint32(0x30, true);
  const fatHeader: CfbFatIndexHeader = {
    sectorSize,
    numFatSectors: view.getUint32(0x2c, true),
    firstDifatSector: view.getUint32(0x44, true),
    numDifatSectors: view.getUint32(0x48, true),
  };
  const fatSectors = collectCfbFatSectors(view, bytes.length, fatHeader);
  if (fatSectors === null) return { kind: 'cfb-unknown', names: new Set() };

  const names = enumerateDirectoryNames(
    view,
    bytes.length,
    sectorSize,
    firstDirSector,
    fatSectors,
  );
  if (names === null) return { kind: 'cfb-unknown', names: new Set() };

  // Encryption wins over a legacy marker: an encrypted .doc is still routed to
  // the crypto path, not the legacy dead-end.
  if (names.has(ENCRYPTION_INFO_NAME)) return { kind: 'encrypted', names };
  for (const n of names) {
    if (LEGACY_STREAM_NAMES.has(n)) return { kind: 'legacy-binary-format', names };
  }
  return { kind: 'cfb-unknown', names };
}

/**
 * Walk the directory-stream sector chain via the FAT and collect entry names.
 * Returns `null` if the structure is out of range / unreadable.
 */
function enumerateDirectoryNames(
  view: DataView,
  totalLen: number,
  sectorSize: number,
  firstDirSector: number,
  fatSectors: number[],
): Set<string> | null {
  if (!isRegularSector(firstDirSector)) return null;

  const names = new Set<string>();
  const entriesPerSector = Math.floor(sectorSize / DIR_ENTRY_SIZE);
  if (entriesPerSector < 1) return null;

  const visited = new Set<number>();
  let sector = firstDirSector;
  let steps = 0;
  let scanned = 0;

  while (isRegularSector(sector)) {
    if (steps++ > MAX_CHAIN_SECTORS) break; // hard loop bound
    if (visited.has(sector)) break; // cycle guard
    visited.add(sector);

    const sectorOffset = fileOffsetOfSector(sector, sectorSize);
    // The whole sector must lie within the buffer.
    if (sectorOffset < 0 || sectorOffset + sectorSize > totalLen) return null;

    for (let i = 0; i < entriesPerSector; i++) {
      if (scanned++ > MAX_DIR_ENTRIES) return names;
      const entryOff = sectorOffset + i * DIR_ENTRY_SIZE;
      const name = readEntryName(view, entryOff);
      if (name) names.add(name);
    }

    const next = nextCfbFatSector(view, totalLen, sectorSize, fatSectors, sector);
    if (next === null) break; // FAT entry unreadable -> stop with what we have
    sector = next;
  }

  return names;
}

/** Read a directory entry's name (§2.6.1): UTF-16LE at +0x00, byte length at
 *  +0x40 (includes the terminating NUL). Returns '' for an empty / invalid
 *  entry. */
function readEntryName(view: DataView, entryOff: number): string {
  // The 64-byte name field + the 2-byte length field must be in range; callers
  // already bound the whole sector, so a full 128-byte entry is in range.
  const nameLenBytes = view.getUint16(entryOff + 0x40, true);
  // Valid range is [0, 64]. 0 => unused entry. Length includes the NUL.
  if (nameLenBytes < 2 || nameLenBytes > 64) return '';
  const units = nameLenBytes / 2 - 1; // drop the terminator
  let s = '';
  for (let i = 0; i < units; i++) {
    const code = view.getUint16(entryOff + i * 2, true);
    if (code === 0) break;
    s += String.fromCharCode(code);
  }
  return s;
}

/** File byte offset of logical sector N: the 512-byte header occupies "sector
 *  -1", so data sector N starts at (N + 1) * sectorSize (§2.2). Assumes
 *  sectorSize >= 512 (guaranteed by the {9, 12} shift check above) so the
 *  512-byte header always fits within one sector and sector 0 cannot overlap
 *  it. */
function fileOffsetOfSector(sector: number, sectorSize: number): number {
  return (sector + 1) * sectorSize;
}

/** A regular (allocatable) sector number: 0 .. MAXREGSECT. The special values
 *  (FREESECT / ENDOFCHAIN / FATSECT / DIFSECT, see above) all sit above
 *  MAXREGSECT and are already excluded by the upper bound, so no separate
 *  check against them is needed. */
function isRegularSector(sector: number): boolean {
  return sector >= 0 && sector <= MAXREGSECT;
}

import { describe, it, expect } from 'vitest';
import { sniffCfb, sniffLegacyOfficeFormat } from './cfb-sniff';
import { buildCfbFixture } from '../testing/cfb-fixture';

/**
 * Minimal synthetic-CFB builder for the sniffer tests. Emits a version-3
 * (512-byte sector) compound file whose FAT lives in sector 0 and whose
 * directory chain starts at sector 1. Enough of [MS-CFB] to exercise the
 * header parse, the FAT-walked directory chain, and the directory-entry name
 * enumeration — not a general CFB writer.
 */
const SECTOR = 512;
const HEADER = 512;
const ENTRY = 128; // directory entry size
const FREESECT = 0xffffffff;
const ENDOFCHAIN = 0xfffffffe;
const FATSECT = 0xfffffffd;

const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

interface DirEntry {
  /** UTF-16 stream / storage name (without the trailing NUL the format stores). */
  name: string;
  /** 0 unknown, 1 storage, 2 stream, 5 root. Defaults to 2 (stream). */
  objType?: number;
}

interface BuildOpts {
  entries: DirEntry[];
  /** Sector size power-of-two shift. Default 9 (512). */
  sectorShift?: number;
  /** Major version. Default 3. */
  majorVersion?: number;
  /** Override first directory sector location (to point it out of range). */
  firstDirSector?: number;
}

/** Write a directory entry's 128 bytes into `view` at `off`. */
function writeEntry(view: DataView, off: number, e: DirEntry): void {
  const name = e.name;
  // Name: UTF-16LE, up to 32 code units (64 bytes) incl. terminator.
  const units = Math.min(name.length, 31);
  for (let i = 0; i < units; i++) {
    view.setUint16(off + i * 2, name.charCodeAt(i), true);
  }
  // Terminator NUL is already zero. Name length in BYTES incl. terminator.
  view.setUint16(off + 0x40, (units + 1) * 2, true);
  view.setUint8(off + 0x42, e.objType ?? 2); // object type
  view.setUint8(off + 0x43, 0); // color flag
  // Left/right/child sibling + CLSID left as 0 — the sniffer walks the
  // directory sequentially, not the red-black tree, so links are unused.
}

function buildCfb(opts: BuildOpts): Uint8Array {
  const sectorShift = opts.sectorShift ?? 9;
  const sectorSize = 1 << sectorShift;
  const major = opts.majorVersion ?? 3;
  const entries = opts.entries;

  // Directory chain sectors needed (starting at logical sector 1).
  const dirSectors = Math.max(1, Math.ceil(entries.length / (sectorSize / ENTRY)));
  // Layout: sector 0 = FAT, sectors 1..dirSectors = directory chain.
  const totalSectors = 1 + dirSectors;
  const size = HEADER + totalSectors * sectorSize;
  const buf = new ArrayBuffer(size);
  const view = new DataView(buf);
  const bytes = new Uint8Array(buf);

  // --- Header ---
  for (let i = 0; i < 8; i++) bytes[i] = SIGNATURE[i];
  view.setUint16(0x18, 0x003e, true); // minor version
  view.setUint16(0x1a, major, true); // major version
  view.setUint16(0x1c, 0xfffe, true); // byte order
  view.setUint16(0x1e, sectorShift, true); // sector shift
  view.setUint16(0x20, 6, true); // mini sector shift
  view.setUint32(0x28, major >= 4 ? dirSectors : 0, true); // number of directory sectors
  view.setUint32(0x2c, 1, true); // number of FAT sectors
  view.setUint32(0x30, opts.firstDirSector ?? 1, true); // first directory sector location
  view.setUint32(0x38, 0x00001000, true); // mini stream cutoff
  view.setUint32(0x3c, ENDOFCHAIN, true); // first mini FAT sector
  view.setUint32(0x40, 0, true); // number of mini FAT sectors
  view.setUint32(0x44, ENDOFCHAIN, true); // first DIFAT sector
  view.setUint32(0x48, 0, true); // number of DIFAT sectors
  // DIFAT[0] = FAT sector location (logical sector 0). Rest FREESECT.
  view.setUint32(0x4c, 0, true);
  for (let i = 1; i < 109; i++) view.setUint32(0x4c + i * 4, FREESECT, true);

  // --- FAT (logical sector 0) ---
  const fatOff = HEADER + 0 * sectorSize;
  const fatEntries = sectorSize / 4;
  for (let i = 0; i < fatEntries; i++) view.setUint32(fatOff + i * 4, FREESECT, true);
  view.setUint32(fatOff + 0 * 4, FATSECT, true); // sector 0 is the FAT
  // Directory chain 1 -> 2 -> ... -> ENDOFCHAIN.
  for (let s = 1; s <= dirSectors; s++) {
    const next = s < dirSectors ? s + 1 : ENDOFCHAIN;
    view.setUint32(fatOff + s * 4, next, true);
  }

  // --- Directory entries (logical sectors 1..dirSectors) ---
  for (let idx = 0; idx < entries.length; idx++) {
    const logicalSector = 1 + Math.floor(idx / (sectorSize / ENTRY));
    const within = idx % (sectorSize / ENTRY);
    const off = HEADER + logicalSector * sectorSize + within * ENTRY;
    writeEntry(view, off, entries[idx]);
  }

  return bytes;
}

/** A well-formed encrypted-OOXML CFB (has an EncryptionInfo stream). */
function encryptedCfb(): Uint8Array {
  return buildCfb({
    entries: [
      { name: 'Root Entry', objType: 5 },
      { name: 'EncryptionInfo', objType: 2 },
      { name: 'EncryptedPackage', objType: 2 },
    ],
  });
}

/**
 * Build the smallest v3 CFB whose directory FAT entry is addressed by the
 * first extension-DIFAT item rather than the 109 locations in the header.
 */
function extendedDifatCfb(directoryName: string): Uint8Array {
  const fatSectorCount = 110;
  const difatSector = fatSectorCount;
  const directorySector = 109 * (SECTOR / 4);
  const totalSectors = directorySector + 1;
  const bytes = new Uint8Array((totalSectors + 1) * SECTOR);
  const view = new DataView(bytes.buffer);
  bytes.set(SIGNATURE);
  view.setUint16(0x18, 0x003e, true);
  view.setUint16(0x1a, 3, true);
  view.setUint16(0x1c, 0xfffe, true);
  view.setUint16(0x1e, 9, true);
  view.setUint16(0x20, 6, true);
  view.setUint32(0x2c, fatSectorCount, true);
  view.setUint32(0x30, directorySector, true);
  view.setUint32(0x38, 0x1000, true);
  view.setUint32(0x3c, ENDOFCHAIN, true);
  view.setUint32(0x44, difatSector, true);
  view.setUint32(0x48, 1, true);

  for (let index = 0; index < 109; index++) {
    view.setUint32(0x4c + index * 4, index, true);
  }
  const difatOffset = HEADER + difatSector * SECTOR;
  view.setUint32(difatOffset, 109, true);
  for (let index = 1; index < SECTOR / 4 - 1; index++) {
    view.setUint32(difatOffset + index * 4, FREESECT, true);
  }
  view.setUint32(difatOffset + SECTOR - 4, ENDOFCHAIN, true);

  for (let fatSector = 0; fatSector < fatSectorCount; fatSector++) {
    const fatOffset = HEADER + fatSector * SECTOR;
    for (let index = 0; index < SECTOR / 4; index++) {
      view.setUint32(fatOffset + index * 4, FREESECT, true);
    }
  }
  const directoryFatOffset = HEADER + 109 * SECTOR;
  view.setUint32(directoryFatOffset, ENDOFCHAIN, true);
  writeEntry(view, HEADER + directorySector * SECTOR, {
    name: directoryName,
    objType: 2,
  });
  return bytes;
}

describe('sniffCfb — signature', () => {
  it('returns null for non-CFB bytes (ZIP magic)', () => {
    const zip = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0, 0, 0, 0, 0, 0]);
    expect(sniffCfb(zip)).toBeNull();
  });

  it('returns null for an empty / too-short buffer', () => {
    expect(sniffCfb(new Uint8Array(0))).toBeNull();
    expect(sniffCfb(new Uint8Array([0xd0, 0xcf, 0x11]))).toBeNull();
  });

  it('returns null when only the signature is present but header is truncated', () => {
    const short = new Uint8Array(16);
    short.set(SIGNATURE);
    expect(sniffCfb(short)).toBeNull();
  });
});

describe('sniffCfb — classification', () => {
  it('detects an encrypted OOXML container (EncryptionInfo stream)', () => {
    expect(sniffCfb(encryptedCfb())).toBe('encrypted');
  });

  it('detects a legacy Word binary (.doc, WordDocument stream)', () => {
    const cfb = buildCfb({
      entries: [
        { name: 'Root Entry', objType: 5 },
        { name: 'WordDocument', objType: 2 },
        { name: '1Table', objType: 2 },
      ],
    });
    expect(sniffCfb(cfb)).toBe('legacy-binary-format');
  });

  it('detects a legacy Excel binary (.xls, Workbook stream)', () => {
    const cfb = buildCfb({
      entries: [
        { name: 'Root Entry', objType: 5 },
        { name: 'Workbook', objType: 2 },
      ],
    });
    expect(sniffCfb(cfb)).toBe('legacy-binary-format');
  });

  it('detects the older Excel "Book" stream as legacy', () => {
    const cfb = buildCfb({
      entries: [
        { name: 'Root Entry', objType: 5 },
        { name: 'Book', objType: 2 },
      ],
    });
    expect(sniffCfb(cfb)).toBe('legacy-binary-format');
  });

  it('detects a legacy PowerPoint binary (.ppt, PowerPoint Document stream)', () => {
    const cfb = buildCfb({
      entries: [
        { name: 'Root Entry', objType: 5 },
        { name: 'PowerPoint Document', objType: 2 },
      ],
    });
    expect(sniffCfb(cfb)).toBe('legacy-binary-format');
  });

  it('prefers "encrypted" when both EncryptionInfo and a legacy stream appear', () => {
    // Encryption should win: an encrypted .doc is still handled by the crypto
    // path in the next PR, not the legacy-binary dead-end.
    const cfb = buildCfb({
      entries: [
        { name: 'Root Entry', objType: 5 },
        { name: 'WordDocument', objType: 2 },
        { name: 'EncryptionInfo', objType: 2 },
      ],
    });
    expect(sniffCfb(cfb)).toBe('encrypted');
  });

  it('returns "cfb-unknown" for a CFB with no recognised stream', () => {
    const cfb = buildCfb({
      entries: [
        { name: 'Root Entry', objType: 5 },
        { name: 'SomeOtherStream', objType: 2 },
      ],
    });
    expect(sniffCfb(cfb)).toBe('cfb-unknown');
  });

  it('spans a multi-sector directory chain to find a stream in the 2nd sector', () => {
    // 4 entries per 512-byte sector; put the marker as the 5th entry so it only
    // resolves if the FAT chain walk advances to the 2nd directory sector.
    const filler: DirEntry[] = [
      { name: 'Root Entry', objType: 5 },
      { name: 'a', objType: 2 },
      { name: 'b', objType: 2 },
      { name: 'c', objType: 2 },
      { name: 'EncryptionInfo', objType: 2 },
    ];
    expect(sniffCfb(buildCfb({ entries: filler }))).toBe('encrypted');
  });

  it('uses extension DIFAT sectors to classify a large legacy container', () => {
    expect(sniffCfb(extendedDifatCfb('WordDocument'))).toBe('legacy-binary-format');
    expect(sniffLegacyOfficeFormat(extendedDifatCfb('WordDocument'))).toBe('doc');
  });

  it('detects an encrypted container built as a major-version-4 (4096-byte sector) CFB', () => {
    // [MS-CFB] §2.2: major version 4 uses SectorShift 0x000C (4096-byte
    // sectors); the 512-byte header is followed by 3584 bytes of padding
    // before sector 0 begins. Uses the shared `buildCfbFixture` (also used by
    // the docx/pptx/xlsx load-guard suites) rather than this file's local
    // `buildCfb`, so the v4 layout is exercised the same way production
    // callers' fixtures are built.
    const cfb = buildCfbFixture(['Root Entry', 'EncryptionInfo', 'EncryptedPackage'], {
      majorVersion: 4,
    });
    expect(sniffCfb(new Uint8Array(cfb))).toBe('encrypted');
  });
});

describe('sniffLegacyOfficeFormat', () => {
  it.each([
    ['WordDocument', 'doc'],
    ['Workbook', 'xls'],
    ['Book', 'xls'],
    ['PowerPoint Document', 'ppt'],
  ] as const)('classifies an unambiguous %s marker as %s', (stream, format) => {
    expect(sniffLegacyOfficeFormat(new Uint8Array(buildCfbFixture(['Root Entry', stream])))).toBe(format);
  });

  it('does not guess when embedded-object markers make the family ambiguous', () => {
    expect(sniffLegacyOfficeFormat(new Uint8Array(buildCfbFixture([
      'Root Entry',
      'WordDocument',
      'Workbook',
    ])))).toBeNull();
  });
});

describe('sniffCfb — robustness (malicious / corrupt input)', () => {
  it('returns "cfb-unknown" when the sector shift is absurd (out-of-range read)', () => {
    // sectorShift 30 => 1 GiB sectors; the directory sector is far past EOF.
    const cfb = buildCfb({ entries: [{ name: 'Root Entry', objType: 5 }] });
    // Overwrite the sector shift in the finished buffer.
    new DataView(cfb.buffer).setUint16(0x1e, 30, true);
    expect(sniffCfb(cfb)).toBe('cfb-unknown');
  });

  it('returns "cfb-unknown" when the first directory sector points past EOF', () => {
    const cfb = buildCfb({
      entries: [{ name: 'Root Entry', objType: 5 }],
      firstDirSector: 9999,
    });
    expect(sniffCfb(cfb)).toBe('cfb-unknown');
  });

  it('rejects a FAT count larger than the file instead of allocating from it', () => {
    const cfb = encryptedCfb();
    new DataView(cfb.buffer).setUint32(0x2c, 0xffffffff, true);
    expect(sniffCfb(cfb)).toBe('cfb-unknown');
  });

  it('does not hang on a cyclic FAT directory chain (returns a value)', () => {
    // Build a normal encrypted CFB, then corrupt the FAT so sector 1 -> 1.
    const cfb = encryptedCfb();
    const view = new DataView(cfb.buffer);
    const fatOff = HEADER + 0 * SECTOR;
    view.setUint32(fatOff + 1 * 4, 1, true); // self-cycle at sector 1
    // Must terminate. The marker is in sector 1, so it is still found before
    // the cycle guard trips; the point is that it returns rather than looping.
    const result = sniffCfb(cfb);
    expect(['encrypted', 'cfb-unknown']).toContain(result);
  });

  it('does not hang when every directory sector points forward in a long cycle', () => {
    // A cycle that does NOT contain a recognised stream must still terminate.
    const cfb = buildCfb({
      entries: [
        { name: 'Root Entry', objType: 5 },
        { name: 'x', objType: 2 },
        { name: 'y', objType: 2 },
        { name: 'z', objType: 2 },
        { name: 'nothing', objType: 2 },
      ],
    });
    const view = new DataView(cfb.buffer);
    const fatOff = HEADER + 0 * SECTOR;
    // dir chain is sectors 1..2; make sector 2 loop back to 1.
    view.setUint32(fatOff + 2 * 4, 1, true);
    const result = sniffCfb(cfb);
    // No recognised stream anywhere -> cfb-unknown, and it must return.
    expect(result).toBe('cfb-unknown');
  });

  it('handles a signature-only buffer padded to header length as cfb-unknown or null', () => {
    const buf = new Uint8Array(HEADER);
    buf.set(SIGNATURE);
    // No valid sector shift / all-zero header. A zero sector shift => 1-byte
    // sectors; the walk should bail to cfb-unknown, never throw.
    const result = sniffCfb(buf);
    expect(result === 'cfb-unknown' || result === null).toBe(true);
  });
});

describe('sniffCfb — deterministic seeded fuzz', () => {
  /**
   * mulberry32: a tiny, fast, deterministic PRNG (public-domain algorithm).
   * Used instead of `Math.random()` so a failure is exactly reproducible from
   * the fixed `FUZZ_SEED` below — no flakiness, no need to capture the failing
   * input separately.
   */
  function mulberry32(seed: number): () => number {
    let a = seed >>> 0;
    return () => {
      a = (a + 0x6d2b79f5) | 0;
      let t = a;
      t = Math.imul(t ^ (t >>> 15), t | 1);
      t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  const FUZZ_SEED = 0xc5b_51ff; // arbitrary, fixed for reproducibility
  const FUZZ_ITERATIONS = 20_000;
  const VALID_CFB_KINDS: ReadonlySet<string> = new Set([
    'encrypted',
    'legacy-binary-format',
    'cfb-unknown',
  ]);

  function randomBytes(rand: () => number, length: number): Uint8Array {
    const bytes = new Uint8Array(length);
    for (let i = 0; i < length; i++) bytes[i] = Math.floor(rand() * 256);
    return bytes;
  }

  /** Category A: pure random bytes, random length (mostly not a CFB at all). */
  function genRandomBytes(rand: () => number): Uint8Array {
    const length = Math.floor(rand() * 4096); // 0 .. 4095 bytes
    return randomBytes(rand, length);
  }

  /**
   * Category B: the 8-byte CFB signature forced at offset 0, everything else
   * (including the rest of the header and an arbitrary-length random body)
   * random. Exercises the header-parse and FAT-walk paths on garbage that
   * nonetheless passes the initial signature check.
   */
  function genForcedSignature(rand: () => number): Uint8Array {
    const length = HEADER + Math.floor(rand() * 8192); // >= one full header
    const bytes = randomBytes(rand, length);
    for (let i = 0; i < SIGNATURE.length; i++) bytes[i] = SIGNATURE[i];
    return bytes;
  }

  /**
   * Category C: a structurally-valid CFB (via `buildCfb`) with `sectorShift`
   * and/or `firstDirSector` mutated to out-of-range / adversarial values —
   * targets the exact fields the sniffer validates defensively.
   */
  function genAnomalousFields(rand: () => number): Uint8Array {
    const cfb = buildCfb({
      entries: [
        { name: 'Root Entry', objType: 5 },
        { name: rand() < 0.5 ? 'EncryptionInfo' : 'WordDocument', objType: 2 },
      ],
    });
    const view = new DataView(cfb.buffer);
    if (rand() < 0.5) {
      // Anomalous sector shift: full 16-bit range, including the valid {9,
      // 12} values (so some fraction of this category still parses cleanly).
      view.setUint16(0x1e, Math.floor(rand() * 65536), true);
    }
    if (rand() < 0.5) {
      // Anomalous first directory sector: full 32-bit range, including
      // FREESECT/ENDOFCHAIN/FATSECT/DIFSECT and far-past-EOF values.
      view.setUint32(0x30, Math.floor(rand() * 0x1_0000_0000), true);
    }
    return cfb;
  }

  it(`never throws and always returns CfbKind | null over ${FUZZ_ITERATIONS} deterministic-seed inputs`, () => {
    const rand = mulberry32(FUZZ_SEED);
    for (let i = 0; i < FUZZ_ITERATIONS; i++) {
      const category = i % 3;
      const input =
        category === 0
          ? genRandomBytes(rand)
          : category === 1
            ? genForcedSignature(rand)
            : genAnomalousFields(rand);

      let result: ReturnType<typeof sniffCfb>;
      expect(() => {
        result = sniffCfb(input);
      }).not.toThrow();
      expect(result! === null || VALID_CFB_KINDS.has(result!)).toBe(true);
    }
  });
});

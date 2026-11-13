// Harness environment parity, not reader behaviour: the faces LibreOffice
// drew each reference PDF with, served to the renderer under the families
// the document requests, so both sides measure text with the same fonts.
//
// The evidence is the PDF itself (`pdffonts`, `pdftohtml -xml`):
// - a requested family this machine lacks, with a metric-compatible
//   replacement in LibreOffice's table (VCL.xcu) that the PDF uses, gets that
//   replacement's faces;
// - any other requested family this machine lacks gets the faces LibreOffice
//   fell back to: the PDF fonts no requested family explains, the one with
//   the most text per weight and style;
// - LibreOffice's bundled fonts the PDF uses are served under their own names.
import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';

const LO_FONTS = process.env.ODF_HARNESS_LO_FONTS ?? '/Applications/LibreOffice.app/Contents/Resources/fonts/truetype';
/** LibreOffice's metric-compatible replacements (VCL.xcu). */
const REPLACEMENTS = { calibri: 'Carlito', cambria: 'Caladea' };
/** Symbol fonts LibreOffice uses for bullets and fields, not as text fallback. */
const SYMBOL_FONTS = new Set(['symbol', 'wingdings', 'opensymbol', 'applecoloremoji']);

const key = (name) => name.normalize('NFKC').toLowerCase().replace(/[^\p{L}\p{N}]/gu, '');
const psBase = (ps) => ps.replace(/^[A-Z]{6}\+/u, '');

/** fontconfig records: family names, PostScript name, weight, slant, face index, file. */
function scan(args) {
  const out = execFileSync(args[0], [...args.slice(1), '--format', '%{family}|%{postscriptname}|%{weight}|%{slant}|%{index}|%{file}\n'], {
    encoding: 'utf8', maxBuffer: 64 << 20, stdio: ['ignore', 'pipe', 'ignore'],
  });
  return out.split('\n').filter(Boolean).map((line) => {
    const [families, ps, weight, slant, index, ...file] = line.split('|');
    return {
      families: families.split(',').map((f) => f.trim()), ps,
      weight: Number(weight) >= 150 ? '700' : '400', style: Number(slant) > 0 ? 'italic' : 'normal',
      index: Number(index), file: file.join('|'),
    };
  });
}

let catalog;
function fontCatalog() {
  if (catalog) return catalog;
  const installed = scan(['fc-list']);
  const bundled = existsSync(LO_FONTS)
    ? execFileSync('ls', [LO_FONTS], { encoding: 'utf8' }).split('\n').filter((f) => /\.(ttf|otf)$/u.test(f))
      .flatMap((f) => scan(['fc-scan', `${LO_FONTS}/${f}`]))
    : [];
  const byPs = new Map();
  for (const font of [...bundled, ...installed]) byPs.set(key(font.ps), font);
  const installedFamilies = new Set(installed.flatMap((font) => [...font.families.map(key), key(font.ps)]));
  catalog = { byPs, bundled, installedFamilies };
  return catalog;
}

/** One sfnt out of a TrueType collection (FontFace takes single fonts). */
function sfntAt(bytes, index) {
  if (bytes.toString('latin1', 0, 4) !== 'ttcf') return bytes;
  const offset = bytes.readUInt32BE(12 + 4 * index);
  const tables = bytes.readUInt16BE(offset + 4);
  const header = 12 + tables * 16;
  const records = [];
  let size = header;
  for (let i = 0; i < tables; i += 1) {
    const at = offset + 12 + i * 16;
    const length = bytes.readUInt32BE(at + 12);
    records.push({ at, from: bytes.readUInt32BE(at + 8), length, to: size });
    size += (length + 3) & ~3;
  }
  const out = Buffer.alloc(size);
  bytes.copy(out, 0, offset, offset + 12);
  for (const [i, record] of records.entries()) {
    bytes.copy(out, 12 + i * 16, record.at, record.at + 8);
    out.writeUInt32BE(record.to, 12 + i * 16 + 8);
    out.writeUInt32BE(record.length, 12 + i * 16 + 12);
    bytes.copy(out, record.to, record.from, record.from + record.length);
  }
  return out;
}

function requestedFamilies(odt) {
  let xml = '';
  for (const part of ['content.xml', 'styles.xml']) {
    try {
      xml += execFileSync('unzip', ['-p', odt, part], { encoding: 'utf8', maxBuffer: 256 << 20, stdio: ['ignore', 'pipe', 'ignore'] });
    } catch { /* a missing part declares nothing */ }
  }
  const families = new Set();
  for (const match of xml.matchAll(/svg:font-family="([^"]*)"/gu)) {
    const value = match[1].replace(/&apos;/gu, "'").replace(/&quot;/gu, '"').replace(/&amp;/gu, '&');
    for (const family of value.split(',')) {
      const name = family.trim().replace(/^['"]|['"]$/gu, '');
      if (name) families.add(name);
    }
  }
  return [...families];
}

/** Characters LibreOffice drew per PDF font on the compared pages. */
function textPerFont(pdf, pages) {
  const limit = Number.isFinite(pages) ? ['-l', String(pages)] : [];
  const xml = execFileSync('pdftohtml', ['-xml', '-i', '-q', '-stdout', ...limit, pdf], {
    encoding: 'utf8', maxBuffer: 256 << 20, stdio: ['ignore', 'pipe', 'ignore'],
  });
  const fontOf = new Map();
  for (const match of xml.matchAll(/<fontspec id="(\d+)"[^>]*family="([^"]*)"/gu)) fontOf.set(match[1], psBase(match[2]));
  const counts = new Map();
  for (const match of xml.matchAll(/<text [^>]*font="(\d+)"[^>]*>(.*?)<\/text>/gu)) {
    const ps = fontOf.get(match[1]);
    if (ps) counts.set(ps, (counts.get(ps) ?? 0) + match[2].replace(/<[^>]*>/gu, '').length);
  }
  return counts;
}

/**
 * Faces to serve for one corpus file: `{ family, weight, style, bytes }`.
 * `pages` limits the evidence to the compared pages.
 */
export function libreOfficeFaces(odt, pdf, pages = Infinity) {
  if (!existsSync(pdf)) return [];
  const { byPs, bundled, installedFamilies } = fontCatalog();
  const used = textPerFont(pdf, pages);
  const requested = requestedFamilies(odt);
  const faces = [];
  const seen = new Set();
  const serve = (family, font) => {
    const id = `${key(family)}|${font.weight}|${font.style}`;
    if (seen.has(id) || !existsSync(font.file)) return;
    seen.add(id);
    faces.push({ family, weight: font.weight, style: font.style, bytes: sfntAt(readFileSync(font.file), font.index) });
  };
  // Bundled LibreOffice fonts the PDF draws with, under their own families.
  for (const ps of used.keys()) {
    const font = byPs.get(key(ps));
    if (font && font.file.startsWith(LO_FONTS) && !installedFamilies.has(key(font.families[0]))) {
      for (const face of bundled.filter((f) => f.families[0] === font.families[0])) serve(font.families[0], face);
    }
  }
  const missing = requested.filter((family) => !installedFamilies.has(key(family)));
  const explained = (ps) => {
    const font = byPs.get(key(ps));
    return requested.some((family) => (font?.families ?? []).some((f) => key(f) === key(family))
      || key(ps).startsWith(key(family)) || key(ps).startsWith(REPLACEMENTS[key(family)]?.toLowerCase() ?? '\u0000'));
  };
  // Fallback faces: unexplained PDF fonts, the most text per weight and style.
  const fallback = new Map();
  for (const [ps, count] of used) {
    const font = byPs.get(key(ps));
    if (!font || SYMBOL_FONTS.has(key(font.families[0])) || explained(ps)) continue;
    const slot = `${font.weight}|${font.style}`;
    if ((fallback.get(slot)?.count ?? -1) < count) fallback.set(slot, { font, count });
  }
  const bundledFamily = (family) => bundled.filter((f) => key(f.families[0]) === key(family));
  for (const family of missing) {
    // LibreOffice draws a requested bundled family with its own files.
    const own = bundledFamily(family);
    if (own.length > 0) {
      for (const face of own) serve(family, face);
      continue;
    }
    const replacement = REPLACEMENTS[key(family)];
    // The replacement table is unconditional in LibreOffice.
    if (replacement) {
      for (const face of bundled.filter((f) => f.families[0] === replacement)) serve(family, face);
      continue;
    }
    for (const { font } of fallback.values()) serve(family, font);
  }
  return faces;
}

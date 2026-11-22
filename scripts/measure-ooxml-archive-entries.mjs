#!/usr/bin/env node
// Reports per-format archive entry counts (maximum, p95) for one or more
// corpus roots, without recording filenames or document content. Use it to
// re-check the `maxArchiveEntries` default in
// packages/ooxml-common/resource-policy.json: the default should leave large
// headroom over observed corpora and stay independent of (stricter than) the
// non-disableable hard ceiling in the same file.
//
// Usage: pnpm measure:archive-entries -- <corpus-root> [more-roots...]

import { readFile, readdir, stat } from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';

const OOXML_EXTENSIONS = new Set(['.docx', '.pptx', '.xlsx']);
const EOCD_SIGNATURE = 0x06054b50;
const ZIP64_EOCD_SIGNATURE = 0x06064b50;
const ZIP64_LOCATOR_SIGNATURE = 0x07064b50;
const MAX_EOCD_SEARCH = 22 + 0xffff;

export function readArchiveEntryCount(bytes) {
  const start = Math.max(0, bytes.length - MAX_EOCD_SEARCH);
  for (let offset = bytes.length - 22; offset >= start; offset -= 1) {
    if (bytes.readUInt32LE(offset) !== EOCD_SIGNATURE) continue;
    const commentLength = bytes.readUInt16LE(offset + 20);
    if (offset + 22 + commentLength !== bytes.length) continue;
    const count = bytes.readUInt16LE(offset + 10);
    if (count !== 0xffff) return count;
    const locator = offset - 20;
    if (locator < 0 || bytes.readUInt32LE(locator) !== ZIP64_LOCATOR_SIGNATURE) {
      throw new Error('ZIP64 end-of-central-directory locator is missing');
    }
    const zip64Offset = Number(bytes.readBigUInt64LE(locator + 8));
    if (!Number.isSafeInteger(zip64Offset)
      || zip64Offset + 40 > bytes.length
      || bytes.readUInt32LE(zip64Offset) !== ZIP64_EOCD_SIGNATURE) {
      throw new Error('ZIP64 end-of-central-directory record is invalid');
    }
    const zip64Count = Number(bytes.readBigUInt64LE(zip64Offset + 32));
    if (!Number.isSafeInteger(zip64Count)) {
      throw new Error('ZIP64 archive entry count exceeds the JavaScript safe-integer range');
    }
    return zip64Count;
  }
  throw new Error('ZIP end-of-central-directory record is missing');
}

async function collectFiles(input) {
  const info = await stat(input);
  if (info.isFile()) {
    const officeLockFile = path.basename(input).startsWith('~$');
    return !officeLockFile && OOXML_EXTENSIONS.has(path.extname(input).toLowerCase())
      ? [input]
      : [];
  }
  if (!info.isDirectory()) return [];
  const entries = await readdir(input, { withFileTypes: true });
  const nested = await Promise.all(entries.map((entry) =>
    collectFiles(path.join(input, entry.name))));
  return nested.flat();
}

function percentile(sorted, fraction) {
  if (sorted.length === 0) return 0;
  return sorted[Math.ceil(sorted.length * fraction) - 1];
}

async function main(inputs) {
  if (inputs.length === 0) {
    throw new Error('Usage: node scripts/measure-ooxml-archive-entries.mjs <file-or-directory> [...]');
  }
  const files = (await Promise.all(inputs.map(collectFiles))).flat();
  const groups = new Map();
  for (const file of files) {
    const format = path.extname(file).slice(1).toLowerCase();
    const values = groups.get(format) ?? [];
    values.push(readArchiveEntryCount(await readFile(file)));
    groups.set(format, values);
  }
  const formats = {};
  for (const [format, values] of [...groups].sort(([left], [right]) => left.localeCompare(right))) {
    values.sort((left, right) => left - right);
    formats[format] = {
      files: values.length,
      maximum: values.at(-1) ?? 0,
      p95: percentile(values, 0.95),
    };
  }
  process.stdout.write(`${JSON.stringify({ files: files.length, formats }, null, 2)}\n`);
}

if (process.argv[1]
  && path.resolve(process.argv[1]) === path.resolve(fileURLToPath(import.meta.url))) {
  main(process.argv.slice(2)).catch((error) => {
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
    process.exitCode = 1;
  });
}

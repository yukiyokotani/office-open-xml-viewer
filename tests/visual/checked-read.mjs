// Built-in-only helper shared by the self-VRT harness and the standalone corpus
// tools, so importing it never requires installed dependencies.
import {
  closeSync, constants, fstatSync, ftruncateSync, lstatSync, openSync, readFileSync, writeFileSync,
} from 'node:fs';

/** Read a whole regular file through one descriptor. The open is non-blocking
 * and refuses a final symlink, and the descriptor is checked with fstat before
 * any read, so a path that turned into a named pipe or device after an earlier
 * check cannot block the worker or be read. */
export function readRegularFile(path) {
  const fd = openSync(
    path,
    constants.O_RDONLY | constants.O_NONBLOCK | (constants.O_NOFOLLOW ?? 0),
  );
  try {
    if (!fstatSync(fd).isFile()) throw new Error(`self-VRT input is not a regular file: ${path}`);
    return readFileSync(fd);
  } finally {
    closeSync(fd);
  }
}

/** Fail unless `path` is absent or a regular file with no other hard link.
 * Only lstat is used, so callers can check every destination of a multi-file
 * operation before they write or remove anything. */
export function assertWritableLeaf(path) {
  let stat;
  try {
    stat = lstatSync(path);
  } catch (error) {
    if (error.code === 'ENOENT') return;
    throw error;
  }
  if (!stat.isFile()) {
    throw new Error(`refusing to write over a symlink or non-regular file: ${path}`);
  }
  if (stat.nlink > 1) throw new Error(`refusing to write over a hard-linked file: ${path}`);
}

/** Write a harness output file. An existing leaf must be a regular file with
 * no other hard link (a symlink or a named pipe there is refused before
 * anything is opened; a hard link would change the file it shares an inode
 * with). The open is non-blocking, refuses a final symlink and truncates
 * nothing: the descriptor is checked first. Callers check the directory
 * components separately. */
export function writeRegularFile(path, data) {
  assertWritableLeaf(path);
  const fd = openSync(
    path,
    constants.O_WRONLY | constants.O_CREAT | constants.O_NONBLOCK | (constants.O_NOFOLLOW ?? 0),
    0o666,
  );
  try {
    const opened = fstatSync(fd);
    if (!opened.isFile()) throw new Error(`self-VRT output is not a regular file: ${path}`);
    if (opened.nlink > 1) throw new Error(`refusing to write over a hard-linked file: ${path}`);
    ftruncateSync(fd, 0);
    writeFileSync(fd, data);
  } finally {
    closeSync(fd);
  }
}

/** The one rule for which names in a corpus folder are inputs: the format's
 * extension, and not an Office lock file (`~$...`). */
export function isCorpusInput(name, format) {
  return name.endsWith(`.${format}`) && !name.startsWith('~$');
}

const CORPUS_INPUT_PATH = /^packages\/(docx|xlsx|pptx)\/public\/corpus\/([^/]+)$/;

/** True for the repository path of a public specimen: an input directly in
 * its own package's corpus folder. */
export function isCorpusInputPath(path) {
  const match = CORPUS_INPUT_PATH.exec(path);
  return match !== null && isCorpusInput(match[2], match[1]);
}

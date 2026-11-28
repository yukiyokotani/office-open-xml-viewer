// Built-in-only helper shared by the self-VRT harness and the standalone corpus
// tools, so importing it never requires installed dependencies.
import {
  closeSync, constants, fstatSync, lstatSync, openSync, readFileSync, writeFileSync,
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

/** Write a harness output file. An existing leaf must be a regular file (a
 * symlink or a named pipe there is refused before anything is opened), and the
 * open is non-blocking and refuses a final symlink. Callers check the
 * directory components separately. */
export function writeRegularFile(path, data) {
  let stat;
  try {
    stat = lstatSync(path);
  } catch (error) {
    if (error.code !== 'ENOENT') throw error;
  }
  if (stat && !stat.isFile()) {
    throw new Error(`refusing to write over a symlink or non-regular file: ${path}`);
  }
  const fd = openSync(
    path,
    constants.O_WRONLY | constants.O_CREAT | constants.O_TRUNC | constants.O_NONBLOCK
      | (constants.O_NOFOLLOW ?? 0),
    0o666,
  );
  try {
    if (!fstatSync(fd).isFile()) throw new Error(`self-VRT output is not a regular file: ${path}`);
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

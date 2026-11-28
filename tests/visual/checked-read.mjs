// Built-in-only helper shared by the self-VRT harness and the standalone corpus
// tools, so importing it never requires installed dependencies.
import { closeSync, constants, fstatSync, openSync, readFileSync } from 'node:fs';

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

import { sniffCfb } from './cfb-sniff';
import { OoxmlError } from './ooxml-error';
import { decryptOoxml } from '../crypto/decrypt-ooxml';

/**
 * Main-thread guard shared by the docx / pptx / xlsx `load()` factories.
 *
 * Given the raw file bytes (before they are handed to the ZIP-based parser
 * worker), throw a typed {@link OoxmlError} if the bytes are a CFB (OLE2)
 * container rather than an OOXML ZIP:
 *
 *   - password-protected OOXML          -> code `'encrypted'`
 *   - legacy binary .doc / .xls / .ppt  -> code `'legacy-binary-format'`
 *   - any other / corrupt CFB           -> code `'not-ooxml'`
 *
 * A ZIP-based OOXML file (or anything that is not a CFB) passes through
 * silently. Detection is done here, on the main thread, so a real `OoxmlError`
 * instance reaches the caller — `instanceof` would not survive the worker's
 * structured-clone boundary.
 *
 * This is the *no-password* path. To decrypt an encrypted file instead of
 * rejecting it, call {@link resolveOoxmlContainer} with a password.
 */
export function assertNotCfbContainer(bytes: Uint8Array | ArrayBuffer): void {
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const kind = sniffCfb(u8);
  if (kind === null) return; // not a CFB — a normal ZIP-based OOXML file

  switch (kind) {
    case 'encrypted':
      throw new OoxmlError(
        'encrypted',
        'This file is password-protected (MS-OFFCRYPTO). Pass LoadOptions.password to decrypt it.',
      );
    case 'legacy-binary-format':
      throw new OoxmlError(
        'legacy-binary-format',
        'This is a legacy binary Office file (.doc/.xls/.ppt), not OOXML.',
      );
    case 'cfb-unknown':
      throw new OoxmlError(
        'not-ooxml',
        'This file is an OLE2/Compound File container, not an OOXML (ZIP) document.',
      );
    default:
      // Exhaustiveness guard: if `CfbKind` grows a new member, this branch
      // stops compiling (the assignment requires `never`) instead of letting
      // an unrecognised CFB fall through and reach the ZIP parser worker
      // un-rejected. Fail closed at runtime too, in case a future change
      // drops the compile-time check (e.g. a widened return type upstream).
      kind satisfies never;
      throw new OoxmlError(
        'not-ooxml',
        'This file is an OLE2/Compound File container of an unrecognised kind, not an OOXML (ZIP) document.',
      );
  }
}

/**
 * Resolve raw input bytes into OOXML ZIP bytes ready for the parser, decrypting
 * an Agile-encrypted container when a `password` is supplied.
 *
 * This is the decrypt-aware superset of {@link assertNotCfbContainer} that the
 * three `load()` factories call. Behaviour:
 *
 *   - **Not a CFB** (a normal ZIP) → returns the bytes unchanged.
 *   - **Encrypted CFB, `password` supplied** → decrypts (Agile / [MS-OFFCRYPTO])
 *     and returns the plaintext ZIP bytes.
 *       - wrong password           → throws `OoxmlError('invalid-password')`
 *       - non-Agile scheme          → throws `OoxmlError('unsupported-encryption')`
 *       - unreadable / corrupt CFB → throws `OoxmlError('not-ooxml')`
 *   - **Encrypted CFB, no `password`** → throws `OoxmlError('encrypted')`.
 *   - **Legacy / unknown CFB** → same typed errors as `assertNotCfbContainer`.
 *
 * Async because WebCrypto is asynchronous; the returned `Uint8Array` is the
 * transferable payload the caller hands to the worker.
 */
export async function resolveOoxmlContainer(
  bytes: Uint8Array | ArrayBuffer,
  password?: string,
): Promise<Uint8Array> {
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const kind = sniffCfb(u8);
  if (kind === null) return u8; // normal ZIP-based OOXML file

  if (kind === 'encrypted') {
    if (password === undefined) {
      throw new OoxmlError(
        'encrypted',
        'This file is password-protected (MS-OFFCRYPTO). Pass LoadOptions.password to decrypt it.',
      );
    }
    const result = await decryptOoxml(u8, password);
    if (result.ok) return result.data;
    switch (result.reason) {
      case 'invalid-password':
        throw new OoxmlError('invalid-password', 'The supplied password is incorrect.');
      case 'unsupported-encryption':
        throw new OoxmlError(
          'unsupported-encryption',
          'This file uses an encryption scheme other than Agile ([MS-OFFCRYPTO]) that is not supported ' +
            '(Standard / Extensible / legacy binary encryption).',
        );
      case 'corrupt':
        throw new OoxmlError(
          'not-ooxml',
          'This file is an encrypted OLE2/Compound File container but its structure could not be read.',
        );
      default:
        result.reason satisfies never;
        throw new OoxmlError('not-ooxml', 'This encrypted file could not be decrypted.');
    }
  }

  // Not encrypted (legacy / unknown) — reuse the same typed errors.
  assertNotCfbContainer(u8);
  return u8; // unreachable: assertNotCfbContainer throws for every non-null kind
}

/**
 * Return a standalone `ArrayBuffer` holding exactly the bytes of `u8`, suitable
 * for transferring to a parser worker. Slices out a sub-view (or a
 * `SharedArrayBuffer`) rather than exposing an over-long / non-transferable
 * backing buffer. When `u8` already owns its whole `ArrayBuffer`, that buffer is
 * returned directly (no copy).
 */
export function toArrayBuffer(u8: Uint8Array): ArrayBuffer {
  if (u8.byteOffset === 0 && u8.byteLength === u8.buffer.byteLength && u8.buffer instanceof ArrayBuffer) {
    return u8.buffer;
  }
  // Do not call the instance `slice()`: Node.js Buffer subclasses Uint8Array
  // but overrides `slice()` to return another view over the original pooled
  // backing store. Copy explicitly so the result contains exactly this view.
  const copy = new Uint8Array(u8.byteLength);
  copy.set(u8);
  return copy.buffer;
}

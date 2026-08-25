/** ECMA-376 §17.13.5 tracked-changes (`<w:ins>` / `<w:del>`) paint policy.
 * Word's "All Markup" view underlines insertions and strikes through deletions
 * in the author's color. The palette mirrors Word's fixed author-color set;
 * the author→slot assignment is renderer-defined (Word's own assignment is
 * per-installation, not document-authored), so this module pins a deterministic
 * one. */

export const WORD_TRACK_CHANGE_AUTHOR_COLORS = Object.freeze([
  '#C00000',
  '#0070C0',
  '#00B050',
  '#7030A0',
  '#E97132',
  '#196B24',
  '#9E480E',
  '#525252',
] as const);

const NO_TRACK_CHANGE_DECORATION = Object.freeze({
  underline: false,
  strike: false,
});
const INSERTION_TRACK_CHANGE_DECORATION = Object.freeze({
  underline: true,
  strike: false,
});
const DELETION_TRACK_CHANGE_DECORATION = Object.freeze({
  underline: false,
  strike: true,
});

export function wordTrackChangeDecoration(
  kind: string | null | undefined,
): Readonly<{ underline: boolean; strike: boolean }> {
  if (kind === 'insertion') return INSERTION_TRACK_CHANGE_DECORATION;
  if (kind === 'deletion') return DELETION_TRACK_CHANGE_DECORATION;
  return NO_TRACK_CHANGE_DECORATION;
}

/** Deterministic author-index policy: a stable FNV-1a hash of the author name
 * with a finalizer mix selects one of the eight palette slots, so a given
 * author keeps the same color across runs, pages, sessions, and documents
 * regardless of layout order. An absent or empty author takes slot 0. Once
 * more than eight authors appear, slots repeat — the same fixed-size
 * author-color behaviour Word itself exhibits. */
export function wordTrackChangeAuthorColor(
  author: string | null | undefined,
): string {
  if (!author) return WORD_TRACK_CHANGE_AUTHOR_COLORS[0];
  let hash = 0x811c9dc5;
  for (let index = 0; index < author.length; index += 1) {
    hash = Math.imul(hash ^ author.charCodeAt(index), 0x01000193);
  }
  hash = Math.imul(hash ^ (hash >>> 16), 0x45d9f3b);
  hash = Math.imul(hash ^ (hash >>> 16), 0x45d9f3b);
  hash ^= hash >>> 16;
  return WORD_TRACK_CHANGE_AUTHOR_COLORS[(hash >>> 0) % WORD_TRACK_CHANGE_AUTHOR_COLORS.length];
}

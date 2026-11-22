/** Serialize a named family as one CSS string (CSS Syntax 3 §4.3.5/4.3.7).
 * OOXML ST_TextTypeface is a string, so quotes and backslashes are data, not
 * font-list syntax. Generic family keywords are handled by callers separately.
 * Hex escapes for controls include a terminator so a following digit is data.
 */
export function quoteCssFontFamily(family: string): string {
  let escaped = '';
  for (const scalar of family) {
    const codePoint = scalar.codePointAt(0) ?? 0;
    if (scalar === '"' || scalar === '\\') escaped += `\\${scalar}`;
    else if (codePoint === 0) escaped += '\uFFFD';
    else if (codePoint <= 0x1f || codePoint === 0x7f) escaped += `\\${codePoint.toString(16)} `;
    else escaped += scalar;
  }
  return `"${escaped}"`;
}

const GENERIC_FAMILIES = new Set([
  'serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui',
  'ui-serif', 'ui-sans-serif', 'ui-monospace', 'ui-rounded', 'math', 'emoji', 'fangsong',
]);

export interface FamilyEntry {
  /** Source text of the entry, as written in the list. */
  readonly text: string;
  /** Family name (quotes and escapes removed). */
  readonly name: string;
  readonly generic: boolean;
}

/** Split a CSS font-family list into entries, honouring quoted strings and
 * escapes (CSS Syntax 3 §4.3.5). Unquoted multi-word names collapse their
 * whitespace (CSS Fonts 4 §2.1). Returns null for malformed lists. */
export function parseFamilyList(list: string): FamilyEntry[] | null {
  const entries: FamilyEntry[] = [];
  let index = 0;
  while (index < list.length) {
    while (index < list.length && /\s/.test(list[index]!)) index += 1;
    const start = index;
    let name = '';
    let quoted = false;
    if (list[index] === '"' || list[index] === "'") {
      quoted = true;
      const quote = list[index]!;
      index += 1;
      let closed = false;
      while (index < list.length) {
        const character = list[index]!;
        if (character === '\\') {
          const hex = /^[0-9a-fA-F]{1,6}\s?/.exec(list.slice(index + 1));
          if (hex) {
            name += String.fromCodePoint(Math.min(parseInt(hex[0].trim(), 16) || 0xfffd, 0x10ffff));
            index += 1 + hex[0].length;
          } else {
            name += list[index + 1] ?? '';
            index += 2;
          }
          continue;
        }
        if (character === quote) { closed = true; index += 1; break; }
        name += character;
        index += 1;
      }
      if (!closed) return null;
      while (index < list.length && /\s/.test(list[index]!)) index += 1;
    } else {
      while (index < list.length && list[index] !== ',') index += 1;
      name = list.slice(start, index).trim().replace(/\s+/g, ' ');
    }
    const text = list.slice(start, index).trim();
    if (!name || (index < list.length && list[index] !== ',')) return null;
    entries.push({ text, name, generic: !quoted && GENERIC_FAMILIES.has(name.toLowerCase()) });
    index += 1;
  }
  return entries.length > 0 ? entries : null;
}

/** Per family map, the lists already scoped (lists repeat on every run). */
const scopedLists = new WeakMap<ReadonlyMap<string, string>, Map<string, string>>();
const SCOPED_LIST_MEMO_LIMIT = 1024;

/**
 * `list` (a CSS font-family list) with a document's served families scoped:
 * each named family the map holds (keyed by its lower-cased, trimmed name)
 * is preceded by the private family its faces are registered under
 * (`LoadedFontSources.families` in `fonts/source-session.ts`). The served
 * faces then win exactly where faces registered under the real name would,
 * and the real name stays as the next entry, so a gap in the served faces'
 * `unicode-range`s falls to that family on the host. Generic keywords and
 * unmapped families are kept as written; an unparseable list is returned
 * unchanged. Idempotent: an entry already preceded by its alias is kept.
 */
export function scopeServedFontFamilies(list: string, families: ReadonlyMap<string, string> | undefined): string {
  if (!families?.size) return list;
  let memo = scopedLists.get(families);
  if (!memo) { memo = new Map(); scopedLists.set(families, memo); }
  const retained = memo.get(list);
  if (retained !== undefined) return retained;
  const entries = parseFamilyList(list);
  let scoped = list;
  if (entries) {
    const out: string[] = [];
    entries.forEach((entry, index) => {
      const alias = entry.generic ? undefined : families.get(entry.name.trim().toLocaleLowerCase('en-US'));
      if (alias !== undefined && entries[index - 1]?.name !== alias) out.push(quoteCssFontFamily(alias));
      out.push(entry.text);
    });
    if (out.length !== entries.length) scoped = out.join(', ');
  }
  if (memo.size >= SCOPED_LIST_MEMO_LIMIT) memo.clear();
  memo.set(list, scoped);
  return scoped;
}

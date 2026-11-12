// CSS font-family lists, parsed as CSS reads them (a small module shared by
// native-route metrics and the symbol-face oracle).

export const GENERIC_FAMILIES = new Set([
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
  // A CSS escape at list[index] (a backslash): the escaped code point and
  // the index after it (CSS Syntax 3 §4.3.7).
  const escape = (): string => {
    const hex = /^[0-9a-fA-F]{1,6}\s?/.exec(list.slice(index + 1));
    if (hex) {
      index += 1 + hex[0].length;
      const code = parseInt(hex[0].trim(), 16);
      return String.fromCodePoint(code === 0 || code > 0x10ffff || (code >= 0xd800 && code <= 0xdfff) ? 0xfffd : code);
    }
    const next = list[index + 1] ?? '';
    index += 2;
    return next;
  };
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
        if (character === '\\') { name += escape(); continue; }
        if (character === quote) { closed = true; index += 1; break; }
        name += character;
        index += 1;
      }
      if (!closed) return null;
      while (index < list.length && /\s/.test(list[index]!)) index += 1;
    } else {
      // Identifiers (escapes decoded) joined by single spaces (§2.1).
      let space = false;
      while (index < list.length && list[index] !== ',') {
        const character = list[index]!;
        if (/\s/.test(character)) { space = name.length > 0; index += 1; continue; }
        if (space) { name += ' '; space = false; }
        if (character === '\\') { name += escape(); continue; }
        name += character;
        index += 1;
      }
    }
    const text = list.slice(start, index).trim();
    if (!name || (index < list.length && list[index] !== ',')) return null;
    entries.push({ text, name, generic: !quoted && GENERIC_FAMILIES.has(name.toLowerCase()) });
    index += 1;
  }
  return entries.length > 0 ? entries : null;
}

import { describe, expect, it } from 'vitest';
import { createFontResolver } from './font-service.js';
import { createTextLayoutService } from './text.js';

// Face names are document text (w:rFonts, fontTable w:font@w:name). Names that
// are Object.prototype members must behave exactly like any other face name:
// an absent entry must not resolve to an inherited value, and an own entry
// (as JSON.parse and structured clone produce for `__proto__`) must be found.
const PROTO_NAMES = ['__proto__', 'constructor', 'toString', 'hasOwnProperty'];
const ORDINARY = 'ordinary face';

/** `{ [name]: value }` with `name` as an own data property, as JSON.parse
 * and structured clone deliver it from the parser or across a worker. */
const own = <T>(name: string, value: T): Record<string, T> =>
  JSON.parse(`{${JSON.stringify(name)}:${JSON.stringify(value)}}`) as Record<string, T>;

const measurer = {
  fingerprint: 'prototype-key-names-v1',
  measure: (request: { text: string }) => ({ advancePt: request.text.length, ascentPt: 1, descentPt: 0 }),
};

describe('DOCX font tables keyed by face names that are prototype members', () => {
  it.each(PROTO_NAMES)('shapes text whose east Asian face is %s like an ordinary face', (name) => {
    const shape = (family: string, charsets: Record<string, string>) => createTextLayoutService({
      fonts: createFontResolver([]), eastAsiaFontCharsets: charsets, measurer,
    }).shape({
      text: 'Ā', fontSizePt: 10, fontHint: 'eastAsia',
      fonts: { highAnsi: 'Latin', eastAsia: family },
    }).spans.map((span) => span.script);
    // Absent charset entry (previously `.trim is not a function` for __proto__).
    expect(shape(name, {})).toEqual(shape(ORDINARY, {}));
    // Own charset entry: the GB2312 charset routes U+0100 to eastAsia.
    expect(shape(name, own(name, '86'))).toEqual(shape(ORDINARY, own(ORDINARY, '86')));
    expect(shape(name, own(name, '86'))).toEqual(['eastAsia']);
  });

  it.each(PROTO_NAMES)('resolves a requested family named %s through own family lists only', (name) => {
    const absent = createFontResolver([]).resolve({ requestedFamily: name });
    expect(absent.route.familyList.startsWith(JSON.stringify(name))).toBe(true);
    const listed = createFontResolver([], { nativeFamilyLists: own(name, '"Listed Face", serif') })
      .resolve({ requestedFamily: name });
    expect(listed.route.familyList).toBe('"Listed Face", serif');
  });
});

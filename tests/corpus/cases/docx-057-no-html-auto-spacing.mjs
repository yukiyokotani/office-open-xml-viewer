import { DOCX } from '../lib/grid.mjs';
import { docxPackage, header } from '../lib/docx.mjs';
import { p, r } from '../lib/wml.mjs';

// Adjacent paragraph spacing with w:doNotUseHTMLParagraphAutoSpacing set in
// w:compat: the space after one paragraph and the space before the next are
// both kept (they add up). Each case is two paragraphs between markers.
//   A. After 24 pt, then before 12 pt: a 36 pt gap.
//   B. After 12 pt, then before 24 pt.
//   C. After 12 pt, then before 12 pt.
//   D. Three paragraphs with automatic spacing (beforeAutospacing and
//      afterAutospacing) under the same compatibility option.
const { LINE } = DOCX;
const raw = (xml) => ({ kind: 'raw', height: NaN, xml });
const sp = (before, after) => `<w:spacing${before ? ` w:before="${before}"` : ''}${after ? ` w:after="${after}"` : ''}/>`;
const marker = (text) => raw(p('', r(text)));
const caption = (text, extra = {}) => ({ kind: 'p', style: 'caption', bold: true, lines: [text], before: 1, height: NaN, ...extra });
const pair = (name, after, before) => [
  caption(`${name}. After ${after / 20} pt, then before ${before / 20} pt`),
  marker(`▼ ${name} starts`),
  raw(p(sp(0, after), r(`${name}1 has ${after / 20} pt after.`))),
  raw(p(sp(before, 0), r(`${name}2 has ${before / 20} pt before.`))),
  marker(`▲ ${name} ends`),
];
const AUTO = '<w:spacing w:beforeAutospacing="1" w:afterAutospacing="1"/>';

export default {
  id: 'docx-057',
  slug: 'no-html-auto-spacing',
  format: 'docx',
  title: 'Spacing without HTML rules',
  covers: [
    'word/settings.xml|w:compat>w:doNotUseHTMLParagraphAutoSpacing',
    'word/document.xml|w:spacing@w:before',
    'word/document.xml|w:spacing@w:after',
    'word/document.xml|w:spacing@w:beforeAutospacing',
  ],
  parts: () => {
    const pkg = docxPackage([
      ...header({
        id: 'docx-057',
        title: 'Spacing without HTML rules',
        verify: [
          'Verify: with the compatibility option set, the space',
          'after and the space before add up (A: 36 pt); D shows',
          'automatic spacing under the same option.',
        ],
      }),
      { kind: 'p', style: 'caption', lines: ['Each case is two paragraphs between markers.'], at: DOCX.rowLine(2) },
      ...pair('A', 2 * LINE, LINE),
      ...pair('B', LINE, 2 * LINE),
      ...pair('C', LINE, LINE),
      caption('D. Automatic spacing'),
      marker('▼ D starts'),
      ...['D1', 'D2', 'D3'].map((n) => raw(p(AUTO, r(`${n} with automatic space before and after.`)))),
      marker('▲ D ends'),
    ]);
    pkg['word/settings.xml'] = pkg['word/settings.xml'].replace('<w:compat>', '<w:compat><w:doNotUseHTMLParagraphAutoSpacing/>');
    return pkg;
  },
};

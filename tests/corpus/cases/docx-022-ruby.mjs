import { DOCX, FONT, LANG } from '../lib/grid.mjs';
import { docxPackage, header, run } from '../lib/docx.mjs';
import { esc } from '../lib/xml.mjs';

// Ruby (furigana): one line per w:rubyAlign value, each with a base longer
// than its ruby, a base shorter than its ruby, and an equal pair. Bases are
// 18 pt, ruby 9 pt raised 18 pt, on lines of three baselines. A last
// paragraph puts ruby into 9 pt body text on single line spacing, where the
// line grows to make room.
const BASE_SZ = 36; // half-points
const RUBY_SZ = 18;
const RAISE = 36;
const JA = `<w:rFonts w:ascii="${FONT.eastAsia}" w:eastAsia="${FONT.eastAsia}" w:hAnsi="${FONT.eastAsia}"/>`;
const ruby = (base, text, align, baseSz = BASE_SZ, rubySz = RUBY_SZ, raise = RAISE) => '<w:r><w:ruby><w:rubyPr>'
  + `<w:rubyAlign w:val="${align}"/><w:hps w:val="${rubySz}"/><w:hpsRaise w:val="${raise}"/><w:hpsBaseText w:val="${baseSz}"/>`
  + `<w:lid w:val="${LANG.eastAsia}"/></w:rubyPr>`
  + `<w:rt><w:r><w:rPr>${JA}<w:sz w:val="${rubySz}"/></w:rPr><w:t>${esc(text)}</w:t></w:r></w:rt>`
  + `<w:rubyBase><w:r><w:rPr>${JA}<w:sz w:val="${baseSz}"/></w:rPr><w:t>${esc(base)}</w:t></w:r></w:rubyBase></w:ruby></w:r>`;
const space = run('　', `${JA}<w:sz w:val="${BASE_SZ}"/>`);
const WORDS = [['紫陽花', 'あじさい'], ['峠', 'とうげ'], ['五月雨', 'さみだれ']];
const ALIGNS = ['center', 'distributeLetter', 'distributeSpace', 'left', 'right', 'rightVertical'];

const RUBY_STYLE = '<w:style w:type="paragraph" w:styleId="RubyLine"><w:name w:val="Ruby line"/><w:basedOn w:val="Normal"/>'
  + `<w:pPr><w:spacing w:line="${3 * DOCX.LINE}" w:lineRule="exact"/></w:pPr></w:style>`;

export default {
  id: 'docx-022',
  slug: 'ruby',
  format: 'docx',
  title: 'Ruby',
  covers: [
    'word/document.xml|w:r>w:ruby',
    'word/document.xml|w:ruby>w:rubyPr',
    'word/document.xml|w:rubyAlign@w:val=center',
    'word/document.xml|w:rubyAlign@w:val=distributeLetter',
    'word/document.xml|w:rubyAlign@w:val=distributeSpace',
    'word/document.xml|w:rubyAlign@w:val=left',
    'word/document.xml|w:rubyAlign@w:val=right',
    'word/document.xml|w:rubyAlign@w:val=rightVertical',
    'word/document.xml|w:rubyPr>w:hps',
    'word/document.xml|w:rubyPr>w:hpsRaise',
    'word/document.xml|w:rubyPr>w:hpsBaseText',
    'word/document.xml|w:ruby>w:rt',
    'word/document.xml|w:ruby>w:rubyBase',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-022',
      title: 'Ruby',
      verify: [
        'Verify: six lines, one per alignment, each 紫陽花 あじさい,',
        '峠 とうげ, 五月雨 さみだれ in 18 pt with 9 pt ruby above;',
        'then 9 pt body text whose ruby lines grow to fit.',
      ],
    }),
    ...ALIGNS.map((align, i) => ({
      kind: 'p',
      before: i ? 0 : 1,
      pStyle: 'RubyLine',
      tabs: [['left', DOCX.colStart(2)]],
      runs: [run(`rubyAlign ${align}\t`, `<w:sz w:val="${DOCX.TYPE.caption.sz}"/>`), ...WORDS.flatMap(([b, r], j) => [...(j ? [space] : []), ruby(b, r, align)])],
      height: 3,
    })),
    {
      kind: 'p',
      before: 1,
      lineAuto: true,
      runs: [
        run('本文の中のルビ。', JA),
        ruby('雪', 'ゆき', 'center', 18, 9, 18),
        run('が降ると、', JA),
        ruby('山', 'やま', 'center', 18, 9, 18),
        run('の', JA),
        ruby('稜線', 'りょうせん', 'distributeSpace', 18, 9, 18),
        run('は白く見える。ルビのある行は、単一行間でルビの高さだけ広がる。', JA),
      ],
    },
  ], { styles: RUBY_STYLE }),
};

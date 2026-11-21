import { DOCX } from '../lib/grid.mjs';
import { docxPackage, header } from '../lib/docx.mjs';

const at = DOCX.rowLine;
const { colStart: start, colEnd: end, GUTTER } = DOCX;

export default {
  id: 'docx-002',
  slug: 'tab-stops-indents',
  format: 'docx',
  title: 'Tab stops and indents',
  covers: [
    'word/document.xml|w:tabs>w:tab',
    'word/document.xml|w:tab@w:val=left',
    'word/document.xml|w:tab@w:val=right',
    'word/document.xml|w:tab@w:val=decimal',
    'word/document.xml|w:tab@w:pos',
    'word/document.xml|w:r>w:tab',
    'word/document.xml|w:ind@w:left',
    'word/document.xml|w:ind@w:hanging',
    'word/document.xml|w:ind@w:firstLine',
    'word/document.xml|w:ind@w:right',
  ],
  parts: () => docxPackage([
    ...header({
      id: 'docx-002',
      title: 'Tab stops and indents',
      verify: [
        'Verify: tab stops and indents fall on the column',
        `edges at 0, ${start(2)}, ${start(3)} and ${start(4)} twips (0, 42.5,`,
        '85 and 127.5 mm from the left margin).',
      ],
    }),

    { kind: 'p', at: at(2), style: 'caption', bold: true, lines: [`Left tab stops at ${start(2)}, ${start(3)} and ${start(4)} twips, the left edges of columns 2 to 4`] },
    {
      kind: 'p',
      tabs: [['left', start(2)], ['left', start(3)], ['left', start(4)]],
      lines: [`Margin 0\tLeft ${start(2)}\tLeft ${start(3)}\tLeft ${start(4)}`],
    },
    { kind: 'p', before: 1, style: 'caption', bold: true, lines: [`Right tab stops at ${end(2)} and ${end(4)} twips, the right edges of columns 2 and 4`] },
    {
      kind: 'p',
      tabs: [['right', end(2)], ['right', end(4)]],
      lines: [`Margin 0\tRight ${end(2)}\tRight ${end(4)}`],
    },

    { kind: 'p', at: at(3), style: 'caption', bold: true, lines: [`Decimal tab stops at ${start(3)} and ${start(4)} twips: the decimal sign sits on the column edge`] },
    {
      kind: 'p',
      bold: true,
      tabs: [['decimal', start(3)], ['decimal', start(4)]],
      lines: ['Unit\tMillimetres\tPoints'],
    },
    {
      kind: 'p',
      tabs: [['decimal', start(3)], ['decimal', start(4)]],
      lines: [
        '1 inch\t25.4\t72',
        '1 pica\t4.2333\t12',
        '1 point\t0.3528\t1',
        '1 millimetre\t1\t2.8346',
        '1 twip\t0.0176\t0.05',
        '1 Q\t0.25\t0.7087',
      ],
    },

    { kind: 'p', at: at(4), style: 'caption', bold: true, lines: [`Left, hanging and first-line indents of ${start(2)} and ${GUTTER} twips`] },
    {
      kind: 'p',
      cols: [2, 4],
      lines: [
        `Left indent ${start(2)} twips, 42.5 mm: both lines of this paragraph`,
        'begin on the left edge of column 2.',
      ],
    },
    {
      kind: 'p',
      before: 1,
      ind: { 'w:left': start(2), 'w:hanging': start(2) },
      lines: [
        `Hanging ${start(2)}\tThe first line begins at 0; the tab after the label reaches`,
        `the indent at ${start(2)}, where the second line also begins.`,
      ],
    },
    {
      kind: 'p',
      before: 1,
      cols: [1, 3],
      ind: { 'w:firstLine': GUTTER },
      lines: [
        `First-line indent ${GUTTER} twips, one 5 mm gutter: only the first line`,
        'of this paragraph is indented.',
      ],
    },

    { kind: 'p', at: at(5), style: 'caption', bold: true, lines: [`Right indent ${DOCX.LIVE_W - end(2)} twips: the measure ends on the right edge of column 2`] },
    {
      kind: 'p',
      cols: [1, 2],
      lines: [
        `Right indent ${DOCX.LIVE_W - end(2)} twips, 85 mm from the right edge of the live area. `
          + 'The lines of this paragraph therefore wrap at 80 mm, the width of two columns '
          + 'and one gutter. Where they break is left to the application: this is the only '
          + 'paragraph in the file whose line breaks are not written into the text.',
      ],
    },
  ]),
};

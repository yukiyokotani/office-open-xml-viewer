import { bench, describe } from 'vitest';
import { layoutBodyModel } from './test-support/document-layout.test-support.js';
import { installStubCanvas } from './testing/synthetic-document.js';
import type { BodyElement, DocParagraph, DocxTextRun, SectionProps } from './types';

/**
 * #1986: layout of a document whose runs carry negative `w:spacing` (the
 * #1975 condensed floor). Run with
 * `pnpm exec vitest bench --run packages/docx/src/condensed-floor.bench.ts`.
 * The stub canvas makes a measurement nearly free, so the time here is the
 * shaping-service overhead only; the printed measureText count is the cost a
 * real canvas pays per unique string.
 */

const section: SectionProps = {
  pageWidth: 595, pageHeight: 842, marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
  headerDistance: 36, footerDistance: 36, titlePage: false, evenAndOddHeaders: false,
};

function body(): BodyElement[] {
  let seed = 1986;
  const random = () => ((seed = (seed * 1103515245 + 12345) % 2147483648) / 2147483648);
  const letters = 'etaoinshrdlucmfwypvbgkqjxz';
  const word = () => Array.from({ length: 2 + Math.floor(random() * 8) },
    () => letters[Math.floor(random() ** 2 * letters.length)]).join('');
  const run = (text: string, charSpacing?: number): DocParagraph['runs'][number] => {
    const value: DocxTextRun = {
      text, bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10.5, color: null, fontFamily: 'Bench', isLink: false, background: null,
      vertAlign: null, hyperlink: null, ...(charSpacing !== undefined ? { charSpacing } : {}),
    };
    return { type: 'text', ...value } as DocParagraph['runs'][number];
  };
  return Array.from({ length: 30 }, () => {
    const words = () => Array.from({ length: 40 }, word).join(' ');
    const p: DocParagraph = {
      alignment: 'both', indentLeft: 0, indentRight: 0, indentFirst: 0, spaceBefore: 0, spaceAfter: 6,
      lineSpacing: null, numbering: null, tabStops: [], defaultFontSize: 10.5, defaultFontFamily: 'Bench',
      runs: [run(words(), -0.5), run(` ${words()} `), run(words(), -0.25)],
    };
    return { type: 'paragraph', ...p } as BodyElement;
  });
}

const model = body();
const counter = installStubCanvas();
const context = new OffscreenCanvas(1, 1).getContext('2d')!;

describe('condensed spacing layout (#1986)', () => {
  const before = counter.measureTextCalls();
  layoutBodyModel(model, section, context);
  console.log(`measureText calls per layout: ${counter.measureTextCalls() - before}`);
  bench('layout, negative w:spacing', () => {
    layoutBodyModel(model, section, context);
  }, { iterations: 20, warmupIterations: 3 });
});

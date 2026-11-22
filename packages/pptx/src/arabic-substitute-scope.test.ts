import { describe, expect, it } from 'vitest';
import { paragraphInputRuns } from './renderer.js';
import type { Paragraph } from './types.js';
import { OFFICE_FONT_SUBSTITUTIONS, fontFamilyKey } from '@silurus/ooxml-core/internal/font-sources';

/** Every Office substitution as if its served face had loaded (the shape the
 * font-source session reports in `LoadedFontSources.substitutions`). */
const LOADED_OFFICE_SUBSTITUTIONS = new Map(OFFICE_FONT_SUBSTITUTIONS.map((entry) =>
  [fontFamilyKey(entry.for), { family: entry.use, ...(entry.script ? { script: entry.script } : {}) }]));


const renderContext = {
  themeMajorFont: null,
  themeMinorFont: null,
  dpr: 1,
  fontSources: { substitutions: LOADED_OFFICE_SUBSTITUTIONS, families: new Map() },
} as const;

function paragraph(text: string): Paragraph {
  return {
    runs: [{
      type: 'text',
      text,
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      fontSize: 18,
      color: null,
      fontFamily: 'Sakkal Majalla',
      fontFamilyCs: 'Sakkal Majalla',
    }],
    eaLnBrk: true,
  } as unknown as Paragraph;
}

describe('script-scoped Arabic visual substitutes', () => {
  it('splits a mixed DrawingML run so Latin does not use the Arabic substitute', () => {
    const { input } = paragraphInputRuns(
      paragraph('Lead العربية Tail'),
      24,
      '#000000',
      1,
      false,
      false,
      1,
      undefined,
      renderContext,
    );

    const textRuns = input.filter((item) => item.type === 'text');
    expect(textRuns.map((item) => item.text)).toEqual(['Lead ', 'العربية', ' Tail']);
    expect(textRuns[0]?.style.font).not.toContain('Noto Naskh Arabic');
    expect(textRuns[1]?.style.font).toContain('Noto Naskh Arabic');
    expect(textRuns[2]?.style.font).not.toContain('Noto Naskh Arabic');
  });

  it('keeps an Arabic grapheme and its combining mark in one substituted segment', () => {
    const vocalised = '\u0645\u064e';
    const { input } = paragraphInputRuns(
      paragraph(`A${vocalised}B`),
      24,
      '#000000',
      1,
      false,
      false,
      1,
      undefined,
      renderContext,
    );

    const textRuns = input.filter((item) => item.type === 'text');
    expect(textRuns.map((item) => item.text)).toEqual(['A', vocalised, 'B']);
    expect(textRuns[1]?.style.font).toContain('Noto Naskh Arabic');
  });
});

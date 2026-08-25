import { defineCompatibilityRule } from './compatibility.js';

export const WORD_TRACK_CHANGE_AUTHOR_PALETTE = defineCompatibilityRule({
  id: 'word-track-change-author-palette',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/compatibility.test.ts#pins the eight track-change author colors independently of author indexing',
  },
  description: 'Use the established eight-color revision-author palette while keeping the renderer deterministic author-index policy outside this compatibility claim.',
});

export const WORD_PARAGRAPH_SHADING_BORDER_BOX = defineCompatibilityRule({
  id: 'word-paragraph-shading-border-box',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/paragraph.test.ts#extends paragraph shading through visible border spacing',
  },
  description: 'Extend paragraph shading through each visible paragraph-border spacing interval so the fill reaches the painted border box.',
});

export const WORD_TRACK_CHANGE_DECORATION = defineCompatibilityRule({
  id: 'word-track-change-decoration',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/compatibility.test.ts#maps visible track-change kinds to their revision decorations',
  },
  description: 'When revision markup is visible, underline inserted text and strike through deleted text in the selected revision-author color.',
});

export const WORD_AUTO_TEXT_CONTRAST_EFFECTIVE_BACKGROUND = defineCompatibilityRule({
  id: 'word-auto-text-contrast-effective-background',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/cell-shading-auto-contrast.test.ts#paints a color-less run white inside a near-black cell',
  },
  description: 'Resolve automatic or never-authored text color against the nearest effective run, paragraph, or cell background before applying the deterministic contrast picker.',
});

export const WORD_RUN_DECORATION_JUSTIFIED_ADVANCE = defineCompatibilityRule({
  id: 'word-run-decoration-justified-advance',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/run-inline-formatting.test.ts#extends the border frame across justified inter-word slack',
  },
  description: 'Extend run shading, borders, underline, and strike decoration through the justification pitch owned by that run, including widened spaces.',
});

export const WORD_SNAP_TO_CHARS_TERMINAL_UNDERLINE = defineCompatibilityRule({
  id: 'word-snap-to-chars-terminal-underline',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'snap-to-chars-terminal-underline-boundaries',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'In the observed horizontal LTR snapToChars matrix, retain trailing character-cell slack in line advance while ending a terminal underline at the retained final-glyph ink extent. Authored trailing spaces remain content, and RTL/vertical text stays outside this rule.',
});

export const WORD_PARAGRAPH_BORDER_FLOW_RESERVATION = defineCompatibilityRule({
  id: 'word-paragraph-border-flow-reservation',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/para-bottom-border-flow.test.ts#a bottom border drops the following paragraph by exactly space + width/2',
  },
  description: 'Reserve a visible bottom paragraph border through its spacing interval and half stroke width so following flow begins below its painted outer edge.',
});

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

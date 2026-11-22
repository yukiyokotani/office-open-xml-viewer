import { defineCompatibilityRule } from './compatibility.js';

export const WORD_DRAWINGML_PATTERN_PAGE_GRID = defineCompatibilityRule({
  id: 'word-drawingml-pattern-page-grid',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'header-body-footer-drawingml-pattern-phase-controls',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'Header, body and footer anchored DrawingML pct50 and dnDiag fills use a 64×64 image tile at 8pt, with 1pt cells on the page-origin 8pt grid. All three stories preserve page axes under 315-degree rotation; header page-sized controls do as well. Chart and VML hosts are outside this rule.',
});

export const WORD_CLASSIC_CHART_SPACE_FRAME = defineCompatibilityRule({
  id: 'word-classic-chart-space-frame',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'classic-chart-space-frame-style-matrix',
    application: 'Microsoft Word',
    version: '16.112.4',
    platform: 'macOS 26.6.2',
  },
  description: 'For classic DOCX charts with omitted direct frame properties, the complete ST_Style 1..48 matrix uses 10pt rounded chart-space corners; omitted style and styles 1..40 use a 0.5pt #898989 outline, while styles 41..48 use no outline. Direct roundedCorners and line paint remain authoritative, followed by a linked chartStyle chartArea role.',
});

export const WORD_TRACK_CHANGE_AUTHOR_PALETTE = defineCompatibilityRule({
  id: 'word-track-change-author-palette',
  evidence: {
    kind: 'library-policy',
    reference: 'packages/docx/src/layout/compatibility.test.ts#pins the eight track-change author colors independently of author indexing',
  },
  description: 'Use the established eight-color revision-author palette while keeping the renderer deterministic author-index policy outside this compatibility claim. Library policy, no Office measurement: stable eight-hue author palette (PR #266); ECMA-376 §17.13 defines no revision display.',
});

export const WORD_TRACK_CHANGE_DECORATION = defineCompatibilityRule({
  id: 'word-track-change-decoration',
  evidence: {
    kind: 'library-policy',
    reference: 'packages/docx/src/layout/compatibility.test.ts#maps visible track-change kinds to their revision decorations',
  },
  description: 'When revision markup is visible, underline inserted text and strike through deleted text in the selected revision-author color. Library policy, no Office measurement: revision decoration styling; ECMA-376 §17.13 defines no display and the Word display settings are not stored in the document.',
});

export const WORD_PARAGRAPH_SHADING_BORDER_BOX = defineCompatibilityRule({
  id: 'word-paragraph-shading-border-box',
  evidence: {
    kind: 'library-policy',
    reference: 'packages/docx/src/layout/paragraph.test.ts#extends paragraph shading through visible border spacing',
  },
  description: 'Extend paragraph shading through each visible paragraph-border spacing interval so the fill reaches the painted border box. Library policy, no Office measurement: shading extent equals the border box; ECMA-376 §17.3.1.31 does not define it (32ffc42a4, test only).',
});

export const WORD_AUTO_TEXT_CONTRAST_EFFECTIVE_BACKGROUND = defineCompatibilityRule({
  id: 'word-auto-text-contrast-effective-background',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/cell-shading-auto-contrast.test.ts#paints a color-less run white inside a near-black cell',
  },
  description: 'Resolve automatic or never-authored text color against the nearest effective run, paragraph, or cell background before applying the deterministic contrast picker. Classification: normative, ECMA-376 §17.3.2.6 color auto (chosen from the background behind the run); the contrast threshold is library policy.',
});

export const WORD_RUN_DECORATION_JUSTIFIED_ADVANCE = defineCompatibilityRule({
  id: 'word-run-decoration-justified-advance',
  evidence: {
    kind: 'library-policy',
    reference: 'packages/docx/src/run-inline-formatting.test.ts#extends the border frame across justified inter-word slack',
  },
  description: 'Extend run shading, borders, underline, and strike decoration through the justification pitch owned by that run, including widened spaces. Library policy, no Office measurement: decorations follow the justified advance (PR #588); asserted, not measured.',
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
  description: 'Reserve a visible bottom paragraph border through its spacing interval and half stroke width so following flow begins below its painted outer edge. Classification: normative, ECMA-376 §17.3.1.7 bottom border space; Word PDF of a corpus document in PR #847.',
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

export const WORD_TRACK_CHANGE_BAR = defineCompatibilityRule({
  id: 'word-track-change-bar',
  evidence: {
    kind: 'library-policy',
    reference: 'packages/docx/src/track-changes-markup-layout.test.ts#emits one margin change bar per line containing revision text in the markup view',
  },
  description: 'In the markup view, draw a vertical change bar in the margin beside every line that contains tracked-change content, matching the Word reviewing-pane convention (an app convention; ECMA-376 defines no bar geometry). Library policy, no Office measurement: change-bar placement is an application convention (PR #1358); ECMA-376 §17.13 defines no display.',
});

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
  // ECMA-376 §17.13.5.22 / §17.13.5.25 — a move renders as a deletion at its
  // source and an insertion at its destination.
  if (kind === 'insertion' || kind === 'moveTo') return INSERTION_TRACK_CHANGE_DECORATION;
  if (kind === 'deletion' || kind === 'moveFrom') return DELETION_TRACK_CHANGE_DECORATION;
  return NO_TRACK_CHANGE_DECORATION;
}

export const WORD_TEXT_BOX_CONTENT_CLIP = defineCompatibilityRule({
  id: 'word-text-box-content-clip',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'cell-content-clip-matrix',
    application: 'Microsoft Word',
    version: '16.0.20625',
    platform: 'Windows',
  },
  description: 'ECMA-376 §20.4.2.37 / §21.1.2.1.1 define the text-box insets but no clip. Windows Word 16.0.20625 clips a fixed-size (noAutofit) horizontal DrawingML text box to one box: horizontally the whole shape (a 300pt box with 2.83pt left and right insets clips at its edges, not at the insets), vertically inside the top and bottom insets (28.35pt and 2.83pt). A picture raised past the first line is cut at the top inset; one lowered past the last line is drawn while it stays above the bottom inset. The clip is the box, not the line, as for table cells (word-cell-content-clip). The clip edges lie a further half outline (0.25pt) inside; the library clips at the shape edge. Library scope: vertical text boxes, VML text boxes and autofit boxes keep their established clip (unmeasured).',
});

/** WORD_TEXT_BOX_CONTENT_CLIP: the clip of a horizontal DrawingML text box,
 * from its shape bounds and its inner (inset) bounds. */
export function wordTextBoxContentClipBounds<Rect extends Readonly<{ xPt: number; yPt: number; widthPt: number; heightPt: number }>>(
  shapeBounds: Rect,
  innerBounds: Rect,
): Readonly<{ xPt: number; yPt: number; widthPt: number; heightPt: number }> {
  void WORD_TEXT_BOX_CONTENT_CLIP;
  return { xPt: shapeBounds.xPt, yPt: innerBounds.yPt, widthPt: shapeBounds.widthPt, heightPt: innerBounds.heightPt };
}

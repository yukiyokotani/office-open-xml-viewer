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
    kind: 'office-observation',
    syntheticFixtureId: 'paragraph-border-extent-matrix',
    application: 'Microsoft Word',
    version: '16.0.20625',
    platform: 'Windows',
  },
  description: 'ECMA-376 §17.3.1.42 and §17.3.1.7 place a top or bottom paragraph border w:space from the text but do not say how much flow it takes or where the stroke lies. Windows Word 16.0.20625 controls (pbdr1, Noto Sans 9pt, single borders of w:sz 2, 4, 8, 12, 24 and 48 with w:space 0, 1 and 4 on all four sides, exact 12pt and automatic single lines, compatibility modes 15 and 14) reserve space + the full stroke width above the first line and below the last line, and draw the whole stroke outside the border box (text ± w:space): a 1.5pt top stroke over w:space 1 spans 0-1.5pt above the paragraph top and the text starts 2.5pt below it, likewise below. Thick, dotted and dashed borders of w:sz 12 take the same width; a top-only or bottom-only border reserves only its own side. Space before and space after add to the reservation rather than overlapping it (a 6pt space after a 5.5pt reservation moves the next paragraph 11.5pt). Two paragraphs with identical borders form one box with no reservation between them; a between border takes its w:space on both sides of its stroke (w:sz 4 w:space 1: 2.5pt; w:sz 12 w:space 4: 9.5pt); different borders keep both reservations. A bar border and shading alone reserve nothing, and shading fills the box between the strokes. The library applies the same reservation to body and table-cell paragraphs; frame groups keep their own gap model. Compound styles reserve their whole painted thickness the same way (pbdr2: 17 styles at w:sz 4, 12, 24 and 48, w:space 0 and 4, box, top-only, bottom-only and a between pair): double 3w, triple 5w, thin-thick and thick-thin small gap w + 1.5 (0.75pt thin stroke and gap), thin-thick-thin small gap w + 3, medium gap 2w (thin stroke and gap w/2), thin-thick-thin medium gap 3w, large gap w + 2.25 (1.5pt and 0.75pt strokes, gap w), thin-thick-thin large gap 2w + 3, 3-D emboss and engrave w plus two outer strokes of w/2 clamped to 0.75-1.5pt, inset and outset w, double wave 5.25pt and wave 3pt, except wave at w:sz 12, which takes 3.75pt (pbdr3: wave and doubleWave at w:sz 2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 24, 36 and 48, w:space 0 and 4 boxes and w:space 1 top-only borders, exact 12pt lines, mode 15: every reservation matches within the 0.1pt PDF precision). The wave ink does not depend on w:sz: Word paints a hairline zigzag about 0.375pt thick (one 1/192in raster cell) that fills the band, rising and falling the band height over the band height of run (a 6pt period in a 3pt band, 7.5pt at w:sz 12); a double wave paints two such zigzags stacked in its 5.25pt band, each about half the band high with a period of about 5.2pt (Word\'s PDF export draws these as raster tiles of which only part of the ink survives, so the zigzag shape beyond its band, period and hairline is inferred). The library paints each style\'s own stroke pattern centred in that band, waves as one or two hairline zigzags filling it. The earlier regression evidence (PR #847, a 1.5pt bottom rule followed by text) measured the rule-to-following-text gap, which the full-width model leaves unchanged because the stroke moves with the reservation.',
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
  description: 'ECMA-376 §20.4.2.37 / §21.1.2.1.1 define the text-box insets but no clip. Windows Word 16.0.20625 clips a fixed-size (noAutofit, which §21.1.2.1.2 also implies when bodyPr names no auto-fit) horizontal DrawingML text box to one box: horizontally the whole shape (a 300pt box with 2.83pt left and right insets clips at its edges, not at the insets), vertically inside the top and bottom insets (28.35pt and 2.83pt). A picture raised past the first line is cut at the top inset; one lowered past the last line is drawn while it stays above the bottom inset. The clip is the box, not the line, as for table cells (word-cell-content-clip). The clip edges lie a further half outline (0.25pt) inside; the library clips at the shape edge. Library scope: vertical text boxes, VML text boxes and autofit boxes keep their established clip (unmeasured).',
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

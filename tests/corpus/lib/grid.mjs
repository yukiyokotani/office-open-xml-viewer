// The grid. Every position, indent, tab stop, column width and size in the
// corpus is derived from the constants in this file; case modules never write
// a free-standing length.

/** One sans family in two weights (400, 700) and its Japanese companion.
 * Roboto is the open-licensed grotesque that the library's Google Fonts
 * substitute registry already resolves (packages/core/src/fonts/google-fonts.ts,
 * 400/700 + italics); Noto Sans JP is the Japanese face of the shared script
 * registry (packages/core/src/fonts/scripts.ts). Both are SIL OFL 1.1. */
export const FONT = { latin: 'Roboto', eastAsia: 'Noto Sans JP' };

/** Black on white; the accent is used only when a feature needs a colour. */
export const COLOR = { ink: '000000', paper: 'FFFFFF', accent: 'D97757' };

/** Text inset (EMU) of a DrawingML shape whose outline is drawn: one 12 pt
 * gutter, so text never touches the rule. Text in unoutlined shapes, text
 * boxes and table cells keeps zero insets and starts on the field edge. */
export const OUTLINED_INSET = 152400;

export const LANG = { latin: 'en-GB', eastAsia: 'ja-JP' };

// ── DOCX: A4 portrait, twentieths of a point (twips) ─────────────────────────

const mmToTwips = (mm) => Math.round((mm * 1440) / 25.4);

/** 4 columns of 37.5 mm with 5 mm gutters; edges are rounded individually from
 * their millimetre positions so rounding never accumulates across the page. */
const COLUMN_MM = 37.5;
const GUTTER_MM = 5;
const colStartMm = (c) => (c - 1) * (COLUMN_MM + GUTTER_MM);

export const DOCX = (() => {
  const LINE = 240; // 12 pt baseline
  const LINES = 60;
  const pageW = mmToTwips(210); // ISO 216 A4
  const pageH = mmToTwips(297);
  const left = mmToTwips(25);
  const top = mmToTwips(20);
  const colStart = (c) => mmToTwips(colStartMm(c));
  const colEnd = (c) => mmToTwips(colStartMm(c) + COLUMN_MM);
  const liveW = colEnd(4); // 165 mm
  const liveH = LINE * LINES; // 254 mm = 720 pt exactly
  return {
    LINE,
    LINES,
    PAGE_W: pageW,
    PAGE_H: pageH,
    MARGIN: { left, top, right: pageW - left - liveW, bottom: pageH - top - liveH },
    LIVE_W: liveW,
    COLUMNS: 4,
    GUTTER: mmToTwips(GUTTER_MM),
    colStart,
    colEnd,
    /** First line (1-based) of grid row r: 6 rows of 9 lines, 1 empty line between. */
    rowLine: (r) => (r - 1) * 10 + 1,
    /** Distance from the top of the live area to the top of grid line n. */
    lineTop: (n) => (n - 1) * LINE,
    /** Gutter between grid column c and c + 1 (283 or 284 after rounding). */
    gutterAfter: (c) => colStart(c + 1) - colEnd(c),
    /** Width of grid columns c1..c2 including the gutters between them. */
    span: (c1, c2) => colEnd(c2) - colStart(c1),
    /** Header line: two baselines above the live area (one empty line between).
     * Footer line: likewise one empty line below it. Values are w:pgMar
     * header/footer distances from the page edge. */
    HEADER_DIST: top - 2 * LINE,
    FOOTER_DIST: pageH - top - liveH - 2 * LINE,
    /** Twips to EMU (DrawingML lengths inside WordprocessingML). */
    emu: (twips) => twips * 635,
    /** Type: size in half-points, leading in twips. */
    TYPE: {
      caption: { sz: 16, leading: LINE },
      text: { sz: 18, leading: LINE },
      title: { sz: 48, leading: 2 * LINE },
    },
  };
})();

// ── PPTX: 16:9, English Metric Units ──────────────────────────────────────────

export const PPTX = (() => {
  const SLIDE_W = 12192000; // 13.333 in
  const SLIDE_H = 6858000; // 7.5 in
  const MARGIN = 457200; // 0.5 in on all four sides
  const GUTTER = 152400; // 12 pt, equal to one baseline
  const COLUMNS = 12;
  const COL = (SLIDE_W - 2 * MARGIN - (COLUMNS - 1) * GUTTER) / COLUMNS; // 800100
  const LINE = 152400; // 12 pt baseline
  const ROW_LINES = 9; // 4 rows of 9 lines + 3 gaps of 1 line = 39 lines
  const ROWS = 4;
  const ROW = ROW_LINES * LINE;
  if (!Number.isInteger(COL)) throw new Error('PPTX column width must be an integer EMU');
  if (2 * MARGIN + ROWS * ROW + (ROWS - 1) * LINE !== SLIDE_H) throw new Error('PPTX rows do not fill the slide');
  const x = (c) => MARGIN + (c - 1) * (COL + GUTTER);
  const y = (r) => MARGIN + (r - 1) * (ROW + LINE);
  return {
    SLIDE_W,
    SLIDE_H,
    MARGIN,
    GUTTER,
    COLUMNS,
    COL,
    LINE,
    ROWS,
    ROW,
    /** Left edge of column c (1-based). */
    x,
    /** Top edge of row r (1-based). */
    y,
    /** Width of columns c1..c2 including the gutters between them. */
    w: (c1, c2) => (c2 - c1 + 1) * COL + (c2 - c1) * GUTTER,
    /** Height of rows r1..r2 including the gaps between them. */
    h: (r1, r2) => (r2 - r1 + 1) * ROW + (r2 - r1) * LINE,
    /** One column plus one gutter: the indent module. */
    STEP: COL + GUTTER,
    /** Type: size and leading in hundredths of a point. */
    TYPE: {
      caption: { sz: 1000, leading: 1200 },
      text: { sz: 2000, leading: 2400 },
      title: { sz: 4000, leading: 4800 },
    },
  };
})();

// ── XLSX: column widths in characters, row heights in points ─────────────────

export const XLSX = {
  /** Column A is the left margin. */
  MARGIN_COL: 2,
  /** The column module, in characters of the default font's maximum digit width. */
  MODULE: 12,
  /** Rows sit on the 12 pt baseline: one baseline per row for 8 and 9 pt
   * text, three for the 24 pt title row. */
  ROW_PT: 12,
  TITLE_ROW_PT: 36,
  /** Type: sizes in points. */
  TYPE: { caption: 8, text: 9, title: 24 },
  /** First content row; rows 2-4 hold the header furniture, row 5 is empty. */
  FIRST_ROW: 6,
};

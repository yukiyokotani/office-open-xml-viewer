// Format-agnostic OOXML types shared across pptx/docx/xlsx packages.
// All positions and sizes are in EMUs (English Metric Units).

export type PathCmd =
  | { cmd: 'moveTo';     x: number; y: number }
  | { cmd: 'lineTo';     x: number; y: number }
  | { cmd: 'cubicBezTo'; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
  | { cmd: 'arcTo';      wr: number; hr: number; stAng: number; swAng: number }
  | { cmd: 'close' };

export type Fill = SolidFill | NoFill | GradientFill;

export interface SolidFill {
  fillType: 'solid';
  color: string; // hex 6-char or 8-char (RRGGBBAA with alpha)
}

export interface NoFill {
  fillType: 'none';
}

export interface GradientStop {
  position: number; // 0.0–1.0
  color: string;    // hex 6 or 8 chars
}

export interface GradientFill {
  fillType: 'gradient';
  stops: GradientStop[];
  /** degrees: 0 = left→right, 90 = top→bottom */
  angle: number;
  /** 'linear' | 'radial' */
  gradType: string;
}

export interface Shadow {
  color: string;  // hex 6 chars
  alpha: number;  // 0.0–1.0
  blur: number;   // EMU
  dist: number;   // EMU
  /** degrees clockwise from East */
  dir: number;
}

export interface ArrowEnd {
  /** OOXML type: "none" | "triangle" | "stealth" | "diamond" | "oval" | "arrow" */
  type: string;
  /** Width multiplier: "sm" | "med" | "lg" */
  w: string;
  /** Length multiplier: "sm" | "med" | "lg" */
  len: string;
}

export interface Stroke {
  color: string;
  /** Width in EMU */
  width: number;
  /** OOXML prstDash value: "dash", "dot", "dashDot", "lgDash", "lgDashDot", etc. */
  dashStyle?: string;
  /** Arrow head at the start of the line */
  headEnd?: ArrowEnd;
  /** Arrow head at the end of the line */
  tailEnd?: ArrowEnd;
}

export interface TextBody {
  /** Vertical anchor: "t" | "ctr" | "b" */
  verticalAnchor: string;
  paragraphs: Paragraph[];
  /** Default pt size from lstStyle (overrides renderer default when present) */
  defaultFontSize: number | null;
  /** Inherited bold from layout/master defRPr (null = not set, use false as final default) */
  defaultBold: boolean | null;
  /** Inherited italic from layout/master defRPr (null = not set, use false as final default) */
  defaultItalic: boolean | null;
  /** Text insets in EMU (defaults: lIns=rIns=91440, tIns=bIns=45720) */
  lIns: number;
  rIns: number;
  tIns: number;
  bIns: number;
  /** "square" = wrap, "none" = no wrap */
  wrap: string;
  /** Text direction: "horz" | "vert" | "vert270" | "eaVert" etc. */
  vert: string;
  /** Auto-fit: "sp" = shape grows to fit text, "norm" = font shrinks, "none" = no fit */
  autoFit: string;
}

export type SpaceLine =
  | { type: 'pct'; val: number }   // val: e.g. 100000 = 100%, 150000 = 150%
  | { type: 'pts'; val: number };  // val in points

export type Bullet =
  | { type: 'none' }
  | { type: 'inherit' }
  | { type: 'char'; char: string; color: string | null; sizePct: number | null; fontFamily: string | null }
  | { type: 'autoNum'; numType: string; startAt: number | null };

export interface TabStop {
  /** Position in EMU from the left edge of the text area (after lIns) */
  pos: number;
  /** Alignment: "l" | "r" | "ctr" | "dec" */
  algn: string;
}

export interface Paragraph {
  /** Alignment: "l" | "ctr" | "r" | "just" */
  alignment: string;
  /** Left margin in EMU */
  marL: number;
  /** Right margin in EMU */
  marR: number;
  /** First-line indent in EMU (negative = hanging indent) */
  indent: number;
  spaceBefore: number | null;
  spaceAfter: number | null;
  spaceLine: SpaceLine | null;
  /** List nesting level (0–8) */
  lvl: number;
  bullet: Bullet;
  defFontSize: number | null;
  defColor: string | null;
  defBold: boolean | null;
  defItalic: boolean | null;
  defFontFamily: string | null;
  /** Tab stops from pPr > tabLst */
  tabStops: TabStop[];
  runs: TextRun[];
}

export type TextRun = TextRunData | LineBreak;

export interface TextRunData {
  type: 'text';
  text: string;
  /** null = not set, inherit from paragraph/body defaults */
  bold: boolean | null;
  /** null = not set, inherit from paragraph/body defaults */
  italic: boolean | null;
  underline: boolean;
  /** true when rPr strike = "sngStrike" or "dblStrike" */
  strikethrough: boolean;
  /** Font size in points */
  fontSize: number | null;
  color: string | null;
  fontFamily: string | null;
  /** Baseline shift in thousandths of a point. Positive = superscript, negative = subscript. */
  baseline?: number;
  /** Set for OOXML field runs (e.g. "slidenum"). When set, renderer replaces text with field value. */
  fieldType?: string;
}

export interface LineBreak {
  type: 'break';
}

export interface RenderOptions {
  width?: number;
  defaultTextColor?: string | null;
  dpr?: number;
  majorFont?: string | null;
  minorFont?: string | null;
}

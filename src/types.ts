// ===== Presentation data model =====
// All positions and sizes are in EMUs (English Metric Units).
// 914400 EMU = 1 inch, 12700 EMU = 1 pt

export interface Presentation {
  slideWidth: number;
  slideHeight: number;
  slides: Slide[];
}

export interface Slide {
  index: number;
  background: Fill | null;
  elements: SlideElement[];
}

export type SlideElement = ShapeElement | PictureElement | TableElement;

/** A single command in a custom geometry path (coordinates normalised to [0,1]). */
export type PathCmd =
  | { cmd: 'moveTo';     x: number; y: number }
  | { cmd: 'lineTo';     x: number; y: number }
  | { cmd: 'cubicBezTo'; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
  | { cmd: 'arcTo';      wr: number; hr: number; stAng: number; swAng: number }
  | { cmd: 'close' };

export interface ShapeElement {
  type: 'shape';
  x: number;
  y: number;
  width: number;
  height: number;
  /** Rotation in degrees, clockwise */
  rotation: number;
  /** Horizontal mirror (a:xfrm flipH) */
  flipH: boolean;
  /** Vertical mirror (a:xfrm flipV) */
  flipV: boolean;
  /** OOXML preset name or "custGeom" when custom paths are used */
  geometry: string;
  fill: Fill | null;
  stroke: Stroke | null;
  textBody: TextBody | null;
  /** Default text color from p:style > fontRef (hex). Used when run/para has no explicit color. */
  defaultTextColor: string | null;
  /** Custom geometry sub-paths (set only when geometry === "custGeom").
   *  Outer array: one entry per <a:path>; inner: path commands with coords in [0,1]. */
  custGeom: PathCmd[][] | null;
  /** First adjustment value from prstGeom avLst (e.g. trapezoid inset). Range 0–100000. */
  adj: number | null;
  /** Second adjustment value from prstGeom avLst (e.g. arrow head width). Range 0–100000. */
  adj2: number | null;
  /** Drop shadow from effectLst > outerShdw (null if not present). */
  shadow: Shadow | null;
}

export interface TableElement {
  type: 'table';
  x: number;
  y: number;
  width: number;
  height: number;
  /** Column widths in EMU */
  cols: number[];
  rows: TableRow[];
}

export interface TableRow {
  /** Row height in EMU */
  height: number;
  cells: TableCell[];
}

export interface TableCell {
  textBody: TextBody | null;
  fill: Fill | null;
  borderL: Stroke | null;
  borderR: Stroke | null;
  borderT: Stroke | null;
  borderB: Stroke | null;
  /** Column span */
  gridSpan: number;
  /** Row span */
  rowSpan: number;
  /** Horizontal merge continuation */
  hMerge: boolean;
  /** Vertical merge continuation */
  vMerge: boolean;
}

export interface PictureElement {
  type: 'picture';
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
  flipH: boolean;
  flipV: boolean;
  /** Data URL, e.g. "data:image/png;base64,..." */
  dataUrl: string;
}

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

export interface Stroke {
  color: string;
  /** Width in EMU */
  width: number;
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
}

export interface LineBreak {
  type: 'break';
}

// ===== Worker message protocol =====

export type WorkerRequest =
  | { kind: 'parse'; id: number; buffer: ArrayBuffer };

export type WorkerResponse =
  | { kind: 'ready' }
  | { kind: 'parsed'; id: number; presentation: Presentation }
  | { kind: 'error'; id: number; message: string };

export { DocxDocument, type LoadOptions, type RenderPageToBitmapOptions } from './document';
export type { WireRenderPageOptions } from './worker-protocol';
export { DocxViewer, type DocxViewerOptions } from './viewer';
export { DocxScrollViewer, type DocxScrollViewerOptions } from './scroll-viewer';
export { buildDocxTextLayer, getDocxTextRunAtNode } from './text-layer';
// IX2 find-in-document: the highlight overlay builder + the docx match-location
// shape. `FindMatch` / `FindMatchesOptions` come from core (shared across formats).
export {
  buildDocxHighlightLayer,
  type DocxHighlightMatch,
  type DocxHighlightColors,
} from './find-highlight-layer';
export type { DocxMatchLocation } from './find';
export type { FindMatch, FindMatchesOptions } from '@silurus/ooxml-core';
export { autoResize, type AutoResizeOptions } from '@silurus/ooxml-core';
// IX1 — the shared hyperlink target shape surfaced by `DocxViewerOptions.
// onHyperlinkClick`, `DocxTextRunInfo.hyperlink`, and the 5th arg of
// `buildDocxTextLayer`, plus the default "open in a new tab, sanitised" helper.
export { type HyperlinkTarget, openExternalHyperlink } from '@silurus/ooxml-core';
export type { TextSourcePathStep, TextSourceRef } from '@silurus/ooxml-core';
export { sliceTextSourceRefs } from '@silurus/ooxml-core';
// Typed load-time error surfaced by DocxDocument.load (e.g. a password-protected
// or legacy-binary .doc file). Re-exported so `@silurus/ooxml/docx` consumers can
// narrow on `err.code`.
export { OoxmlError, type OoxmlErrorCode } from '@silurus/ooxml-core';
export { noteText } from './types';
export type {
  DocxDocumentModel,
  DocSettings,
  // Embedded-font reference (reachable via DocxDocumentModel.embeddedFonts,
  // ECMA-376 §17.8.3.3-.6). The viewer de-obfuscates + registers these.
  EmbeddedFontRef,
  SectionProps,
  // Per-section page geometry (reachable via the BodyElement sectionBreak arm's
  // `geom` and PaginatedBodyElement's `sectionGeom`, ECMA-376 §17.6.13/§17.6.11).
  SectionGeom,
  // Per-section page-numbering settings (reachable via SectionProps.pageNumType
  // and the BodyElement sectionBreak arm's `pageNumType`, ECMA-376 §17.6.12).
  PageNumType,
  // Per-section page decorations (reachable via SectionProps): page borders
  // (§17.6.10) and line numbering (§17.6.8).
  PageBorders,
  PageBorderEdge,
  LineNumbering,
  // Multi-column section sub-types (reachable via SectionProps.columns).
  ColumnsSpec,
  ColSpec,
  HeadersFooters,
  HeaderFooter,
  NumberingInfo,
  BodyElement,
  DocParagraph,
  DocRun,
  // Absolute-position tab run (reachable via the DocRun union's `ptab` arm,
  // ECMA-376 §17.3.3.23).
  PTabRun,
  DocxTextRun,
  FieldRun,
  ImageRun,
  // DrawingML chart run (reachable via the DocRun union's `chart` arm,
  // ECMA-376 §21.2).
  ChartRun,
  AnchorHostMetrics,
  ShapeRun,
  // VML `<v:textpath>` watermark text (reachable via ShapeRun.textPath).
  TextPath,
  ShapeText,
  // Per-run shape-text formatting (reachable via ShapeText.runs).
  ShapeTextRun,
  RubyAnnotation,
  RenderPageOptions,
  RunRevision,
  DocRevision,
  DocComment,
  DocNote,
  NoteRef,
  // Paragraph / line-spacing sub-types.
  LineSpacing,
  // Text-frame / drop-cap properties (reachable via DocParagraph.framePr).
  FramePr,
  TabStop,
  ParagraphBorders,
  ParaBorderEdge,
  DocxRunBorder,
  // Table model (reachable via BodyElement table variant).
  DocTable,
  TblpPr,
  DocTableRow,
  DocTableCell,
  CellElement,
  TableBorders,
  CellBorders,
  BorderSpec,
  // Shape geometry / fill sub-types.
  PathCmd,
  GradientStop,
  LineEnd,
} from './types';
export type { DocxTextRunInfo } from './renderer';

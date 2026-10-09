/**
 * Private, projection-owned canonical CSS widths for ACTIVE view-only column
 * resizes.
 *
 * Settled view-only policy: a user drag captures a logical CSS pixel size, so
 * the drag's intent is that pixel size, not the stored-width number written
 * beside it. A later MDW change (retained font rebind) must keep a user-resized
 * column at its dragged CSS px, while authored stored widths keep decoding via
 * the current MDW. Raw `colWidths` values stay as they are (the drag still
 * writes `pxToColWidth`), and this sidecar is never a public Worksheet field.
 * `null` removes an override; 0 means hidden. No DPR/device-unit math lives
 * here. The drag captures integer CSS px; existing per-band display-scale
 * rounding is unchanged, so this is not fractional Office-width support.
 * Worksheets without an override allocate nothing.
 */
export { columnCssWidths, getColumnCssWidth, setColumnCssWidth, inheritColumnCssWidths }
  from './worksheet-size-context.js';

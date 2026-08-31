import { overlayPercent, type HyperlinkTarget } from '@silurus/ooxml-core';
import type { PptxTextRunInfo } from './renderer';
import { pptxRunFrameKey, pptxRunFrameTransform } from './run-frame-transform.js';

function setSelectionData(element: HTMLElement, key: string, value?: string): void {
  if (element.dataset) {
    if (value === undefined) delete element.dataset[key];
    else element.dataset[key] = value;
  } else if (value !== undefined) {
    element.setAttribute?.(`data-${key.replace(/[A-Z]/g, (char) => `-${char.toLowerCase()}`)}`, value);
  }
}

/**
 * Build the transparent text-selection overlay for a rendered pptx slide. Unlike
 * docx (flat spans), pptx groups runs into one positioned `<div>` per shape frame
 * (keyed by the shape's geometry + total rotation) and applies a CSS `rotate()` to
 * the group when the shape is rotated, so the browser selection tracks the drawn,
 * rotated text as a unit. Each run's `<span>` is absolutely positioned INSIDE its
 * shape div (`inShapeX`/`inShapeY`). Extracted verbatim from
 * `PptxViewer._buildTextLayer` so the pager (PptxViewer) and the continuous-scroll
 * viewer (PptxScrollViewer, WS4) share one implementation; public API for
 * integrators (design §10). IX6 — usable in BOTH render modes: worker mode
 * collects the same `PptxTextRunInfo[]` off-thread and ships it back beside the
 * bitmap, so the overlay is built from identical geometry regardless of thread.
 *
 * IX1 — when a run carries a resolved `hyperlink` (from `<a:hlinkClick>`) and an
 * `onHyperlinkClick` callback is supplied, its span becomes a click target
 * (`cursor:pointer`, a `title` tooltip, and a `click` handler). A plain span
 * (no hyperlink) is byte-identical to before. A JS click handler is used rather
 * than an `<a href>` so the URL never bypasses the viewer's sanitisation.
 *
 * The overlay's coordinates are all PERCENTAGES of `cssWidth`/`cssHeight` (the
 * slide's intended CSS-px box), never literal px, and the container's own
 * `width`/`height` are left untouched (the caller sizes it `width:100%;
 * height:100%` so it fills the wrapper). This lets the overlay track the canvas's
 * ACTUAL rendered box even when a consumer scales the canvas down with external
 * CSS (e.g. `width:100%!important;height:auto`): the wrapper (and therefore the
 * `100%` container) shrinks with the canvas, and every `%`-placed child scales
 * with it, so nothing overflows the wrapper into an ancestor's scroll area.
 *
 * @param layer     the overlay div (sized `width:100%;height:100%` by the caller).
 * @param runs      per-run + per-shape geometry from `renderSlide({ onTextRun })`.
 * @param cssWidth  the slide's intended CSS width (px, number) — the % denominator.
 * @param cssHeight the slide's intended CSS height (px, number) — the % denominator.
 * @param onHyperlinkClick called with the run's resolved {@link HyperlinkTarget}
 *                         when a hyperlink span is clicked. Omit to leave links
 *                         non-interactive (spans stay plain, selectable text).
 */
export function buildPptxTextLayer(
  layer: HTMLDivElement,
  runs: PptxTextRunInfo[],
  cssWidth: number,
  cssHeight: number,
  onHyperlinkClick?: (target: HyperlinkTarget) => void,
  slideIndex?: number,
): void {
  layer.innerHTML = '';
  setSelectionData(layer, 'ooxmlSelectionSurface', 'pptx');
  setSelectionData(layer, 'slideIndex', slideIndex === undefined ? undefined : String(slideIndex));

  // Group runs by shape (same shapeX/shapeY/rotation)
  const shapeMap = new Map<string, { div: HTMLDivElement; w: number; h: number }>();

  const ownerDocument = layer.ownerDocument ?? document;
  for (const [runIndex, run] of runs.entries()) {
    const transform = pptxRunFrameTransform(run);
    const key = pptxRunFrameKey(run, transform);
    let shape = shapeMap.get(key);
    if (!shape) {
      const div = ownerDocument.createElement('div');
      // The shape frame is placed as a % of the slide box so it tracks the
      // canvas's actual rendered size; its width/height are % too, so the child
      // spans (positioned as % of THIS box) scale with it. rotate() is applied to
      // the %-sized box unchanged — the rotation centre is within the box, so it
      // composes with the outer scale.
      div.style.cssText =
        `position:absolute;` +
        `left:${overlayPercent(run.shapeX, cssWidth)};top:${overlayPercent(run.shapeY, cssHeight)};` +
        `width:${overlayPercent(run.shapeW, cssWidth)};height:${overlayPercent(run.shapeH, cssHeight)};` +
        // DrawingML text may paint beyond its shape rectangle when the body does
        // not autofit. Match the canvas renderer by keeping those runs selectable;
        // the outer slide text layer still clips anything past the slide edge.
        `pointer-events:all;overflow:visible;`;
      if (transform) {
        div.style.transformOrigin = 'center center';
        div.style.transform = transform;
      }
      shape = { div, w: run.shapeW, h: run.shapeH };
      shapeMap.set(key, shape);
      layer.appendChild(div);
    }

    const span = ownerDocument.createElement('span');
    setSelectionData(span, 'ooxmlSelectionRun', 'pptx');
    setSelectionData(span, 'runIndex', String(runIndex));
    if (run.shapeId !== undefined) setSelectionData(span, 'shapeId', run.shapeId);
    if (run.elementIndex !== undefined) setSelectionData(span, 'elementIndex', String(run.elementIndex));
    if (run.origin !== undefined) setSelectionData(span, 'elementOrigin', run.origin);
    span.textContent = run.text;
    // The `font` shorthand must precede `line-height` because the shorthand
    // resets `line-height` to `normal`. Reset `letter-spacing` so a parent
    // CSS rule cannot drift the trailing edge of the selection. Kerning /
    // ligatures are left at the browser default ('auto') because canvas
    // `measureText` / `fillText` also apply them by default — forcing them
    // off here would make the span wider than the drawn text.
    // Hyperlink runs become clickable when a handler is wired: pointer cursor +
    // a tooltip + a JS click handler. `color:transparent` is kept (the glyphs
    // are painted on the canvas underneath, incl. the theme hyperlink colour and
    // underline) — the span only supplies the hit region. `cursor:text` stays
    // the default for plain runs so selection UX is unchanged.
    const link = onHyperlinkClick ? run.hyperlink : undefined;
    // Position the span as a % of its shape frame (shapeW/shapeH) so it tracks
    // the frame when the whole overlay is scaled by external CSS. `font` /
    // `line-height` stay px: they set the glyph metrics of the transparent hit
    // text, which the browser lays out inside the box; the box position is what
    // must track the canvas (byte-identical to the prior overlay at 100% scale).
    span.style.cssText =
      `position:absolute;` +
      `left:${overlayPercent(run.inShapeX, shape.w)};top:${overlayPercent(run.inShapeY, shape.h)};` +
      `font:${run.font};line-height:${run.h}px;letter-spacing:0;` +
      `white-space:pre;color:transparent;cursor:${link ? 'pointer' : 'text'};`;
    if (link && onHyperlinkClick) {
      span.title = link.kind === 'external' ? link.url : link.ref;
      span.addEventListener('click', (e) => {
        e.preventDefault();
        onHyperlinkClick(link);
      });
    }
    shape.div.appendChild(span);
  }
}

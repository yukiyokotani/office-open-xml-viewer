/**
 * Paint a Canvas only after the browser's font set has settled.
 *
 * The first paint is deliberate: Canvas font selection can discover a pending
 * face that was not requested by DOM text. Once that paint has made every
 * dependency visible, wait for the FontFaceSet and repaint so the captured
 * pixels cannot contain a transient fallback glyph. Two animation frames give
 * the browser a stable capture boundary without relying on machine-dependent
 * wall-clock delays.
 */
export async function stableCanvasRender({ render, fonts, nextFrame }) {
  if (fonts) await fonts.ready;
  await render();
  if (fonts) await fonts.ready;
  await render();
  await nextFrame();
  await nextFrame();
}

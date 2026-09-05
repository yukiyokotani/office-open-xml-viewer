/** Build the shared progressive-slide loading surface used by both PPTX viewers. */
export function createPptxLoadingLayer(ownerDocument: Document): HTMLSpanElement {
  const layer = ownerDocument.createElement('span');
  layer.style.cssText = [
    'position:absolute',
    'top:0',
    'right:0',
    'bottom:0',
    'left:0',
    'display:none',
    'align-items:center',
    'justify-content:center',
    'background:rgba(255,255,255,0.72)',
    'backdrop-filter:blur(2px)',
    'pointer-events:none',
    'z-index:4',
  ].join(';');
  layer.setAttribute('role', 'status');
  layer.setAttribute('aria-live', 'polite');
  layer.setAttribute('aria-label', 'Loading slide');

  const circle = ownerDocument.createElement('span');
  circle.className = 'ooxml-pptx-progress-circle';
  circle.style.cssText = [
    'width:34px',
    'height:34px',
    'box-sizing:border-box',
    'border-radius:50%',
    'border:3px solid var(--border-bright, rgba(100,116,139,0.28))',
    'border-top-color:var(--signal, #12bfd8)',
    'box-shadow:0 2px 10px rgba(15,23,42,0.08)',
  ].join(';');
  circle.setAttribute('aria-hidden', 'true');
  layer.appendChild(circle);

  const reducedMotion = ownerDocument.defaultView
    ?.matchMedia?.('(prefers-reduced-motion: reduce)').matches ?? false;
  if (!reducedMotion && typeof circle.animate === 'function') {
    circle.animate(
      [{ transform: 'rotate(0deg)' }, { transform: 'rotate(360deg)' }],
      { duration: 800, iterations: Infinity, easing: 'linear' },
    );
  }

  return layer;
}

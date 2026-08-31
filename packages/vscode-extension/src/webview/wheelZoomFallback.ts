import { zoomStepScale, type ZoomableViewer } from '@silurus/ooxml-core';

export interface WheelZoomEvent {
  ctrlKey: boolean;
  metaKey: boolean;
  deltaY: number;
  deltaMode: number;
  preventDefault(): void;
}

/**
 * Capture a Chromium trackpad pinch before the VS Code webview can consume it.
 *
 * The scroll viewers normally handle the same Ctrl/Command+wheel event on their
 * private scroll host. We defer the fallback by one microtask and only act when
 * that handler did not change the scale, avoiding a double zoom while still
 * covering events targeted at the webview frame rather than the scroll host.
 */
export function captureUnhandledWheelZoom(
  event: WheelZoomEvent,
  viewer: ZoomableViewer | null,
  defer: (callback: () => void) => void = queueMicrotask,
  onError?: (error: unknown) => void,
): boolean {
  if (!viewer || !(event.ctrlKey || event.metaKey) || event.deltaY === 0) return false;

  const initialScale = viewer.getScale();
  event.preventDefault();
  defer(() => {
    if (viewer.getScale() !== initialScale) return;
    try {
      Promise.resolve(
        viewer.setScale(zoomStepScale(initialScale, event.deltaY, event.deltaMode)),
      ).catch(onError);
    } catch (error) {
      onError?.(error);
    }
  });
  return true;
}

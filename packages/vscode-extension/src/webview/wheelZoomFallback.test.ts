import { describe, expect, it, vi } from 'vitest';
import type { ZoomableViewer } from '@silurus/ooxml-core';
import { captureUnhandledWheelZoom, type WheelZoomEvent } from './wheelZoomFallback.js';

function fakeViewer(initialScale = 1): ZoomableViewer {
  let scale = initialScale;
  return {
    getScale: () => scale,
    setScale: (next) => {
      scale = next;
    },
    zoomIn: vi.fn(),
    zoomOut: vi.fn(),
    fitWidth: vi.fn(),
    fitPage: vi.fn(),
  };
}

function fakeEvent(overrides: Partial<WheelZoomEvent> = {}): WheelZoomEvent {
  return {
    ctrlKey: true,
    metaKey: false,
    deltaY: -100,
    deltaMode: 0,
    preventDefault: vi.fn(),
    ...overrides,
  };
}

describe('VS Code webview wheel-zoom fallback', () => {
  it('zooms when the viewer-owned handler did not receive the pinch event', () => {
    const viewer = fakeViewer();
    const event = fakeEvent();
    let deferred: (() => void) | undefined;

    expect(
      captureUnhandledWheelZoom(event, viewer, (callback) => {
        deferred = callback;
      }),
    ).toBe(true);
    expect(event.preventDefault).toHaveBeenCalledOnce();
    expect(viewer.getScale()).toBe(1);

    deferred?.();
    expect(viewer.getScale()).toBeCloseTo(1.1, 10);
  });

  it('does not zoom twice when the scroll viewer already handled the event', () => {
    const viewer = fakeViewer();
    const event = fakeEvent();
    let deferred: (() => void) | undefined;

    captureUnhandledWheelZoom(event, viewer, (callback) => {
      deferred = callback;
    });
    viewer.setScale(1.1);
    deferred?.();

    expect(viewer.getScale()).toBe(1.1);
  });

  it('leaves an unmodified wheel event to native scrolling', () => {
    const viewer = fakeViewer();
    const event = fakeEvent({ ctrlKey: false });

    expect(captureUnhandledWheelZoom(event, viewer)).toBe(false);
    expect(event.preventDefault).not.toHaveBeenCalled();
    expect(viewer.getScale()).toBe(1);
  });

  it('accepts Command-modified wheel gestures', () => {
    const viewer = fakeViewer();
    let deferred: (() => void) | undefined;

    expect(
      captureUnhandledWheelZoom(
        fakeEvent({ ctrlKey: false, metaKey: true }),
        viewer,
        (callback) => {
          deferred = callback;
        },
      ),
    ).toBe(true);

    deferred?.();
    expect(viewer.getScale()).toBeGreaterThan(1);
  });

  it('preserves line-mode normalization in the webview fallback', () => {
    const viewer = fakeViewer();
    let deferred: (() => void) | undefined;

    captureUnhandledWheelZoom(
      fakeEvent({ deltaY: -3, deltaMode: 1 }),
      viewer,
      (callback) => {
        deferred = callback;
      },
    );
    deferred?.();

    expect(viewer.getScale()).toBeCloseTo(1.1, 10);
  });

  it('reports a synchronous viewer failure', () => {
    const error = new Error('viewer destroyed');
    const viewer = fakeViewer();
    viewer.setScale = () => {
      throw error;
    };
    const onError = vi.fn();
    let deferred: (() => void) | undefined;

    captureUnhandledWheelZoom(
      fakeEvent(),
      viewer,
      (callback) => {
        deferred = callback;
      },
      onError,
    );
    deferred?.();

    expect(onError).toHaveBeenCalledWith(error);
  });
});

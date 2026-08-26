import { afterEach, describe, expect, it, vi } from 'vitest';
import { XlsxViewer } from './viewer.js';
import { installDom, makeContainer } from './viewer-destroy-test-dom.js';
import type { XlsxChromeColors } from './types.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

interface ViewerPrivate {
  chromeColors: XlsxChromeColors;
  renderGutters(): void;
  scheduleRender(): void;
}

describe('XlsxViewer chrome theme', () => {
  it('re-reads inherited CSS variables and repaints Canvas-owned chrome once', () => {
    const doc = installDom();
    const values: Record<string, string> = {
    };
    let themeChanged: () => void = () => undefined;
    const observe = vi.fn();
    const disconnect = vi.fn();
    class ThemeObserver {
      constructor(callback: () => void) { themeChanged = callback; }
      observe = observe;
      disconnect = disconnect;
    }
    Object.assign(doc.defaultView, {
      getComputedStyle: () => ({
        getPropertyValue: (property: string) => values[property] ?? '',
      }),
      MutationObserver: ThemeObserver,
    });

    const viewer = new XlsxViewer(makeContainer() as unknown as HTMLElement);
    const subject = viewer as unknown as ViewerPrivate;
    subject.renderGutters = vi.fn();
    subject.scheduleRender = vi.fn();

    Object.assign(values, {
      '--ooxml-xlsx-chrome-background': '#101820',
      '--ooxml-xlsx-chrome-surface': '#182430',
      '--ooxml-xlsx-chrome-text': '#f5f7fa',
      '--ooxml-xlsx-chrome-border': '#52606d',
    });
    themeChanged();
    expect(subject.chromeColors).toEqual({
      background: '#101820',
      surface: '#182430',
      text: '#f5f7fa',
      border: '#52606d',
    });
    expect(subject.renderGutters).toHaveBeenCalledOnce();
    expect(subject.scheduleRender).toHaveBeenCalledOnce();

    themeChanged();
    expect(subject.renderGutters).toHaveBeenCalledOnce();
    expect(subject.scheduleRender).toHaveBeenCalledOnce();
    expect(observe).toHaveBeenCalled();
    viewer.destroy();
    expect(disconnect).toHaveBeenCalledOnce();
  });
});

import type { XlsxChromeColors } from '../../types.js';

const XLSX_CHROME_COLOR_PROPERTIES = {
  background: '--ooxml-xlsx-chrome-background',
  surface: '--ooxml-xlsx-chrome-surface',
  mutedSurface: '--ooxml-xlsx-chrome-surface-muted',
  text: '--ooxml-xlsx-chrome-text',
  mutedText: '--ooxml-xlsx-chrome-text-muted',
  border: '--ooxml-xlsx-chrome-border',
  selectedSurface: '--ooxml-xlsx-chrome-selection-background',
  accent: '--ooxml-xlsx-chrome-accent',
} as const satisfies Record<keyof XlsxChromeColors, string>;

function sameChromeColors(left: XlsxChromeColors, right: XlsxChromeColors): boolean {
  return Object.keys(XLSX_CHROME_COLOR_PROPERTIES).every((key) =>
    left[key as keyof XlsxChromeColors] === right[key as keyof XlsxChromeColors]);
}

/** Mount elements and repaint hook the theme tracker needs from its engine. */
export interface ChromeThemeHost {
  readonly hostWindow: Window & typeof globalThis;
  /** Caller-owned mount; it and its ancestors carry the application theme. */
  readonly container: HTMLElement;
  /** Viewer root whose computed style resolves the inherited variables. */
  readonly wrapper: HTMLElement;
  isDestroyed(): boolean;
  /** Repaint Canvas-owned chrome after the resolved colors changed. */
  onChange(): void;
}

/**
 * Tracks the `--ooxml-xlsx-chrome-*` CSS custom properties for chrome the
 * viewer paints into Canvas pixels (row/column headers, outline gutters). DOM
 * chrome follows inherited CSS variables without help; Canvas chrome needs an
 * explicit repaint, so this owns the MutationObserver on the mount's ancestor
 * chain and the `prefers-color-scheme` listener that trigger re-reads.
 */
export class ChromeTheme {
  colors: XlsxChromeColors = {};
  private styleObserver: MutationObserver | null = null;
  private schemeMedia: MediaQueryList | null = null;
  private schemeListener: (() => void) | null = null;

  constructor(private readonly host: ChromeThemeHost) {}

  /** Re-read the CSS custom properties that affect Canvas-painted chrome. */
  refresh(): void {
    if (this.host.isDestroyed()) return;
    const hostWindow = this.host.hostWindow;
    const getComputedStyle = hostWindow.getComputedStyle?.bind(hostWindow);
    if (!getComputedStyle) return;
    const computed = getComputedStyle(this.host.wrapper);
    // ast-grep-ignore: no-plain-object-string-map -- keys are the fixed chrome color property list.
    const next: Record<string, string> = {};
    for (const [key, property] of Object.entries(XLSX_CHROME_COLOR_PROPERTIES)) {
      const value = computed.getPropertyValue(property).trim();
      if (value) next[key] = value;
    }
    const nextColors = next as XlsxChromeColors;
    if (sameChromeColors(this.colors, nextColors)) return;
    this.colors = nextColors;
    this.host.onChange();
  }

  /** Read the initial colors and observe the ordinary ways an application
   *  changes theme state. */
  install(): void {
    this.refresh();
    const hostWindow = this.host.hostWindow;

    const MutationObserverClass = hostWindow.MutationObserver ?? globalThis.MutationObserver;
    if (MutationObserverClass) {
      this.styleObserver = new MutationObserverClass(() => this.refresh());
      for (let target: HTMLElement | null = this.host.container; target; target = target.parentElement) {
        this.styleObserver.observe(target, {
          attributes: true,
          attributeFilter: ['class', 'style', 'data-theme'],
        });
      }
    }

    const media = hostWindow.matchMedia?.('(prefers-color-scheme: dark)') ?? null;
    if (media) {
      const listener = () => this.refresh();
      media.addEventListener?.('change', listener);
      this.schemeMedia = media;
      this.schemeListener = listener;
    }
  }

  /** Stop observing theme changes. */
  destroy(): void {
    this.styleObserver?.disconnect();
    this.styleObserver = null;
    if (this.schemeMedia && this.schemeListener) {
      this.schemeMedia.removeEventListener?.('change', this.schemeListener);
    }
    this.schemeMedia = null;
    this.schemeListener = null;
  }
}

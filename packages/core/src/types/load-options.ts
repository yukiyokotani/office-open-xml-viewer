/**
 * Common load-time options shared by the docx / pptx / xlsx viewer
 * `load(source, opts)` methods.
 *
 * Each viewer narrows or extends this in its package-local `LoadOptions` if
 * needed, but the names that overlap (currently `useGoogleFonts`) keep the
 * same shape so application code can pass the same options object.
 */
export interface LoadOptions {
  /**
   * Opt in to loading webfont substitutes from Google Fonts
   * (`fonts.googleapis.com`). Default `false` — the canvas falls back to
   * locally available fonts.
   *
   * When enabled, end-user IP / User-Agent is sent to Google, which may
   * have privacy / GDPR implications for your application. To avoid the
   * third-party request, host the substitutes yourself and reference them
   * via `@font-face` in your application CSS.
   */
  useGoogleFonts?: boolean;
}

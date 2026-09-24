import * as vscode from 'vscode';

/** Origin that serves the Google Fonts *stylesheets* (`<link rel="stylesheet">`). */
const GOOGLE_FONTS_CSS_ORIGIN = 'https://fonts.googleapis.com';
/** Origin that serves the actual `.woff2` font binaries referenced by the CSS. */
const GOOGLE_FONTS_FILES_ORIGIN = 'https://fonts.gstatic.com';

/**
 * Build the webview Content-Security-Policy string.
 *
 * Pure function (no VSCode API) so it can be unit-tested for both states.
 *
 * When `useGoogleFonts` is false the policy is fully offline: the only allowed
 * origin is the extension's own `cspSource`. When true we widen exactly two
 * directives, no more:
 *   - `style-src` gains {@link GOOGLE_FONTS_CSS_ORIGIN} because the library
 *     loads each Google Fonts CSS via an injected `<link rel="stylesheet">`.
 *   - `font-src` gains {@link GOOGLE_FONTS_FILES_ORIGIN} because the `@font-face`
 *     rules in that CSS point their `src:` at `fonts.gstatic.com` woff2 files.
 * `connect-src` is deliberately NOT widened: the preload path
 * (`packages/core/src/fonts/preload.ts`) never `fetch()`es either origin — the
 * browser font engine fetches the binaries, governed by `font-src`.
 */
export function buildContentSecurityPolicy(
  cspSource: string,
  nonce: string,
  useGoogleFonts: boolean,
): string {
  const fontSrc = useGoogleFonts
    ? `font-src ${cspSource} ${GOOGLE_FONTS_FILES_ORIGIN};`
    : `font-src ${cspSource};`;
  const styleSrc = useGoogleFonts
    ? `style-src 'unsafe-inline' ${GOOGLE_FONTS_CSS_ORIGIN};`
    : `style-src 'unsafe-inline';`;

  return [
    `default-src 'none';`,
    `base-uri ${cspSource};`,
    `img-src ${cspSource} data: blob:;`,
    `media-src ${cspSource} blob:;`,
    fontSrc,
    `script-src 'nonce-${nonce}' 'wasm-unsafe-eval';`,
    `worker-src data: blob:;`,
    styleSrc,
    `connect-src ${cspSource} data: blob:;`,
  ].join(' ');
}

/**
 * Generate the HTML for the webview panel.
 * The webview script (dist/webview.js) is allowed via the content security policy,
 * and receives the file bytes via a `ooxml-init` message posted from the extension host.
 *
 * When `useGoogleFonts` is true the CSP is widened to allow the optional
 * font CDN (see {@link buildContentSecurityPolicy}); the flag is also forwarded to
 * the viewers via the `ooxml-init` message in the editor providers.
 */
export function getWebviewHtml(
  webview: vscode.Webview,
  extensionUri: vscode.Uri,
  fileType: 'docx' | 'xlsx' | 'pptx',
  useGoogleFonts = false,
  selectionSession = 0,
): string {
  const scriptUri = webview.asWebviewUri(
    vscode.Uri.joinPath(extensionUri, 'dist', 'webview.js'),
  );
  // esbuild emits parser WASM assets as relative URLs. The webview document
  // has its own origin, so resolve those URLs beside webview.js.
  const assetBaseUri = new URL('.', scriptUri.toString()).toString();

  const nonce = getNonce();
  const csp = buildContentSecurityPolicy(webview.cspSource, nonce, useGoogleFonts);

  return /* html */ `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <meta http-equiv="Content-Security-Policy" content="${csp}" />
  <base href="${assetBaseUri}" />
  <title>OOXML Viewer</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    html, body {
      width: 100%;
      height: 100%;
      overflow: hidden;
      background: var(--vscode-editor-background);
      color: var(--vscode-foreground);
      font-family: var(--vscode-font-family, sans-serif);
    }
    #viewer-root {
      position: relative;
      width: 100%;
      height: 100%;
      min-height: 0;
      display: flex;
      flex-direction: column;
    }
    #viewer-toolbar {
      height: 35px;
      flex: 0 0 35px;
      display: flex;
      align-items: center;
      justify-content: flex-end;
      gap: 4px;
      padding: 0 8px;
      border-bottom: 1px solid var(--vscode-panel-border, rgba(127, 127, 127, 0.35));
      background: var(--vscode-editorGroupHeader-tabsBackground, var(--vscode-editor-background));
      user-select: none;
    }
    #viewer-toolbar button {
      width: 26px;
      height: 24px;
      border: 1px solid transparent;
      border-radius: 3px;
      color: var(--vscode-foreground);
      background: transparent;
      font: 16px/1 var(--vscode-font-family, sans-serif);
      cursor: pointer;
    }
    #viewer-toolbar button:hover:not(:disabled) {
      background: var(--vscode-toolbar-hoverBackground, rgba(127, 127, 127, 0.2));
    }
    #viewer-toolbar button:focus-visible {
      outline: 1px solid var(--vscode-focusBorder);
      outline-offset: -1px;
    }
    #viewer-toolbar button:disabled {
      cursor: default;
      opacity: 0.4;
    }
    #zoom-label {
      min-width: 46px;
      text-align: center;
      font-size: 12px;
      font-variant-numeric: tabular-nums;
    }
    #viewer-container {
      position: relative;
      width: 100%;
      min-width: 0;
      min-height: 0;
      flex: 1 1 auto;
      overflow: hidden;
    }
    #find-popup {
      position: absolute;
      z-index: 20;
      top: 43px;
      right: 12px;
      display: flex;
      align-items: center;
      gap: 4px;
      padding: 6px;
      border: 1px solid var(--vscode-widget-border, var(--vscode-panel-border));
      border-radius: 4px;
      background: var(--vscode-editorWidget-background, var(--vscode-editor-background));
      color: var(--vscode-editorWidget-foreground, var(--vscode-foreground));
      box-shadow: 0 2px 8px var(--vscode-widget-shadow, rgba(0, 0, 0, 0.35));
    }
    #find-popup[hidden] {
      display: none;
    }
    #find-input {
      width: min(240px, 42vw);
      height: 24px;
      padding: 2px 6px;
      border: 1px solid var(--vscode-input-border, transparent);
      outline: none;
      background: var(--vscode-input-background);
      color: var(--vscode-input-foreground);
      font: 12px/1.4 var(--vscode-font-family, sans-serif);
    }
    #find-input:focus {
      border-color: var(--vscode-focusBorder);
    }
    #find-status {
      min-width: 70px;
      text-align: center;
      white-space: nowrap;
      font-size: 12px;
      font-variant-numeric: tabular-nums;
    }
    #find-popup button {
      width: 24px;
      height: 24px;
      border: 1px solid transparent;
      border-radius: 3px;
      color: inherit;
      background: transparent;
      font: 14px/1 var(--vscode-font-family, sans-serif);
      cursor: pointer;
    }
    #find-popup button:hover:not(:disabled) {
      background: var(--vscode-toolbar-hoverBackground, rgba(127, 127, 127, 0.2));
    }
    #find-popup button:focus-visible {
      outline: 1px solid var(--vscode-focusBorder);
      outline-offset: -1px;
    }
    #find-popup button:disabled {
      cursor: default;
      opacity: 0.4;
    }
    #status {
      position: absolute;
      inset: 0;
      display: flex;
      align-items: center;
      justify-content: center;
      pointer-events: none;
      z-index: 10;
    }
    #status[data-state="error"] {
      pointer-events: auto;
      color: var(--vscode-errorForeground, #f44747);
      font-size: 13px;
      padding: 16px;
      text-align: center;
    }
    .spinner {
      width: 28px;
      height: 28px;
      border: 3px solid color-mix(in srgb, var(--vscode-foreground) 20%, transparent);
      border-top-color: var(--vscode-progressBar-background, var(--vscode-foreground));
      border-radius: 50%;
      animation: spin 0.9s linear infinite;
    }
    @keyframes spin { to { transform: rotate(360deg); } }
  </style>
</head>
<body>
  <div id="viewer-root">
    <div id="viewer-toolbar" role="toolbar" aria-label="Zoom controls">
      <button id="zoom-out" type="button" aria-label="Zoom out" title="Zoom out" disabled>−</button>
      <span id="zoom-label" aria-live="polite">100%</span>
      <button id="zoom-in" type="button" aria-label="Zoom in" title="Zoom in" disabled>+</button>
    </div>
    <div id="find-popup" role="dialog" aria-label="Find" hidden>
      <input id="find-input" type="text" aria-label="Find" placeholder="Find" autocomplete="off" spellcheck="false" />
      <span id="find-status" aria-live="polite"></span>
      <button id="find-previous" type="button" aria-label="Previous match" title="Previous match (Shift+Enter)" disabled>↑</button>
      <button id="find-next" type="button" aria-label="Next match" title="Next match (Enter)" disabled>↓</button>
      <button id="find-close" type="button" aria-label="Close find" title="Close (Escape)">×</button>
    </div>
    <div id="viewer-container">
      <div id="status"><div class="spinner"></div></div>
    </div>
  </div>
  <script nonce="${nonce}">
    window.__OOXML_FILE_TYPE__ = ${JSON.stringify(fileType)};
    window.__OOXML_SELECTION_SESSION__ = ${JSON.stringify(selectionSession)};
  </script>
  <script nonce="${nonce}" src="${scriptUri}"></script>
</body>
</html>`;
}

function getNonce(): string {
  let text = '';
  const possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  for (let i = 0; i < 32; i++) {
    text += possible.charAt(Math.floor(Math.random() * possible.length));
  }
  return text;
}

import * as vscode from 'vscode';

/** Configuration section root for all extension settings. */
export const CONFIG_SECTION = 'ooxmlViewer';

/** Setting id (relative to {@link CONFIG_SECTION}) for the Google Fonts opt-in. */
export const USE_GOOGLE_FONTS_SETTING = 'useGoogleFonts';

/** Fully-qualified id, e.g. for `event.affectsConfiguration(...)`. */
export const USE_GOOGLE_FONTS_CONFIG_ID = `${CONFIG_SECTION}.${USE_GOOGLE_FONTS_SETTING}`;

/**
 * Resolve whether the webview may load optional font substitutes from the Google
 * Fonts CDN.
 *
 * Two gates, both must pass:
 *   1. The user opted in via `ooxmlViewer.useGoogleFonts` (default `false`, so the
 *      extension is offline out of the box).
 *   2. The workspace is *trusted*. Loading remote fonts is a network egress, which
 *      Workspace Trust governs — an untrusted/restricted workspace must never reach
 *      out to a CDN regardless of the setting value. `isTrusted` is `true` for
 *      windows without a folder (e.g. a single loose file), which is the desired
 *      behaviour: the user is in control there.
 */
export function shouldUseGoogleFonts(): boolean {
  if (!vscode.workspace.isTrusted) return false;
  return vscode.workspace
    .getConfiguration(CONFIG_SECTION)
    .get<boolean>(USE_GOOGLE_FONTS_SETTING, false);
}

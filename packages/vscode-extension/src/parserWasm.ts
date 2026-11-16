import * as vscode from 'vscode';

/** Written by `parserWasmManifest` in esbuild-asset-sidecars.mjs. */
const MANIFEST = 'parser-wasm.json';
const ASSET_PATH = /^assets\/[A-Za-z0-9_-]+\.wasm$/;

/**
 * Read the parser WASM for `fileType` from the extension's own `dist`.
 *
 * The extension host posts these bytes to the webview rather than letting the
 * webview fetch them (#2212). A webview fetch of an extension resource is
 * served by VS Code's resource service worker, which answers 408 when the
 * workbench does not deliver the file within its timeout; this was observed
 * over Remote SSH. The host reads the file wherever the extension runs (the
 * remote machine for a remote workspace), and the bytes travel over the same
 * message channel as the document itself.
 */
export async function readParserWasm(
  extensionUri: vscode.Uri,
  fileType: 'docx' | 'xlsx' | 'pptx',
): Promise<Uint8Array> {
  const dist = vscode.Uri.joinPath(extensionUri, 'dist');
  const manifest: unknown = JSON.parse(
    new TextDecoder().decode(await vscode.workspace.fs.readFile(vscode.Uri.joinPath(dist, MANIFEST))),
  );
  const asset = (manifest as Record<string, unknown> | null)?.[fileType];
  if (typeof asset !== 'string' || !ASSET_PATH.test(asset)) {
    throw new Error(`The ${fileType} parser WebAssembly is missing from the extension build`);
  }
  // Copy into a plain, exact-length Uint8Array (see the document read in
  // baseEditor.ts): the webview transport cannot carry a Node Buffer.
  return new Uint8Array(await vscode.workspace.fs.readFile(vscode.Uri.joinPath(dist, ...asset.split('/'))));
}

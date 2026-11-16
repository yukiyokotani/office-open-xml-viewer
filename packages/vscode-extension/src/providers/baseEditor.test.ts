import { describe, expect, it, vi } from 'vitest';

const mocks = vi.hoisted(() => ({
  readFile: vi.fn(),
}));

vi.mock('vscode', () => ({
  Uri: {
    joinPath: (base: { path: string }, ...parts: string[]) => ({
      path: [base.path, ...parts].join('/'),
    }),
  },
  workspace: {
    fs: { readFile: mocks.readFile },
  },
}));

vi.mock('../webviewHtml', () => ({ getWebviewHtml: () => '<html></html>' }));
vi.mock('../config', () => ({ shouldUseGoogleFonts: () => false }));

import { BaseEditorProvider } from './baseEditor';
import { SelectionContextRegistry } from '../selectionContextRegistry';
import { exactArrayBuffer } from '../webview/postedBytes';

/**
 * Mirrors VS Code's webview message transport: `JSON.stringify` with a
 * replacer that moves ArrayBuffer views out as transferable references, then
 * revives them as views over the transferred buffers. `JSON.stringify` calls
 * `Buffer#toJSON` before the replacer sees the value, so a Node Buffer
 * arrives as `{ type: 'Buffer', data: [...] }`, exactly as in the editor.
 */
function webviewRoundTrip(message: unknown): Record<string, unknown> {
  const views: ArrayBufferView[] = [];
  const json = JSON.stringify(message, (_key, value: unknown) => {
    if (ArrayBuffer.isView(value)) {
      views.push(value);
      return { $$vscode_array_buffer_reference$$: views.length - 1 };
    }
    return value;
  });
  const transferred = views.map((view) => {
    const copy = new ArrayBuffer(view.buffer.byteLength);
    new Uint8Array(copy).set(new Uint8Array(view.buffer as ArrayBuffer));
    return new Uint8Array(copy, view.byteOffset, view.byteLength);
  });
  return JSON.parse(json, (_key, value: unknown) => {
    const index = (value as { $$vscode_array_buffer_reference$$?: number } | null)
      ?.$$vscode_array_buffer_reference$$;
    return typeof index === 'number' ? transferred[index] : value;
  });
}

const WASM = new Uint8Array([0x00, 0x61, 0x73, 0x6d, 1, 0, 0, 0]);

/** Route reads: the build manifest and parser WASM under /ext/dist, else the document. */
function mockReads(document: () => Promise<Uint8Array>, manifest: Record<string, string> = {
  docx: 'assets/docx_parser_bg-ABC123.wasm',
}) {
  mocks.readFile.mockReset().mockImplementation(async (uri: { path: string }) => {
    if (uri.path === '/ext/dist/parser-wasm.json') {
      return new TextEncoder().encode(JSON.stringify(manifest));
    }
    // A pooled Node Buffer view, as `workspace.fs.readFile` may return.
    if (uri.path === '/ext/dist/assets/docx_parser_bg-ABC123.wasm') {
      return Buffer.from([9, ...WASM, 9]).subarray(1, 1 + WASM.length);
    }
    if (uri.path.startsWith('/ext/')) throw new Error(`unexpected extension read ${uri.path}`);
    return document();
  });
}

class TestProvider extends BaseEditorProvider {
  protected readonly fileType = 'docx' as const;
}

function fakePanel() {
  let onMessage: ((msg: unknown) => Promise<void>) | undefined;
  const posted: Record<string, unknown>[] = [];
  const webview = {
    options: {} as { localResourceRoots?: { path: string }[] },
    html: '',
    cspSource: 'vscode-webview://test',
    asWebviewUri: vi.fn((uri: { path: string }) => ({ toString: () => `webview:${uri.path}` })),
    onDidReceiveMessage: (listener: (msg: unknown) => Promise<void>) => {
      onMessage = listener;
      return { dispose: () => undefined };
    },
    postMessage: vi.fn(async (msg: Record<string, unknown>) => {
      posted.push(msg);
      return true;
    }),
  };
  const panel = {
    webview,
    active: true,
    onDidDispose: () => ({ dispose: () => undefined }),
    onDidChangeViewState: () => ({ dispose: () => undefined }),
  };
  return { panel, webview, posted, send: (msg: unknown) => onMessage!(msg) };
}

describe('BaseEditorProvider webview resource scope', () => {
  it('exposes only bundled assets and posts the document bytes itself', async () => {
    const bytes = new Uint8Array([0x50, 0x4b, 0x03, 0x04]);
    mockReads(async () => bytes);
    const provider = new TestProvider(
      { extensionUri: { path: '/ext' } } as never,
      new SelectionContextRegistry(),
    );
    const documentUri = {
      scheme: 'file',
      path: '/docs/report.docx',
      fsPath: '/docs/report.docx',
      toString: () => 'file:///docs/report.docx',
    };
    const { panel, webview, posted, send } = fakePanel();

    await provider.resolveCustomEditor({ uri: documentUri } as never, panel as never);

    // Neither the document's directory nor the document is loadable.
    expect(webview.options.localResourceRoots).toEqual([{ path: '/ext/dist' }]);
    expect(webview.asWebviewUri).not.toHaveBeenCalled();

    await send({ type: 'webview-ready', selectionSession: 1 });
    expect(mocks.readFile).toHaveBeenCalledWith(documentUri);
    expect(posted).toHaveLength(1);
    expect(posted[0]).toMatchObject({
      type: 'ooxml-init',
      fileType: 'docx',
      selectionSession: 1,
    });
    expect(posted[0].data).toEqual(bytes);
    // #2212: the parser WASM is posted, so the webview never fetches it.
    expect([...(posted[0].wasm as Uint8Array)]).toEqual([...WASM]);
    expect(posted[0]).not.toHaveProperty('url');
  });

  it('posts a Node Buffer view as bytes that survive the webview transport', async () => {
    const pool = Buffer.from([9, 9, 0x50, 0x4b, 0x03, 0x04, 9]);
    const pooled = pool.subarray(2, 6);
    expect(pooled.byteOffset).toBeGreaterThan(0);
    expect(pooled.buffer.byteLength).toBeGreaterThan(pooled.byteLength);
    // The transport itself cannot carry a Buffer.
    expect(exactArrayBuffer(webviewRoundTrip({ data: pooled }).data)).toBeNull();

    mockReads(async () => pooled);
    const provider = new TestProvider(
      { extensionUri: { path: '/ext' } } as never,
      new SelectionContextRegistry(),
    );
    const { panel, posted, send } = fakePanel();
    await provider.resolveCustomEditor(
      { uri: { scheme: 'file', path: '/docs/x.docx', fsPath: '/docs/x.docx' } } as never,
      panel as never,
    );
    await send({ type: 'webview-ready', selectionSession: 1 });

    const message = webviewRoundTrip(posted[0]);
    const received = exactArrayBuffer(message.data);
    expect(received).not.toBeNull();
    expect([...new Uint8Array(received as ArrayBuffer)]).toEqual([0x50, 0x4b, 0x03, 0x04]);
    const wasm = exactArrayBuffer(message.wasm);
    expect(wasm).not.toBeNull();
    expect([...new Uint8Array(wasm as ArrayBuffer)]).toEqual([...WASM]);
  });

  it('reports a read failure to the webview instead of leaving it loading', async () => {
    mockReads(async () => { throw new Error('EACCES'); });
    const provider = new TestProvider(
      { extensionUri: { path: '/ext' } } as never,
      new SelectionContextRegistry(),
    );
    const { panel, posted, send } = fakePanel();

    await provider.resolveCustomEditor(
      { uri: { scheme: 'file', path: '/docs/x.docx', fsPath: '/docs/x.docx' } } as never,
      panel as never,
    );
    await send({ type: 'webview-ready', selectionSession: 1 });

    expect(posted).toHaveLength(1);
    expect(posted[0]).toMatchObject({ type: 'ooxml-init', error: 'EACCES' });
    expect(posted[0].data).toBeUndefined();
  });

  it('reports a parser WASM missing from the build instead of posting the document alone', async () => {
    mockReads(async () => new Uint8Array([0x50, 0x4b]), { pptx: 'assets/pptx_parser_bg-ABC123.wasm' });
    const provider = new TestProvider(
      { extensionUri: { path: '/ext' } } as never,
      new SelectionContextRegistry(),
    );
    const { panel, posted, send } = fakePanel();

    await provider.resolveCustomEditor(
      { uri: { scheme: 'file', path: '/docs/x.docx', fsPath: '/docs/x.docx' } } as never,
      panel as never,
    );
    await send({ type: 'webview-ready', selectionSession: 1 });

    expect(posted).toHaveLength(1);
    expect(posted[0].error).toMatch(/docx parser WebAssembly is missing/);
    expect(posted[0].data).toBeUndefined();
    expect(posted[0].wasm).toBeUndefined();
  });
});

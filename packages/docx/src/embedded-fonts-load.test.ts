import { readFileSync } from 'node:fs';
import { afterEach, beforeAll, expect, it, vi } from 'vitest';
import { OoxmlResourceLimitError } from '@silurus/ooxml-core';
import { storedZip } from '@silurus/ooxml-core/testing';
import { resourcePolicyForWasm, serializeWorkerError } from '@silurus/ooxml-core/worker';
import { DocxDocument } from './document.js';
import { DocumentPullWorker, isDocumentPullCommand } from './document-pull-worker.js';
import type { WorkerRequest } from './types.js';
import { initSync, DocxArchive } from './wasm/docx_parser.js';

vi.mock('./worker.ts?worker&inline', () => ({ default: function () { return new ArchiveWorker(); } }));

// Keep the production document cursor and bridge: the failure must cross the
// public load boundary after model admission, rather than just reject a helper.
class ArchiveWorker {
  static latest: ArchiveWorker;
  terminated = false;
  extracted: string[] = [];
  private archive?: DocxArchive;
  private pull = new DocumentPullWorker(() => this.archive);
  private listeners = new Set<(event: MessageEvent) => void>();
  constructor() { ArchiveWorker.latest = this; }
  addEventListener(type: string, listener: (event: MessageEvent) => void) {
    if (type === 'message') this.listeners.add(listener);
  }
  removeEventListener(_type: string, listener: (event: MessageEvent) => void) { this.listeners.delete(listener); }
  postMessage(request: WorkerRequest) {
    queueMicrotask(() => { void this.dispatch(request); });
  }
  private async dispatch(request: WorkerRequest) {
    try {
      if (isDocumentPullCommand(request)) {
        await this.pull.dispatch(request, response => this.reply(response));
      } else if (request.type === 'parse') {
        this.archive = new DocxArchive(new Uint8Array(request.data), ...resourcePolicyForWasm(request.resourcePolicy));
        const identity = { sessionId: 1, operationId: 1, generation: 1 };
        this.pull.open(identity);
        this.reply({ type: 'documentSessionOpened', id: request.id, ...identity });
      } else if (request.type === 'extractImage') {
        this.extracted.push(request.path);
        this.reply({ type: 'imageExtracted', id: request.id, bytes: this.archive!.extract_image(request.path).buffer });
      }
    } catch (error) {
      this.reply({ type: 'error', id: 'id' in request ? request.id : undefined, ...serializeWorkerError(error) });
    }
  }
  extractFont(path: string) { this.extracted.push(path); return this.archive!.extract_image(path); }
  private reply(data: unknown) { for (const listener of this.listeners) listener({ data } as MessageEvent); }
  terminate() { this.terminated = true; this.archive?.free(); this.archive = undefined; }
}

beforeAll(() => {
  initSync({ module: readFileSync(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});
afterEach(() => { ArchiveWorker.latest?.terminate(); vi.unstubAllGlobals(); vi.restoreAllMocks(); });

it.each([false, true])('rejects load when lazy font extraction crosses the entry limit (direct WASM=%s)', async direct => {
  if (direct) vi.spyOn(DocxDocument.prototype, 'getFontBytes').mockImplementation(async path => ArchiveWorker.latest.extractFont(path));
  const add = vi.fn();
  vi.stubGlobal('document', { fonts: { add, delete: vi.fn() } });
  vi.stubGlobal('FontFace', class {
    constructor(readonly family: string) {}
    load() { return Promise.resolve(this); }
  });
  vi.stubGlobal('location', { href: 'http://localhost/' });
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const files: Array<[string, string]> = [
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>'],
    ['word/document.xml', `<w:document xmlns:w="${w}"><w:body><w:p/></w:body></w:document>`],
    ['word/fontTable.xml', `<w:fonts xmlns:w="${w}" xmlns:r="${rel}">${[1, 2, 3].map(i => `<w:font w:name="Font ${i}"><w:embedRegular r:id="f${i}" w:fontKey="{00112233-4455-6677-8899-AABBCCDDEEFF}"/></w:font>`).join('')}</w:fonts>`],
    ['word/_rels/fontTable.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${[1, 2, 3].map(i => `<Relationship Id="f${i}" Type="${rel}/font" Target="fonts/font${i}.ttf"/>`).join('')}</Relationships>`],
    ['word/fonts/font1.ttf', 'first'], ['word/fonts/font2.ttf', 'second'],
    ['word/fonts/font3.ttf', 'x'.repeat(4097)],
  ];
  const loading = DocxDocument.load(storedZip(files).buffer, { resourceLimits: { maxArchiveEntryBytes: 4096 } });
  await expect(loading).rejects.toBeInstanceOf(OoxmlResourceLimitError);
  await expect(loading).rejects.toMatchObject({ details: { violation: {
    resource: 'archive-entry', metric: 'declared-inflated-bytes', limit: 4096, observed: 4097,
  } } });
  expect(ArchiveWorker.latest.extracted).toEqual([
    'word/fonts/font1.ttf', 'word/fonts/font2.ttf', 'word/fonts/font3.ttf',
  ]);
  // DOCX prepares all resources before registration, so no partial set escapes.
  // (The document's own route sentinels are not embedded fonts.)
  expect(add.mock.calls.filter(([face]) => !/^__ooxml_route_sentinel_/.test((face as { family: string }).family))).toEqual([]);
  expect(ArchiveWorker.latest.terminated).toBe(true);
});

import { afterEach, describe, expect, it, vi } from 'vitest';
import type { ModelSource, ModelSourceModuleDescriptor } from '@silurus/ooxml-core';
import { buildCfbFixture } from '@silurus/ooxml-core/testing';
import { XlsxWorkbook } from './workbook.js';
import { XLSX_HOST_LAYOUT_REQUEST, XLSX_HOST_LAYOUT_RESULT } from './internal/host-layout.js';
import type { ParsedWorkbook } from './types.js';

const mocks = vi.hoisted(() => ({ computeMdw: vi.fn(() => 8) }));
vi.mock('./renderer.js', async (load) => ({
  ...await load<typeof import('./renderer.js')>(),
  computeMdw: mocks.computeMdw,
}));

type Message = Record<string, unknown>;
type Script = (worker: ProtocolWorker, message: Message) => void;

/** A worker that answers the XLSX parse protocol from a test script. */
class ProtocolWorker {
  static instances: ProtocolWorker[] = [];
  static script: Script = () => undefined;
  terminated = false;
  readonly messages: Message[] = [];
  readonly transfers: (Transferable[] | undefined)[] = [];
  private readonly listeners = new Set<(event: MessageEvent) => void>();
  constructor() { ProtocolWorker.instances.push(this); }
  postMessage(message: unknown, transfer?: Transferable[]): void {
    this.messages.push(message as Message);
    this.transfers.push(transfer);
    const script = ProtocolWorker.script;
    queueMicrotask(() => script(this, message as Message));
  }
  addEventListener(type: string, listener: (event: MessageEvent) => void): void {
    if (type === 'message') this.listeners.add(listener);
  }
  removeEventListener(type: string, listener: (event: MessageEvent) => void): void {
    if (type === 'message') this.listeners.delete(listener);
  }
  reply(data: unknown): void {
    for (const listener of [...this.listeners]) listener({ data } as MessageEvent);
  }
  terminate(): void { this.terminated = true; }
}

const globals = globalThis as Record<string, unknown>;
const originals = { Worker: globals.Worker, location: globals.location };
const descriptor: ModelSourceModuleDescriptor = {
  protocol: 'ooxml-model-source-module/v1',
  target: 'xlsx',
  moduleUrl: 'https://example.test/source.mjs',
  config: {},
};
const CALIBRI_11 = { family: 'Calibri', sizePt: 11, bold: false, italic: false };
const cfbBytes = () => buildCfbFixture(['Root Entry', 'Workbook']);
const workbookJson = () => new TextEncoder().encode(JSON.stringify({
  workbook: { sheets: [{ name: 'Sheet1' }] }, styles: {}, sharedStrings: [],
})).buffer;

/**
 * The parse worker's side of a model-source load: ask the page to measure the
 * Normal font, then answer the parse with the width the page returned.
 */
function parseWorkerScript(): Script {
  let pendingParse: number | undefined;
  return (worker, message) => {
    if (message.type === 'parse') {
      pendingParse = message.id as number;
      worker.reply({ type: XLSX_HOST_LAYOUT_REQUEST, requestId: 41, font: CALIBRI_11 });
    } else if (message.type === XLSX_HOST_LAYOUT_RESULT && pendingParse !== undefined) {
      const width = message.maximumDigitWidth as number | undefined;
      worker.reply({
        type: 'parsed', id: pendingParse, workbookJson: workbookJson(), usage: undefined,
        ...(width === undefined ? {} : { layoutMetrics: { maximumDigitWidth: width } }),
      });
    } else if (message.type === 'toMarkdown') {
      worker.reply({
        type: 'error', id: message.id, name: 'Error',
        message: 'Markdown conversion is unsupported for this source',
      });
    }
  };
}

function fakeSource(claim = true) {
  const release = vi.fn();
  const beginLoad = vi.fn(() => ({ module: descriptor, release }));
  const source = { target: 'xlsx', claim: vi.fn(() => claim), beginLoad } as unknown as ModelSource<'xlsx'>;
  return { source, beginLoad, release };
}

afterEach(() => {
  globals.Worker = originals.Worker;
  globals.location = originals.location;
  ProtocolWorker.instances = [];
  mocks.computeMdw.mockClear();
});

function install(script: Script): void {
  globals.Worker = ProtocolWorker;
  globals.location = { href: 'http://localhost/' };
  ProtocolWorker.script = script;
}

describe('XlsxWorkbook.load with model sources', () => {
  it('forwards a claimed load and answers its host layout request with the renderer measurement', async () => {
    install(parseWorkerScript());
    const { source, release } = fakeSource();
    const workbook = await XlsxWorkbook.load(cfbBytes(), { modelSources: [source] });

    const worker = ProtocolWorker.instances[0]!;
    expect(worker.messages.filter((message) => message.type === 'parse')).toEqual([
      expect.objectContaining({ source: descriptor }),
    ]);
    expect(mocks.computeMdw).toHaveBeenCalledExactlyOnceWith('Calibri', 11, undefined, false, 400, 'normal');
    expect(worker.messages).toContainEqual({
      type: XLSX_HOST_LAYOUT_RESULT, requestId: 41, maximumDigitWidth: 8,
    });
    // The parse response's width becomes the workbook's authoritative MDW input.
    expect((workbook as unknown as { parsedWorkbook: ParsedWorkbook }).parsedWorkbook.layoutMetrics)
      .toEqual({ maximumDigitWidth: 8 });
    expect(workbook.sheetNames).toEqual(['Sheet1']);
    expect(release).toHaveBeenCalledOnce();
    await expect(workbook.toMarkdown()).rejects.toThrow('Markdown conversion is unsupported for this source');
    workbook.destroy();
  });

  it('keeps unclaimed input on the OOXML path and releases a failed load once', async () => {
    install(() => undefined);
    const unclaimed = fakeSource(false);
    await expect(XlsxWorkbook.load(cfbBytes(), { modelSources: [unclaimed.source] }))
      .rejects.toMatchObject({ code: 'legacy-binary-format' });
    expect(unclaimed.beginLoad).not.toHaveBeenCalled();
    expect(ProtocolWorker.instances).toEqual([]);

    install((worker, message) => {
      if (message.type === 'parse') {
        worker.reply({ type: 'error', id: message.id, name: 'Error', message: 'source parse failed' });
      }
    });
    const failing = fakeSource();
    await expect(XlsxWorkbook.load(cfbBytes(), { modelSources: [failing.source] }))
      .rejects.toThrow('source parse failed');
    expect(failing.release).toHaveBeenCalledOnce();
    expect(ProtocolWorker.instances[0]!.terminated).toBe(true);
  });
});

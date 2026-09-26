// Test-only XLSX model source module: wraps the real XLSX parser archive and
// optionally asks the host to measure a Normal font (host layout).
//
// config.hostLayoutFamily (string | null)  font family for host_layout_request()
// config.hostLayoutSizePt (number)         font size for host_layout_request()
// config.minimal          (boolean)        expose only the required archive methods
// transfer[0]             (ArrayBuffer)    probe Float64Array: [close() calls,
//                                          configure_host_layout() calls, last width or NaN]
import { readFileSync } from 'node:fs';
import { initSync, XlsxArchive } from '../../../xlsx/src/wasm/xlsx_parser.js';

const REQUIRED = [
  'open_sheet_cursor', 'pull_sheet_cursor', 'sheet_cursor_pull_finished',
  'acknowledge_sheet_cursor_terminal', 'cancel_sheet_cursor',
  'close_sheet_cursor', 'assert_healthy', 'parse', 'extract_image',
];

export async function openModelSource(bytes, config, _signal, transfer = []) {
  initSync({ module: readFileSync(new URL('../../../xlsx/src/wasm/xlsx_parser_bg.wasm', import.meta.url)) });
  const parser = new XlsxArchive(bytes);
  const probe = transfer[0] instanceof ArrayBuffer ? new Float64Array(transfer[0]) : undefined;
  const archive = config.minimal || typeof config.hostLayoutFamily === 'string'
    ? Object.fromEntries(REQUIRED.map((method) => [method, parser[method].bind(parser)]))
    : parser;
  if (typeof config.hostLayoutFamily === 'string') {
    const font = { family: config.hostLayoutFamily, sizePt: config.hostLayoutSizePt, bold: false, italic: false };
    archive.host_layout_request = () => new TextEncoder().encode(JSON.stringify(font));
    archive.configure_host_layout = (width) => {
      if (!probe) return;
      probe[1] += 1;
      probe[2] = width === undefined ? Number.NaN : width;
    };
  }
  return {
    archive,
    close() {
      if (probe) probe[0] += 1;
      parser.free();
    },
  };
}

// Test-only PPTX model source module: wraps the real PPTX parser archive.
//
// config.minimal (boolean)      expose only the required archive methods
// transfer[0]    (ArrayBuffer)  probe: Float64Array[0] counts close() calls
import { readFileSync } from 'node:fs';
import { initSync, PptxArchive } from '../../../pptx/src/wasm/pptx_parser.js';

const REQUIRED = [
  'pull_slide', 'slide_cursor_resource_usage', 'acknowledge_slide', 'cancel_slide',
  'close_presentation_session', 'assert_healthy', 'presentation_bootstrap', 'extract_image',
];

export async function openModelSource(bytes, config, _signal, transfer = []) {
  initSync({ module: readFileSync(new URL('../../../pptx/src/wasm/pptx_parser_bg.wasm', import.meta.url)) });
  const parser = new PptxArchive(bytes);
  const probe = transfer[0] instanceof ArrayBuffer ? new Float64Array(transfer[0]) : undefined;
  return {
    archive: config.minimal
      ? Object.fromEntries(REQUIRED.map((method) => [method, parser[method].bind(parser)]))
      : parser,
    close() {
      if (probe) probe[0] += 1;
      parser.free();
    },
  };
}

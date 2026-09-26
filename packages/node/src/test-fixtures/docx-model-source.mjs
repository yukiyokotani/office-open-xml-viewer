// Test-only DOCX model source module: opens the input with the real DOCX parser
// archive, so a fake ModelSource can exercise the generic contract end to end.
//
// config.showTrackedChanges (boolean | null)  reported view default
// config.extraViewDefault   (string | null)   an additional, unsupported view key
// config.minimal            (boolean)         expose only the required archive methods
// transfer[0]               (ArrayBuffer)     probe: Float64Array[0] counts close() calls
import { readFileSync } from 'node:fs';
import { DocxArchive, initSync } from '../../../docx/src/wasm/docx_parser.js';

const REQUIRED = [
  'open_document_cursor', 'pull_document_chunk', 'document_chunk_done',
  'acknowledge_document_chunk', 'cancel_document_cursor', 'close_document_session',
  'assert_healthy', 'extract_image',
];

export async function openModelSource(bytes, config, _signal, transfer = []) {
  initSync({ module: readFileSync(new URL('../../../docx/src/wasm/docx_parser_bg.wasm', import.meta.url)) });
  const parser = new DocxArchive(bytes);
  const probe = transfer[0] instanceof ArrayBuffer ? new Float64Array(transfer[0]) : undefined;
  const viewDefaults = {};
  if (typeof config.showTrackedChanges === 'boolean') viewDefaults.showTrackedChanges = config.showTrackedChanges;
  if (typeof config.extraViewDefault === 'string') viewDefaults[config.extraViewDefault] = true;
  return {
    archive: config.minimal ? only(parser, REQUIRED) : parser,
    viewDefaults,
    close() {
      if (probe) probe[0] += 1;
      parser.free();
    },
  };
}

function only(archive, methods) {
  return Object.fromEntries(methods.map((method) => [method, archive[method].bind(archive)]));
}

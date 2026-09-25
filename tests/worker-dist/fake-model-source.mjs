// Test-only model source module. It opens OOXML packages with the ordinary
// parser WASM glue, so the production workers exercise the generic
// ModelSource contract (module URL import, view defaults, host layout and
// missing optional capabilities) against a real archive. The page supplies
// all behaviour switches through the descriptor config.
import initDocx, { DocxArchive } from '/packages/docx/src/wasm/docx_parser.js';
import initXlsx, { XlsxArchive } from '/packages/xlsx/src/wasm/xlsx_parser.js';

const DOCX_REQUIRED = [
  'open_document_cursor', 'pull_document_chunk', 'document_chunk_done',
  'acknowledge_document_chunk', 'cancel_document_cursor', 'close_document_session',
  'assert_healthy', 'extract_image',
];
const DOCX_OPTIONAL = ['document_cursor_resource_usage', 'resource_usage', 'to_markdown'];
const XLSX_REQUIRED = [
  'open_sheet_cursor', 'pull_sheet_cursor', 'sheet_cursor_pull_finished',
  'sheet_cursor_resource_usage', 'acknowledge_sheet_cursor_terminal', 'cancel_sheet_cursor',
  'close_sheet_cursor', 'assert_healthy', 'parse', 'extract_image',
];
const XLSX_OPTIONAL = ['resource_usage', 'to_markdown'];

let docxReady;
let xlsxReady;

function delegate(archive, methods) {
  const wrapper = {};
  for (const method of methods) wrapper[method] = (...args) => archive[method](...args);
  return wrapper;
}

export async function openModelSource(bytes, config) {
  if (config.format === 'docx') {
    docxReady ??= initDocx({
      module_or_path: new URL('/packages/docx/src/wasm/docx_parser_bg.wasm', import.meta.url),
    });
    await docxReady;
    const archive = new DocxArchive(bytes);
    return {
      archive: delegate(archive, config.minimal ? DOCX_REQUIRED : [...DOCX_REQUIRED, ...DOCX_OPTIONAL]),
      viewDefaults: config.showTrackedChanges === true ? { showTrackedChanges: true } : {},
      close: () => archive.free(),
    };
  }
  if (config.format === 'xlsx') {
    xlsxReady ??= initXlsx({
      module_or_path: new URL('/packages/xlsx/src/wasm/xlsx_parser_bg.wasm', import.meta.url),
    });
    await xlsxReady;
    const archive = new XlsxArchive(bytes);
    const wrapper = delegate(archive, config.minimal ? XLSX_REQUIRED : [...XLSX_REQUIRED, ...XLSX_OPTIONAL]);
    if (typeof config.layoutFamily === 'string') {
      const request = config.layoutFamily === ''
        ? null
        : { family: config.layoutFamily, sizePt: config.layoutSizePt, bold: false, italic: false };
      wrapper.host_layout_request = () => new TextEncoder().encode(JSON.stringify(request));
      wrapper.configure_host_layout = () => undefined;
    }
    return { archive: wrapper, close: () => archive.free() };
  }
  throw new TypeError('unsupported fake model source format');
}

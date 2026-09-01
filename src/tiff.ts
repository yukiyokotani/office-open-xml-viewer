// Opt-in TIFF image codec entry point: `@silurus/ooxml/tiff`.

import { renderTiffToBitmap } from '../packages/core/src/image/tiff.js';
import type { TiffRenderer } from '../packages/core/src/image/tiff-contract.js';
import { registerBuiltinWorkerRenderer } from '../packages/core/src/worker/renderer-module-contract.js';

/** Optional TIFF 6.0 codec shared by DOCX, XLSX and PPTX viewers. */
export const tiff: TiffRenderer = registerBuiltinWorkerRenderer({
  render: renderTiffToBitmap,
}, 'tiff');

export type { TiffRenderer } from '../packages/core/src/image/tiff-contract.js';

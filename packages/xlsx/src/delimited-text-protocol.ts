import type { NormalizedXlsxWorksheetPolicy } from './internal/worksheet-policy.js';
import type {
  WorkerRendererDescriptors,
  WorkerErrorPayload,
} from '@silurus/ooxml-core/worker';
import type { ParsedWorkbook } from './types.js';
import type { ResolvedDelimitedTextOptions } from './delimited-text.js';

export type DelimitedTextParseRequest = {
  readonly type: 'parseDelimitedText';
  readonly id: number;
  readonly data: ArrayBuffer;
  readonly options: ResolvedDelimitedTextOptions;
  readonly worksheetPolicy?: NormalizedXlsxWorksheetPolicy;
  /** Validated `LoadOptions.fonts` descriptors (plain, structured-clone
   * safe data). The render worker re-validates them and runs the same demand
   * and session as main mode in its own FontFaceSet; the parse-only worker
   * ignores them. Never the deprecated `useGoogleFonts`. */
  readonly fonts?: readonly import('@silurus/ooxml-core/internal/font-sources').FontSourceDescriptor[];
  /** Caller-owned load-time cell-font language gate, shared with main layout. */
  readonly japaneseCellFonts?: boolean;
  readonly cjkFallback?: import('@silurus/ooxml-core').CjkLang;
  readonly renderers?: WorkerRendererDescriptors;
};

export type DelimitedTextParseResponse =
  | {
      readonly type: 'delimitedTextParsed';
      readonly id: number;
      readonly workbook: ParsedWorkbook;
      readonly worksheetJson: ArrayBuffer;
    }
  | ({ readonly type: 'error'; readonly id: number } & WorkerErrorPayload);

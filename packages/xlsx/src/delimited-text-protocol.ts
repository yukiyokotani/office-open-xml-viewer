import type { WorkerRendererDescriptors, WorkerErrorPayload } from '@silurus/ooxml-core/worker';
import type { ParsedWorkbook } from './types.js';
import type { ResolvedDelimitedTextOptions } from './delimited-text.js';

export type DelimitedTextParseRequest = {
  readonly type: 'parseDelimitedText';
  readonly id: number;
  readonly data: ArrayBuffer;
  readonly options: ResolvedDelimitedTextOptions;
  readonly useGoogleFonts?: boolean;
  readonly googleFontsCssOrigin?: import('@silurus/ooxml-core').LoadOptions['googleFontsCssOrigin'];
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

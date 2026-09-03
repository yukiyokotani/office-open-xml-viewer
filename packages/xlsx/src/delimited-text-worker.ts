import { serializeWorkerError } from '@silurus/ooxml-core/worker';
import { parseDelimitedWorksheet } from './delimited-text.js';
import type {
  DelimitedTextParseRequest,
  DelimitedTextParseResponse,
} from './delimited-text-protocol.js';

self.onmessage = (event: MessageEvent<DelimitedTextParseRequest>): void => {
  const request = event.data;
  try {
    const { workbook, worksheet } = parseDelimitedWorksheet(request.data, request.options);
    const worksheetJson = new TextEncoder().encode(JSON.stringify(worksheet)).buffer as ArrayBuffer;
    const response: DelimitedTextParseResponse = {
      type: 'delimitedTextParsed',
      id: request.id,
      workbook,
      worksheetJson,
    };
    (self.postMessage as (message: unknown, transfer: Transferable[]) => void)(response, [
      worksheetJson,
    ]);
  } catch (error) {
    const response: DelimitedTextParseResponse = {
      type: 'error',
      id: request.id,
      ...serializeWorkerError(error),
    };
    self.postMessage(response);
  }
};

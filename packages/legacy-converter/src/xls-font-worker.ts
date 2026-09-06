import { measureXlsFont, type LegacyXlsFontMeasurement, type LegacyXlsNormalFont } from './xls-font-metrics.js';

export const XLS_FONT_REQUEST = 'legacy-xls-font-request';
export const XLS_FONT_RESULT = 'legacy-xls-font-result';

interface MessagePort {
  addEventListener(type: 'message', listener: EventListener): void;
  removeEventListener(type: 'message', listener: EventListener): void;
  postMessage(message: unknown): void;
}

function isFont(value: unknown): value is LegacyXlsNormalFont {
  if (typeof value !== 'object' || value === null) return false;
  const font = value as Partial<LegacyXlsNormalFont>;
  return typeof font.family === 'string' && font.family.length > 0 && font.family.length <= 255
    && typeof font.sizePoints === 'number' && Number.isFinite(font.sizePoints)
    && font.sizePoints > 0 && font.sizePoints <= 65535 / 20
    && typeof font.bold === 'boolean' && typeof font.italic === 'boolean';
}

/** One metric exchange for one disposable worker; functions never cross realms. */
export function attachXlsFontMeasurement(worker: MessagePort, measure: LegacyXlsFontMeasurement): () => void {
  const controller = new AbortController();
  let requested = false;
  const listener: EventListener = (event) => {
    const data: unknown = (event as MessageEvent).data;
    if (typeof data !== 'object' || data === null || !('type' in data) || data.type !== XLS_FONT_REQUEST) return;
    if (requested || controller.signal.aborted) return;
    requested = true;
    const font = 'font' in data ? data.font : undefined;
    const result = isFont(font)
      ? measureXlsFont(measure, { family: font.family, sizePoints: font.sizePoints, bold: font.bold, italic: font.italic }, controller.signal)
      : Promise.reject(new Error('invalid XLS font request'));
    void result.then((width) => {
      if (!controller.signal.aborted) worker.postMessage({ type: XLS_FONT_RESULT, width });
    }, () => {
      if (!controller.signal.aborted) worker.postMessage({ type: XLS_FONT_RESULT, failed: true });
    }).catch(() => { /* Worker termination/message errors are owned by the adapter. */ });
  };
  worker.addEventListener('message', listener);
  return () => { controller.abort(); worker.removeEventListener('message', listener); };
}

/** Worker-side measurement. Messages intentionally have no conversion requestId:
 * the generic core protocol remains unchanged and cannot mistake them for output. */
export function requestXlsFontMeasurement(scope: MessagePort): LegacyXlsFontMeasurement {
  let used = false;
  return (font, signal) => new Promise<number | undefined>((resolve, reject) => {
    if (used || signal.aborted) { reject(new Error('XLS measurement unavailable')); return; }
    used = true;
    const cleanup = () => {
      scope.removeEventListener('message', listener);
      signal.removeEventListener('abort', aborted);
    };
    const aborted = () => { cleanup(); reject(new Error('XLS measurement aborted')); };
    const listener: EventListener = (event) => {
      const data: unknown = (event as MessageEvent).data;
      if (typeof data !== 'object' || data === null || !('type' in data) || data.type !== XLS_FONT_RESULT) return;
      cleanup();
      const width = 'width' in data ? data.width : undefined;
      if (('failed' in data && data.failed) || (width !== undefined && (typeof width !== 'number' || !Number.isInteger(width) || width < 1 || width > 4096))) {
        reject(new Error('XLS measurement failed'));
      } else { resolve(width as number | undefined); }
    };
    scope.addEventListener('message', listener);
    signal.addEventListener('abort', aborted, { once: true });
    try { scope.postMessage({ type: XLS_FONT_REQUEST, font }); }
    catch (error) { cleanup(); reject(error); }
  });
}

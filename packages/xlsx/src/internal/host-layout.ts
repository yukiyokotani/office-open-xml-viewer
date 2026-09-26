/**
 * Host layout for XLSX model sources.
 *
 * Worksheet column widths and cell-anchored drawings are expressed in the
 * Normal style font's maximum digit width (ECMA-376 §18.3.1.13). A model
 * source that stores drawing anchors in that unit cannot resolve them without
 * the width the grid is actually painted with, so it may ask the host: its
 * archive's optional `host_layout_request()` names the Normal font, and the
 * host answers through `configure_host_layout(mdw?)` exactly once, before
 * `parse()` and before the first sheet cursor opens.
 *
 * The host measures with the renderer's own `computeMdw` — the same code that
 * sizes the painted grid — in the realm that renders: locally in a render
 * worker, through the page for a parse worker whose renderer lives on the main
 * thread, and with the session canvas factory in Node. This module holds only
 * the worker-safe protocol; it never imports the renderer.
 */

/** The Normal-style font a model source asks the host to measure. */
export interface HostLayoutFont {
  readonly family: string;
  readonly sizePt: number;
  readonly bold: boolean;
  readonly italic: boolean;
}

/** Optional archive capabilities an XLSX model source may implement. */
export interface HostLayoutArchive {
  host_layout_request?(): Uint8Array;
  configure_host_layout?(maximumDigitWidth?: number): void;
}

/** Worker -> page request (parse worker only; carries no document identity). */
export const XLSX_HOST_LAYOUT_REQUEST = 'xlsx-host-layout-request';
/** Page -> worker reply to {@link XLSX_HOST_LAYOUT_REQUEST}. */
export const XLSX_HOST_LAYOUT_RESULT = 'xlsx-host-layout-result';

/** Implementation bound: the request is a font tuple, never document data. */
const MAX_REQUEST_BYTES = 4096;
const MAX_FAMILY_LENGTH = 255;
/** BIFF/ECMA font sizes are bounded by 16-bit twips (409.55 pt). */
const MAX_FONT_SIZE_PT = 65535 / 20;
/** Implementation bound on a measured digit width, in CSS pixels. */
const MAX_DIGIT_WIDTH_PX = 4096;

/** Decode `host_layout_request()`: `null` means no font needs measuring. */
export function decodeHostLayoutRequest(bytes: Uint8Array): HostLayoutFont | null {
  if (!(bytes instanceof Uint8Array) || bytes.byteLength > MAX_REQUEST_BYTES) {
    throw new RangeError('XLSX host layout request byte budget exceeded');
  }
  let value: unknown;
  try {
    value = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes));
  } catch {
    throw new TypeError('invalid XLSX host layout request');
  }
  if (value === null) return null;
  return validateHostLayoutFont(value);
}

export function validateHostLayoutFont(value: unknown): HostLayoutFont {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new TypeError('invalid XLSX host layout font');
  }
  const keys = Object.keys(value);
  const expected = ['family', 'sizePt', 'bold', 'italic'];
  if (keys.length !== expected.length || !expected.every((key) => Object.hasOwn(value, key))) {
    throw new TypeError('invalid XLSX host layout font');
  }
  const { family, sizePt, bold, italic } = value as Record<string, unknown>;
  if (
    typeof family !== 'string' || family.length === 0 || family.length > MAX_FAMILY_LENGTH
    || typeof sizePt !== 'number' || !Number.isFinite(sizePt)
    || sizePt <= 0 || sizePt > MAX_FONT_SIZE_PT
    || typeof bold !== 'boolean' || typeof italic !== 'boolean'
  ) {
    throw new TypeError('invalid XLSX host layout font');
  }
  return Object.freeze({ family, sizePt, bold, italic });
}

/** Admit a measured width: a positive integer pixel count, else `undefined`. */
export function admitMaximumDigitWidth(value: unknown): number | undefined {
  return typeof value === 'number'
    && Number.isInteger(value)
    && value >= 1
    && value <= MAX_DIGIT_WIDTH_PX
    ? value
    : undefined;
}

/**
 * Run the one host layout decision a model-source archive asks for. Returns
 * the width passed to `configure_host_layout`, or `undefined` when the
 * archive does not ask, names no font, or the host cannot measure it.
 */
export async function configureHostLayout(
  archive: HostLayoutArchive,
  measure: (font: HostLayoutFont) => number | undefined | Promise<number | undefined>,
): Promise<number | undefined> {
  if (typeof archive.host_layout_request !== 'function') return undefined;
  if (typeof archive.configure_host_layout !== 'function') {
    throw new TypeError('XLSX model source archive must implement configure_host_layout()');
  }
  const font = decodeHostLayoutRequest(archive.host_layout_request());
  const width = font ? admitMaximumDigitWidth(await measure(font)) : undefined;
  archive.configure_host_layout(width);
  return width;
}

interface MessageScope {
  addEventListener(type: 'message', listener: EventListener): void;
  removeEventListener(type: 'message', listener: EventListener): void;
  postMessage(message: unknown): void;
}

let nextRequestId = 1;

/** Parse-worker side: ask the page to measure `font` with its renderer. */
export function requestHostLayoutFromPage(
  scope: MessageScope,
  font: HostLayoutFont,
): Promise<number | undefined> {
  const requestId = nextRequestId++;
  return new Promise((resolve) => {
    const listener: EventListener = (event) => {
      const data: unknown = (event as MessageEvent).data;
      if (!isHostLayoutResult(data) || data.requestId !== requestId) return;
      scope.removeEventListener('message', listener);
      resolve(admitMaximumDigitWidth(data.maximumDigitWidth));
    };
    scope.addEventListener('message', listener);
    try {
      scope.postMessage({ type: XLSX_HOST_LAYOUT_REQUEST, requestId, font });
    } catch {
      scope.removeEventListener('message', listener);
      resolve(undefined);
    }
  });
}

export function isHostLayoutResult(value: unknown): value is Readonly<{
  type: typeof XLSX_HOST_LAYOUT_RESULT;
  requestId: number;
  maximumDigitWidth?: number;
}> {
  return typeof value === 'object' && value !== null
    && (value as { type?: unknown }).type === XLSX_HOST_LAYOUT_RESULT
    && typeof (value as { requestId?: unknown }).requestId === 'number';
}

/**
 * Page side: answer one {@link XLSX_HOST_LAYOUT_REQUEST}. Returns false for any
 * other message. A measurement failure answers with no width.
 */
export function respondToHostLayoutRequest(
  post: (message: unknown) => void,
  message: unknown,
  measure: (font: HostLayoutFont) => number | undefined,
): boolean {
  if (
    typeof message !== 'object' || message === null
    || (message as { type?: unknown }).type !== XLSX_HOST_LAYOUT_REQUEST
  ) return false;
  const requestId = (message as { requestId?: unknown }).requestId;
  if (typeof requestId !== 'number') return true;
  let maximumDigitWidth: number | undefined;
  try {
    maximumDigitWidth = admitMaximumDigitWidth(
      measure(validateHostLayoutFont((message as { font?: unknown }).font)),
    );
  } catch {
    maximumDigitWidth = undefined;
  }
  post({
    type: XLSX_HOST_LAYOUT_RESULT,
    requestId,
    ...(maximumDigitWidth === undefined ? {} : { maximumDigitWidth }),
  });
  return true;
}

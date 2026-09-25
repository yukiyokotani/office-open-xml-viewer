import { validateLegacySourceDescriptor, type LegacyDirectSourceDescriptor } from './legacy-source-descriptor.js';

/** First-party native decoder selection for XLS. */
export interface LegacyXlsDirectSourceDescriptor extends LegacyDirectSourceDescriptor<'xls'> {}

export const MAX_LEGACY_XLS_SOURCE_BYTES = 256 * 1024 * 1024;

export function validateLegacyXlsSourceDescriptor(value: unknown): LegacyXlsDirectSourceDescriptor {
  return validateLegacySourceDescriptor(value, 'xls');
}

/** The Normal-style font a direct XLS worker asks its host to measure. */
export interface LegacyXlsNormalFont {
  readonly family: string;
  readonly sizePoints: number;
  readonly bold: boolean;
  readonly italic: boolean;
}

/**
 * Measure digits 0–9 in this font at 96 dpi and return the rounded maximum
 * advance in pixels (integer 1–4096), or undefined when it cannot be
 * measured. The signal is aborted if the load is cancelled. Never fetch a
 * font URL supplied by the document: family is an untrusted name.
 */
export type LegacyXlsFontMeasurement = (
  font: Readonly<LegacyXlsNormalFont>,
  signal: AbortSignal,
) => number | undefined | Promise<number | undefined>;

/** The worker message port a host service listens on. */
export interface LegacyXlsHostPort {
  addEventListener(type: 'message', listener: EventListener): void;
  removeEventListener(type: 'message', listener: EventListener): void;
  postMessage(message: unknown): void;
}

/**
 * Host-side services the XLS source supplies to the spreadsheet host, which
 * only calls them: the source owns its measurement policy, including the
 * default used when the caller passes no measurement.
 */
export interface LegacyXlsHostServices {
  /** The measurement a load uses for this caller override, if any. */
  resolve(override: LegacyXlsFontMeasurement | undefined): LegacyXlsFontMeasurement | undefined;
  /** Serve one worker's measurement request; returns the detach function. */
  attach(worker: LegacyXlsHostPort, measure: LegacyXlsFontMeasurement): () => void;
}

const HOST_SERVICES = new WeakMap<object, LegacyXlsHostServices>();

/** Associate host services with a source descriptor object (main realm only;
 *  the descriptor itself stays structured-clone-safe). */
export function bindLegacyXlsHostServices<T extends LegacyXlsDirectSourceDescriptor>(
  descriptor: T,
  services: LegacyXlsHostServices,
): T {
  HOST_SERVICES.set(descriptor, services);
  return descriptor;
}

/** The host services bound to a caller-supplied source, if any. */
export function legacyXlsHostServices(source: unknown): LegacyXlsHostServices | undefined {
  return typeof source === 'object' && source !== null ? HOST_SERVICES.get(source) : undefined;
}

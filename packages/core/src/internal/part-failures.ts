import { graphemeClusterOffsets } from '../text/sea-break.js';
import { safeUtf16Prefix } from './utf16-prefix.js';

/**
 * Replacement text for the library's own "could not be displayed" placeholder
 * — the card a viewer paints in place of a document, slide or sheet whose part
 * failed to parse. Only this library-invented UI text is customizable: the
 * glyph, colours, card geometry and page/slide/sheet size never change.
 *
 * An omitted field keeps the built-in text. `detail: null` (or `''`) hides the
 * detail line. The text is painted as canvas text only — never as markup.
 */
export interface ErrorPlaceholderText {
  readonly title?: string;
  readonly detail?: string | null;
}

/**
 * Static placeholder text, or a function of the part failure that returns it.
 *
 * The function always runs on the main thread (also in `mode: 'worker'`), only
 * for a part that failed, and possibly on every repaint of that part, so it
 * must be cheap and side-effect free. Returning `undefined` (or `{}`) keeps the
 * built-in text for that part. A function that throws or returns an invalid
 * value makes that render call reject.
 */
export type ErrorPlaceholderOption<F> =
  | ErrorPlaceholderText
  | ((failure: Readonly<F>) => ErrorPlaceholderText | undefined);

/** Normalized text: `title: null` is the built-in title; `detail: undefined`
 *  is the built-in detail and `null` hides the line. @internal */
export interface ResolvedErrorPlaceholder {
  readonly title: string | null;
  readonly detail: string | null | undefined;
}

/**
 * Structured-clone-safe form of the option sent with a worker render request.
 * Absent is equivalent to `builtin`. `resolve-on-main` is sent only when a
 * function option is configured but the main thread does not yet know whether
 * the part failed. @internal
 */
export type WireErrorPlaceholder =
  | { readonly state: 'builtin' }
  | { readonly state: 'resolved'; readonly title: string | null; readonly detail: string | null | undefined }
  | { readonly state: 'resolve-on-main' };

export const BUILTIN_ERROR_PLACEHOLDER: WireErrorPlaceholder = Object.freeze({ state: 'builtin' });

const DESCRIBE = (value: unknown): string => (value === null ? 'null' : typeof value);

/**
 * Validate and copy caller placeholder text. Each property is read exactly
 * once. Returns `null` for `undefined` / `{}` (the built-in text).
 */
export function normalizeErrorPlaceholderText(
  value: unknown,
  label = 'errorPlaceholder',
): ResolvedErrorPlaceholder | null {
  if (value === undefined) return null;
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${label} must be an object with optional title/detail strings, got ${DESCRIBE(value)}`);
  }
  const record = value as { readonly title?: unknown; readonly detail?: unknown };
  const title = record.title;
  const detail = record.detail;
  if (title !== undefined && typeof title !== 'string') {
    throw new TypeError(`${label}.title must be a string, got ${DESCRIBE(title)}`);
  }
  if (detail !== undefined && detail !== null && typeof detail !== 'string') {
    throw new TypeError(`${label}.detail must be a string or null, got ${DESCRIBE(detail)}`);
  }
  if (title === undefined && detail === undefined) return null;
  return Object.freeze({ title: title ?? null, detail: detail === '' ? null : detail });
}

/** A caller option checked once at the engine/viewer boundary. @internal */
export type PreparedErrorPlaceholder<F> =
  | { readonly kind: 'none' }
  | { readonly kind: 'static'; readonly text: ResolvedErrorPlaceholder | null }
  | { readonly kind: 'function'; readonly resolve: (failure: Readonly<F>) => unknown };

const NO_ERROR_PLACEHOLDER = Object.freeze({ kind: 'none' as const });

/** Validate the public option. Static text is normalized now; a function is
 *  kept to be called per failed part. @internal */
export function prepareErrorPlaceholder<F>(
  option: unknown,
  label = 'errorPlaceholder',
): PreparedErrorPlaceholder<F> {
  if (option === undefined) return NO_ERROR_PLACEHOLDER;
  if (typeof option === 'function') {
    return Object.freeze({ kind: 'function', resolve: option as (failure: Readonly<F>) => unknown });
  }
  return Object.freeze({ kind: 'static', text: normalizeErrorPlaceholderText(option, label) });
}

/** Resolve the text for one failed part (main thread). A throwing or invalid
 *  function result propagates to the render call. @internal */
export function resolveErrorPlaceholderFor<F>(
  prepared: PreparedErrorPlaceholder<F>,
  failure: Readonly<F>,
  label = 'errorPlaceholder',
): ResolvedErrorPlaceholder | null {
  if (prepared.kind === 'none') return null;
  if (prepared.kind === 'static') return prepared.text;
  return normalizeErrorPlaceholderText(prepared.resolve(failure), `${label}() result`);
}

export function wireErrorPlaceholder(resolved: ResolvedErrorPlaceholder | null): WireErrorPlaceholder {
  return resolved === null
    ? BUILTIN_ERROR_PLACEHOLDER
    : Object.freeze({ state: 'resolved', title: resolved.title, detail: resolved.detail });
}

/** Read a wire field back into resolved text (`builtin`/absent → `null`).
 *  `resolve-on-main` must be handled by the caller first. @internal */
export function resolvedFromWire(wire: WireErrorPlaceholder | undefined): ResolvedErrorPlaceholder | null {
  if (wire === undefined || wire.state === 'builtin') return null;
  if (wire.state !== 'resolved') throw new Error('errorPlaceholder was not resolved on the main thread');
  const title = typeof wire.title === 'string' ? wire.title : null;
  const detail = wire.detail === null || typeof wire.detail === 'string' ? wire.detail : undefined;
  return Object.freeze({ title, detail: detail === '' ? null : detail });
}

/** The painted title for a failed part. */
export function errorPlaceholderTitle(resolved: ResolvedErrorPlaceholder | null, builtin: string): string {
  return resolved?.title ?? builtin;
}

/** The painted detail for a failed part; `null` paints no detail line. */
export function errorPlaceholderDetail(
  resolved: ResolvedErrorPlaceholder | null,
  builtin: string,
): string | null {
  if (resolved === null || resolved.detail === undefined) return builtin;
  return resolved.detail;
}

/** Bytes kept for one failure record's `message` (the painted default detail
 *  keeps the full parser string). */
export const PART_FAILURE_MESSAGE_MAX_UTF8 = 1024;
/** Aggregate `message` bytes kept per engine; later records get `''`. */
export const PART_FAILURE_MESSAGES_MAX_UTF8 = 1024 * 1024;

const ELLIPSIS = '…';
const ELLIPSIS_UTF8 = 3;

function utf8Length(value: string): number {
  let bytes = 0;
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if (code < 0x80) bytes += 1;
    else if (code < 0x800) bytes += 2;
    else if (code >= 0xD800 && code <= 0xDBFF && index + 1 < value.length) {
      const next = value.charCodeAt(index + 1);
      if (next >= 0xDC00 && next <= 0xDFFF) {
        bytes += 4;
        index += 1;
      } else bytes += 3;
    } else bytes += 3;
  }
  return bytes;
}

/**
 * Clamp a file-derived diagnostic to `maxUtf8` bytes at a grapheme boundary,
 * ending in `…` when shortened. Work is bounded by `maxUtf8`, not by the input.
 */
export function clampPartFailureMessage(message: string, maxUtf8 = PART_FAILURE_MESSAGE_MAX_UTF8): string {
  // Every UTF-16 code unit is at least one UTF-8 byte, so this prefix holds
  // every candidate cut; one extra unit tells whether anything was dropped.
  const window = safeUtf16Prefix(message, maxUtf8 + 1);
  if (window.length === message.length && utf8Length(message) <= maxUtf8) return message;
  const budget = maxUtf8 - ELLIPSIS_UTF8;
  const boundaries = graphemeClusterOffsets(window);
  // The window may end inside a cluster, so its last boundary is not a cut.
  let cut = 0;
  let previous = 0;
  let bytes = 0;
  for (const boundary of boundaries) {
    if (boundary >= window.length) break;
    bytes += utf8Length(window.slice(previous, boundary));
    if (bytes > budget) break;
    cut = boundary;
    previous = boundary;
  }
  return `${window.slice(0, cut)}${ELLIPSIS}`;
}

/**
 * Per-engine list of failed parts. Records are frozen, deduplicated by key
 * (the first copy wins), carry a bounded copy of the parser message, and are
 * kept until the engine is destroyed. Never subtracted on cache eviction.
 * @internal
 */
export class PartFailureLog<F extends { readonly message: string }> {
  readonly #records = new Map<string | number, F>();
  readonly #listeners = new Set<(failure: F) => void>();
  #messageBytes = 0;
  #snapshot: readonly F[] | null = null;
  #closed = false;

  constructor(
    private readonly keyOf: (failure: F) => string | number,
    private readonly compare: (left: F, right: F) => number,
  ) {}

  get size(): number {
    return this.#records.size;
  }

  has(key: string | number): boolean {
    return this.#records.has(key);
  }

  get(key: string | number): F | undefined {
    return this.#records.get(key);
  }

  /** Frozen snapshot in the format's order. */
  get list(): readonly F[] {
    if (this.#snapshot === null) {
      this.#snapshot = Object.freeze([...this.#records.values()].sort(this.compare));
    }
    return this.#snapshot;
  }

  /**
   * Record a failure. `failure.message` is the full parser string; the record
   * keeps a clamped copy within the aggregate budget. Returns the stored
   * record, or `undefined` when the key is already recorded or the log closed.
   */
  add(failure: F): F | undefined {
    if (this.#closed) return undefined;
    const key = this.keyOf(failure);
    if (this.#records.has(key)) return undefined;
    let message = clampPartFailureMessage(String(failure.message));
    const bytes = utf8Length(message);
    if (this.#messageBytes + bytes > PART_FAILURE_MESSAGES_MAX_UTF8) message = '';
    else this.#messageBytes += bytes;
    const record = Object.freeze({ ...failure, message }) as F;
    this.#records.set(key, record);
    this.#snapshot = null;
    for (const listener of [...this.#listeners]) {
      try {
        listener(record);
      } catch (error) {
        console.error('[ooxml] part failure listener failed:', error);
      }
    }
    return record;
  }

  subscribe(listener: (failure: F) => void): () => void {
    if (this.#closed) return () => {};
    this.#listeners.add(listener);
    return () => { this.#listeners.delete(listener); };
  }

  close(): void {
    this.#closed = true;
    this.#listeners.clear();
    this.#records.clear();
    this.#snapshot = Object.freeze([]);
    this.#messageBytes = 0;
  }
}

const partFailureLogs = new WeakMap<object, PartFailureLog<{ readonly message: string }>>();

/**
 * The failure log owned by an engine, created on first use. Keyed by identity
 * so the log never appears on the engine's public surface and engines built
 * off-prototype in tests work unchanged. @internal
 */
export function partFailureLogOf<F extends { readonly message: string }>(
  owner: object,
  create: () => PartFailureLog<F>,
): PartFailureLog<F> {
  let log = partFailureLogs.get(owner) as PartFailureLog<F> | undefined;
  if (!log) {
    log = create();
    partFailureLogs.set(owner, log as unknown as PartFailureLog<{ readonly message: string }>);
  }
  return log;
}

interface QueuedPartFailure<F> {
  readonly generation: number;
  readonly log: object;
  readonly key: string | number;
  readonly failure: F;
}

/**
 * Viewer-side `onPartFailure` delivery for the committed engine.
 *
 * `attach` switches to a newly committed engine: it subscribes and replays the
 * failures that engine already holds. Each delivery carries the attachment
 * generation and is re-checked before every callback, so a callback that
 * destroys the viewer or starts a reload drops the rest. Delivery is once per
 * (engine, key). Callback errors are logged, never routed to `onError`.
 * @internal
 */
export class PartFailureDelivery<F extends { readonly message: string }> {
  #generation = 0;
  #unsubscribe: (() => void) | null = null;
  #closed = false;
  #draining = false;
  readonly #queue: QueuedPartFailure<F>[] = [];
  readonly #delivered = new WeakMap<object, Set<string | number>>();

  constructor(
    private readonly viewerName: string,
    private readonly callback: ((failure: Readonly<F>) => void) | undefined,
    private readonly keyOf: (failure: F) => string | number,
  ) {}

  attach(log: PartFailureLog<F> | null): void {
    if (this.#closed) return;
    const generation = ++this.#generation;
    this.#unsubscribe?.();
    this.#unsubscribe = null;
    if (!log || !this.callback) return;
    this.#unsubscribe = log.subscribe((failure) => {
      this.#enqueue(generation, log, failure);
    });
    for (const failure of log.list) this.#enqueue(generation, log, failure);
  }

  close(): void {
    if (this.#closed) return;
    this.#closed = true;
    this.#generation++;
    this.#unsubscribe?.();
    this.#unsubscribe = null;
    this.#queue.length = 0;
  }

  #enqueue(generation: number, log: object, failure: F): void {
    if (this.#closed || generation !== this.#generation) return;
    this.#queue.push({ generation, log, key: this.keyOf(failure), failure });
    if (this.#draining) return;
    this.#draining = true;
    try {
      while (this.#queue.length > 0) {
        const entry = this.#queue.shift()!;
        if (this.#closed || entry.generation !== this.#generation) continue;
        let delivered = this.#delivered.get(entry.log);
        if (!delivered) {
          delivered = new Set();
          this.#delivered.set(entry.log, delivered);
        }
        if (delivered.has(entry.key)) continue;
        delivered.add(entry.key);
        try {
          this.callback!(entry.failure);
        } catch (error) {
          console.error(`[ooxml] ${this.viewerName} onPartFailure callback failed:`, error);
        }
      }
    } finally {
      this.#draining = false;
    }
  }
}

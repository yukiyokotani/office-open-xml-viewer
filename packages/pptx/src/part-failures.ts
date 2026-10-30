import { PartFailureLog, partFailureLogOf } from '@silurus/ooxml-core/internal/part-failures';

/**
 * A PPTX slide part that failed to parse (#1974). The deck still loads; that
 * slide shows a "Slide N could not be displayed" placeholder and every other
 * slide renders normally. A broken slide's `show` flag cannot be read, so it
 * is never treated as hidden.
 *
 * `message` is the parser's part-tagged diagnostic (for example
 * `ppt/slides/slide3.xml: …`), copied and clamped to 1 KiB. It is derived from
 * the file: treat it as untrusted text (use `textContent`, never `innerHTML`).
 */
export interface PptxPartFailure {
  readonly kind: 'slide';
  /** 0-based slide index (presentation order). */
  readonly slideIndex: number;
  /** 1-based slide number, as shown in the built-in placeholder title. */
  readonly slideNumber: number;
  readonly message: string;
}

/** @internal */
export function createPptxPartFailureLog(): PartFailureLog<PptxPartFailure> {
  return new PartFailureLog<PptxPartFailure>(
    (failure) => failure.slideIndex,
    (left, right) => left.slideIndex - right.slideIndex,
  );
}

/** Validate one failure entry received from a worker. @internal */
export function normalizeWirePptxPartFailure(value: unknown, slideCount: number): PptxPartFailure {
  const entry = value as Partial<Record<keyof PptxPartFailure, unknown>> | null;
  if (
    entry === null || typeof entry !== 'object'
    || entry.kind !== 'slide'
    || !Number.isInteger(entry.slideIndex) || (entry.slideIndex as number) < 0
    || (entry.slideIndex as number) >= slideCount
    || !Number.isInteger(entry.slideNumber)
    || typeof entry.message !== 'string'
  ) {
    throw new Error('PPTX worker returned an invalid part failure');
  }
  return {
    kind: 'slide',
    slideIndex: entry.slideIndex as number,
    slideNumber: entry.slideNumber as number,
    message: entry.message,
  };
}

/** The engine's failure log (created on first use). @internal */
export function pptxPartFailureLog(engine: object): PartFailureLog<PptxPartFailure> {
  return partFailureLogOf(engine, createPptxPartFailureLog);
}

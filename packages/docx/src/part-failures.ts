import { PartFailureLog, partFailureLogOf } from '@silurus/ooxml-core/internal/part-failures';

/**
 * A DOCX package part that failed to parse (#1974). The document still loads
 * and shows a "could not be displayed" placeholder page instead of its content.
 *
 * `message` is the parser's part-tagged diagnostic (for example
 * `word/document.xml: …`), copied and clamped to 1 KiB. It is derived from the
 * file: treat it as untrusted text (use `textContent`, never `innerHTML`).
 */
export interface DocxPartFailure {
  readonly kind: 'document';
  readonly message: string;
}

/** @internal */
export function createDocxPartFailureLog(): PartFailureLog<DocxPartFailure> {
  return new PartFailureLog<DocxPartFailure>(() => 'document', () => 0);
}

/** The engine's failure log (created on first use). @internal */
export function docxPartFailureLog(engine: object): PartFailureLog<DocxPartFailure> {
  return partFailureLogOf(engine, createDocxPartFailureLog);
}

/** The committed engine's log for viewer delivery; reading `partFailures`
 * first lets the document record a failure from its loaded state. @internal */
export function syncedDocxPartFailureLog(
  document: { readonly partFailures?: readonly DocxPartFailure[] },
): PartFailureLog<DocxPartFailure> {
  void document.partFailures;
  return docxPartFailureLog(document);
}

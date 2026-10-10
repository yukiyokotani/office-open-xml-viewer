/** A validated caller request cannot poison an already acquired reading layout.
 * Keep this exact code across the ordinary worker error wire; parser/resource
 * RangeError and TypeError failures retain the terminal acquisition policy. */
export class DocxCallerInputError extends RangeError {
  readonly code = 'docx-caller-input';
}
export function isDocxCallerInputError(error: unknown): boolean {
  return error instanceof Error && 'code' in error && error.code === 'docx-caller-input';
}

/**
 * Admission of one `fontResolver` answer, shared by the realm that paints
 * (`source-session.ts`) and by the main-thread side of the worker bridge
 * (`resolver-bridge-responder.ts`), so a resolver answer is judged by the same
 * rules whether the function runs in the painting realm or is bridged to it.
 * Library resource policy, not OOXML behaviour.
 */
import { fontSourceLimitError } from '../errors/font-source-limit.js';
import { HARD_MAX_FONT_SOURCE_DECLARATIONS } from '../worker/resource-policy.generated.js';
import { validateFontFaceDeclaration, type FontFaceDeclaration } from './source.js';

/** A resolver answer: validated declarations, or a failure reason that keeps
 * every demanded family on its host fallback. */
export type AdmittedFontResolverResult =
  | { readonly ok: true; readonly declarations: readonly FontFaceDeclaration[] }
  | { readonly ok: false; readonly reason: string };

/**
 * Validate what a resolver function returned. A non-array, a malformed
 * declaration, more than `HARD_MAX_FONT_SOURCE_DECLARATIONS` entries or a
 * face beyond the face ceiling is a failure of the call's families (`ok:
 * false`: they keep their host fallback with a warning). A run-time answer
 * never rejects the load: only factory input is rejected outright.
 */
export function admitFontResolverResult(raw: unknown): AdmittedFontResolverResult {
  if (!Array.isArray(raw)) return { ok: false, reason: 'resolver did not return an array' };
  if (raw.length > HARD_MAX_FONT_SOURCE_DECLARATIONS) {
    return { ok: false, reason: fontSourceLimitError('declarations', HARD_MAX_FONT_SOURCE_DECLARATIONS, raw.length).message };
  }
  try {
    return {
      ok: true,
      declarations: Object.freeze(raw.map((entry: unknown, index) =>
        validateFontFaceDeclaration(entry, `font resolver declaration ${index}`))),
    };
  } catch (error) {
    return { ok: false, reason: error instanceof Error ? error.message : String(error) };
  }
}

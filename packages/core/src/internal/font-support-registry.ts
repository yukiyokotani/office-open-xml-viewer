/** Intrinsic parsed-resource facts. This store owns no analyzer or catalogue;
 * immutable owners determine lifetime, and metadata-only parsers never attach
 * certificates. PPTX reads them only after successful font registration. */
export interface FontSupportFacts {
  readonly schema: 'ot-definedness-1';
  readonly glyphCount: number | undefined;
  readonly nonzeroPreserved: boolean | undefined;
  readonly missingIsolated: boolean | undefined;
  readonly noErasure?: boolean;
  readonly anyIndic3ScriptPresent?: boolean;
  readonly gsubLookupCount?: number;
  /** Unicode presence whose default glyph cannot reach any deletion, intersected
   * over all eligible cmaps. Unsafe is unknown, never known absence. */
  readonly erasureSafeRanges?: readonly (readonly [number, number])[];
  readonly gsubDisposition?: 'identity' | 'active-open-type' | 'profile-inactive-major';
  readonly profile?: 'canonical-static-v1';
  readonly reason?: 'glyph-domain' | 'unsupported' | 'malformed' | 'budget' | 'cycle';
}
export type FontTable = Readonly<{ offset: number; length: number }>;
const facts = new WeakMap<object, FontSupportFacts>();
export function retainFontSupportFacts(owner: object, value: FontSupportFacts): void { facts.set(owner, value); }
export function fontSupportFacts(owner: object | undefined): FontSupportFacts | undefined { return owner ? facts.get(owner) : undefined; }

import type {
  LayoutDiagnostic,
  LayoutDiagnosticCode,
  SourceRef,
} from './types.js';

export interface ParseDiagnosticWire {
  readonly code: string;
  readonly severity: 'warning' | 'error';
  readonly part: string;
  readonly path: readonly number[];
}

interface ParseDiagnosticContractEntry {
  readonly severity: ParseDiagnosticWire['severity'];
  readonly layoutCode: LayoutDiagnosticCode;
  readonly message: string;
}

/** Private cross-language contract. The Rust constants in parser/src/types.rs
 * are compared with these keys by diagnostics.test.ts so a new emitter code
 * cannot silently disappear in a newer parser/older renderer mismatch. */
export const PARSER_DIAGNOSTIC_CONTRACT = Object.freeze({
  UNSUPPORTED_TEXT_EFFECT: Object.freeze({
    severity: 'warning',
    layoutCode: 'UNSUPPORTED_FEATURE',
    message: 'WordprocessingML text effects are not rendered',
  }),
  INVALID_TEXT_EFFECT_VALUE: Object.freeze({
    severity: 'warning',
    layoutCode: 'INVALID_VALUE',
    message: 'An invalid WordprocessingML text-effect value was ignored',
  }),
  MISSING_DRAWING_EXTENT: Object.freeze({
    severity: 'error',
    layoutCode: 'INVALID_GEOMETRY',
    message: 'A drawing with a missing required extent was omitted',
  }),
  INVALID_DRAWING_EXTENT: Object.freeze({
    severity: 'error',
    layoutCode: 'INVALID_GEOMETRY',
    message: 'A drawing with an invalid extent was omitted',
  }),
  DEGENERATE_DRAWING_EXTENT: Object.freeze({
    severity: 'warning',
    layoutCode: 'INVALID_GEOMETRY',
    message: 'A drawing has a schema-valid zero-area extent',
  }),
  UNSUPPORTED_NESTED_QUARTER_TURN_GROUP_TRANSFORM: Object.freeze({
    severity: 'warning',
    layoutCode: 'UNSUPPORTED_FEATURE',
    message: 'A nested non-uniform group transform uses the standard DrawingML fallback',
  }),
} satisfies Readonly<Record<string, ParseDiagnosticContractEntry>>);

type KnownParseDiagnosticCode = keyof typeof PARSER_DIAGNOSTIC_CONTRACT;

const CONTRACT_MISMATCH_DIAGNOSTIC = Object.freeze({
  code: 'INVALID_VALUE' as const,
  severity: 'warning' as const,
  message: 'The parser diagnostic contract did not match this renderer build',
});

function isRecord(value: unknown): value is Readonly<Record<string, unknown>> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function validPath(value: unknown, bodyLength: number): value is readonly number[] {
  if (!Array.isArray(value) || !value.every((entry) =>
    Number.isSafeInteger(entry) && entry >= 0)) {
    return false;
  }
  const [bodyIndex] = value;
  return bodyIndex === undefined || bodyIndex < bodyLength;
}

function bodySource(path: readonly number[]): SourceRef {
  return Object.freeze({
    story: 'body',
    storyInstance: 'body',
    path: Object.freeze([...path]),
  });
}

/** Convert the parser-only wire into immutable retained-layout diagnostics.
 *
 * Unknown fields are never reflected: a fixed sentinel exposes a binary
 * contract mismatch without leaking authored text, part names, or raw invalid
 * values. Geometry convergence never sees these immutable parse facts; the
 * paginator attaches the mapped list exactly once to its final result. */
export function mapParseDiagnostics(
  value: unknown,
  bodyLength: number,
): readonly LayoutDiagnostic[] {
  if (value === undefined) return Object.freeze([]);
  if (!Array.isArray(value)) return Object.freeze([CONTRACT_MISMATCH_DIAGNOSTIC]);
  const mapped: LayoutDiagnostic[] = [];
  let mismatch = false;
  for (const candidate of value) {
    if (!isRecord(candidate)
      || typeof candidate.code !== 'string'
      || !Object.hasOwn(PARSER_DIAGNOSTIC_CONTRACT, candidate.code)
      || candidate.part !== 'word/document.xml'
      || !validPath(candidate.path, bodyLength)) {
      mismatch = true;
      continue;
    }
    const code = candidate.code as KnownParseDiagnosticCode;
    const contract = PARSER_DIAGNOSTIC_CONTRACT[code];
    if (candidate.severity !== contract.severity) {
      mismatch = true;
      continue;
    }
    mapped.push(Object.freeze({
      code: contract.layoutCode,
      severity: contract.severity,
      source: bodySource(candidate.path),
      message: contract.message,
    }));
  }
  if (mismatch) mapped.push(CONTRACT_MISMATCH_DIAGNOSTIC);
  return Object.freeze(mapped);
}

export class LayoutInvariantError extends Error {
  readonly code: LayoutDiagnosticCode;

  constructor(code: LayoutDiagnosticCode, detail: string) {
    super(`${code}: ${detail}`);
    this.name = 'LayoutInvariantError';
    this.code = code;
  }
}

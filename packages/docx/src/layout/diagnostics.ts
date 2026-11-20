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
  /** The one source part that may carry this code, compared exactly. */
  readonly part: string;
  readonly layoutCode: LayoutDiagnosticCode;
  readonly message: string;
}

const WORDPROCESSINGML_DOCUMENT_PART = 'word/document.xml';
/** Main document story of a producer that has no WordprocessingML part name. */
const PRODUCER_MAIN_DOCUMENT = 'document';

/** Private cross-language contract. The Rust constants in model/src/lib.rs
 * are compared with these keys, severities and parts by diagnostics.test.ts so
 * a new emitter code cannot silently disappear in a newer parser/older
 * renderer mismatch. Each code is accepted only from its own part, so a
 * WordprocessingML fact and another producer's fact cannot claim each other. */
export const PARSER_DIAGNOSTIC_CONTRACT = Object.freeze({
  REF_BOOKMARK_RESOURCE_LIMIT: Object.freeze({
    severity: 'warning',
    part: WORDPROCESSINGML_DOCUMENT_PART,
    layoutCode: 'UNSUPPORTED_FEATURE',
    message: 'REF bookmark acquisition exceeded its resource limits; cached field results were retained',
  }),
  UNSUPPORTED_VML_CONTENT: Object.freeze({
    severity: 'warning',
    part: WORDPROCESSINGML_DOCUMENT_PART,
    layoutCode: 'UNSUPPORTED_FEATURE',
    message: 'Some VML content could not be rendered',
  }),
  UNSUPPORTED_TEXT_EFFECT: Object.freeze({
    severity: 'warning',
    part: WORDPROCESSINGML_DOCUMENT_PART,
    layoutCode: 'UNSUPPORTED_FEATURE',
    message: 'WordprocessingML text effects are not rendered',
  }),
  INVALID_TEXT_EFFECT_VALUE: Object.freeze({
    severity: 'warning',
    part: WORDPROCESSINGML_DOCUMENT_PART,
    layoutCode: 'INVALID_VALUE',
    message: 'An invalid WordprocessingML text-effect value was ignored',
  }),
  MISSING_DRAWING_EXTENT: Object.freeze({
    severity: 'error',
    part: WORDPROCESSINGML_DOCUMENT_PART,
    layoutCode: 'INVALID_GEOMETRY',
    message: 'A drawing with a missing required extent was omitted',
  }),
  INVALID_DRAWING_EXTENT: Object.freeze({
    severity: 'error',
    part: WORDPROCESSINGML_DOCUMENT_PART,
    layoutCode: 'INVALID_GEOMETRY',
    message: 'A drawing with an invalid extent was omitted',
  }),
  DEGENERATE_DRAWING_EXTENT: Object.freeze({
    severity: 'warning',
    part: WORDPROCESSINGML_DOCUMENT_PART,
    layoutCode: 'INVALID_GEOMETRY',
    message: 'A drawing has a schema-valid zero-area extent',
  }),
  // A producer's retained framePr facts of nested-cell paragraphs; the shared
  // consumer keeps cell-owned tables in ordinary cell flow (table-owner-runs.ts
  // tableRowsElectCarriers), whose Word evidence covers WML/DOCX sources only.
  // No other positioning rule is implemented or claimed.
  NESTED_CELL_FRAME_FLOW: Object.freeze({
    severity: 'warning',
    part: PRODUCER_MAIN_DOCUMENT,
    layoutCode: 'UNSUPPORTED_FEATURE',
    message: 'Nested-cell paragraph frames are retained but laid out in ordinary cell flow; their frame positioning is not implemented for this source',
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
      || !validPath(candidate.path, bodyLength)) {
      mismatch = true;
      continue;
    }
    const code = candidate.code as KnownParseDiagnosticCode;
    const contract: ParseDiagnosticContractEntry = PARSER_DIAGNOSTIC_CONTRACT[code];
    if (candidate.part !== contract.part || candidate.severity !== contract.severity) {
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

import { describe, expect, it } from 'vitest';
import rustDiagnosticTypes from '../../parser/src/types.rs?raw';
import { createBodyLayoutInput } from '../body-layout-input.js';
import { createLayoutServices } from '../layout-runtime.js';
import { layoutDocument } from '../document-layout.js';
import type { InternalDocxDocumentModel } from '../parser-model.js';
import type { BodyElement, DocxDocumentModel, SectionProps } from '../types.js';
import {
  mapParseDiagnostics,
  PARSER_DIAGNOSTIC_CONTRACT,
  type ParseDiagnosticWire,
} from './diagnostics.js';
import { assertDocumentLayout } from './invariants.js';

const wire = (
  code: string,
  severity: ParseDiagnosticWire['severity'],
  path: readonly number[],
): ParseDiagnosticWire => ({
  code,
  severity,
  part: 'word/document.xml',
  path,
});

const section: SectionProps = {
  pageWidth: 200,
  pageHeight: 100,
  marginTop: 10,
  marginRight: 10,
  marginBottom: 10,
  marginLeft: 10,
  headerDistance: 5,
  footerDistance: 5,
  titlePage: false,
  evenAndOddHeaders: false,
  sectionStart: 'nextPage',
  columns: null,
} as SectionProps;

const paragraph = (): BodyElement => ({
  type: 'pageBreak',
} as unknown as BodyElement);

function document(
  diagnostics?: readonly ParseDiagnosticWire[],
): InternalDocxDocumentModel {
  return {
    section,
    body: [paragraph()],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    footnotes: [],
    endnotes: [],
    fontFamilyClasses: {},
    ...(diagnostics === undefined ? {} : { diagnostics }),
  } as unknown as InternalDocxDocumentModel;
}

function measureContext(): CanvasRenderingContext2D {
  return {
    font: '10px serif',
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText: (text: string) => ({
      width: [...text].length * 5,
      actualBoundingBoxAscent: 8,
      actualBoundingBoxDescent: 2,
      fontBoundingBoxAscent: 8,
      fontBoundingBoxDescent: 2,
    } as TextMetrics),
  } as unknown as CanvasRenderingContext2D;
}

describe('private parser diagnostic mapping', () => {
  it('maps allowlisted parse facts to fixed layout categories and source refs', () => {
    expect(mapParseDiagnostics([
      wire('UNSUPPORTED_TEXT_EFFECT', 'warning', [0]),
      wire('INVALID_TEXT_EFFECT_VALUE', 'warning', [0]),
      wire('MISSING_DRAWING_EXTENT', 'error', [0]),
      wire('INVALID_DRAWING_EXTENT', 'error', [0]),
      wire('DEGENERATE_DRAWING_EXTENT', 'warning', [0]),
      wire('UNSUPPORTED_NESTED_QUARTER_TURN_GROUP_TRANSFORM', 'warning', [0]),
    ], 1)).toEqual([
      {
        code: 'UNSUPPORTED_FEATURE',
        severity: 'warning',
        source: { story: 'body', storyInstance: 'body', path: [0] },
        message: 'WordprocessingML text effects are not rendered',
      },
      {
        code: 'INVALID_VALUE',
        severity: 'warning',
        source: { story: 'body', storyInstance: 'body', path: [0] },
        message: 'An invalid WordprocessingML text-effect value was ignored',
      },
      {
        code: 'INVALID_GEOMETRY',
        severity: 'error',
        source: { story: 'body', storyInstance: 'body', path: [0] },
        message: 'A drawing with a missing required extent was omitted',
      },
      {
        code: 'INVALID_GEOMETRY',
        severity: 'error',
        source: { story: 'body', storyInstance: 'body', path: [0] },
        message: 'A drawing with an invalid extent was omitted',
      },
      {
        code: 'INVALID_GEOMETRY',
        severity: 'warning',
        source: { story: 'body', storyInstance: 'body', path: [0] },
        message: 'A drawing has a schema-valid zero-area extent',
      },
      {
        code: 'UNSUPPORTED_FEATURE',
        severity: 'warning',
        source: { story: 'body', storyInstance: 'body', path: [0] },
        message: 'A nested non-uniform group transform uses the standard DrawingML fallback',
      },
    ]);
  });

  it('surfaces one fixed mismatch sentinel without reflecting malformed wire data', () => {
    const mapped = mapParseDiagnostics([
      wire('PRIVATE_SENTINEL_CODE', 'warning', [0]),
      wire('INVALID_DRAWING_EXTENT', 'warning', [0]),
      { ...wire('INVALID_DRAWING_EXTENT', 'error', [0]), part: 'private/sentinel.xml' },
      wire('INVALID_DRAWING_EXTENT', 'error', [-1]),
      wire('INVALID_DRAWING_EXTENT', 'error', [1]),
      wire('INVALID_DRAWING_EXTENT', 'error', [Number.MAX_SAFE_INTEGER + 1]),
    ], 1);

    expect(mapped).toEqual([{
      code: 'INVALID_VALUE',
      severity: 'warning',
      message: 'The parser diagnostic contract did not match this renderer build',
    }]);
    expect(JSON.stringify(mapped)).not.toContain('PRIVATE_SENTINEL');
    expect(JSON.stringify(mapped)).not.toContain('private/sentinel.xml');
    expect(mapParseDiagnostics(null, 1)).toEqual([{
      code: 'INVALID_VALUE',
      severity: 'warning',
      message: 'The parser diagnostic contract did not match this renderer build',
    }]);
  });

  it('preserves validated nested numeric source coordinates', () => {
    expect(mapParseDiagnostics([
      wire('UNSUPPORTED_TEXT_EFFECT', 'warning', [0, 3, 2]),
    ], 1)).toEqual([{
      code: 'UNSUPPORTED_FEATURE',
      severity: 'warning',
      source: { story: 'body', storyInstance: 'body', path: [0, 3, 2] },
      message: 'WordprocessingML text effects are not rendered',
    }]);
  });

  it('keeps the Rust emitter and TypeScript mapper code sets exhaustive', () => {
    const rustCodes = [...rustDiagnosticTypes.matchAll(
      /pub const PARSE_DIAGNOSTIC_CODE_[A-Z_]+:\s*&str\s*=\s*"([A-Z_]+)";/g,
    )].map((match) => match[1]!).sort();

    expect(rustCodes).toEqual(
      Object.keys(PARSER_DIAGNOSTIC_CONTRACT).sort(),
    );

    const rustSeverities = Object.fromEntries(
      [...rustDiagnosticTypes.matchAll(
        /pub const PARSE_DIAGNOSTIC_SEVERITY_([A-Z_]+):\s*DiagnosticSeverity\s*=\s*DiagnosticSeverity::(Warning|Error);/g,
      )].map((match) => [match[1]!, match[2]!.toLowerCase()]),
    );
    expect(rustSeverities).toEqual(Object.fromEntries(
      Object.entries(PARSER_DIAGNOSTIC_CONTRACT)
        .map(([code, contract]) => [code, contract.severity]),
    ));
  });

  it('crosses acquisition once and freezes the final layout diagnostic graph', () => {
    const model = document([
      wire('UNSUPPORTED_TEXT_EFFECT', 'warning', [0]),
    ]);
    const input = createBodyLayoutInput(model as DocxDocumentModel);

    expect(structuredClone(input).parserDiagnostics).toEqual(input.parserDiagnostics);
    const services = createLayoutServices(model, { measureContext: measureContext() });
    const first = layoutDocument(model, services, { currentDateMs: 0 });
    const second = layoutDocument(model, services, { currentDateMs: 0 });

    expect(first.diagnostics).toEqual([{
      code: 'UNSUPPORTED_FEATURE',
      severity: 'warning',
      source: { story: 'body', storyInstance: 'body', path: [0] },
      message: 'WordprocessingML text effects are not rendered',
    }]);
    expect(second.diagnostics).toEqual(first.diagnostics);
    expect(Object.isFrozen(first)).toBe(true);
    expect(Object.isFrozen(first.diagnostics)).toBe(true);
    expect(Object.isFrozen(first.diagnostics[0])).toBe(true);
    expect(Object.isFrozen(first.diagnostics[0]!.source)).toBe(true);
    expect(Object.isFrozen(first.diagnostics[0]!.source!.path)).toBe(true);

    expect(() => assertDocumentLayout({
      ...first,
      diagnostics: [{
        code: 'PRIVATE_SENTINEL' as never,
        severity: 'warning',
        message: 'fixed',
      }],
    })).toThrow(/diagnostics\[0\]\.code is unknown/);
    expect(() => assertDocumentLayout({
      ...first,
      diagnostics: [{
        code: 'INVALID_VALUE',
        severity: 'warning',
        source: {
          story: 'body',
          storyInstance: 'body',
          path: [-1],
        },
        message: 'fixed',
      }],
    })).toThrow(/diagnostics\[0\]\.source is invalid/);
  });
});

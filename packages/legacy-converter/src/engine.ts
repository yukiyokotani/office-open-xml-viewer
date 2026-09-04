import initWasm, {
  convert_legacy_office,
  type InitInput,
} from './wasm/legacy_office_converter.js';
import {
  HARD_MAX_LEGACY_CONVERSION_BYTES,
  LegacyOfficeConversionError,
  type LegacyOfficeConversionInput,
  type LegacyOfficeConversionResult,
  type LegacyOfficeConverter,
} from '@silurus/ooxml-core';

export const LEGACY_OFFICE_WASM_ENGINE = 'silurus-legacy-office';
export const LEGACY_OFFICE_WASM_ENGINE_VERSION = '0.1.0';

/** Internal source-injected form keeps relative WASM URLs out of inline Workers. */
export function createLegacyOfficeWasmConverterFromSource(
  resolveWasm: () => InitInput | Promise<InitInput>,
): LegacyOfficeConverter {
  let initialized: Promise<void> | undefined;
  const initialize = (): Promise<void> => {
    initialized ??= Promise.resolve(resolveWasm())
      .then((wasm) => initWasm({ module_or_path: wasm }))
      .then(() => undefined);
    return initialized;
  };
  return {
    async convert(input: Readonly<LegacyOfficeConversionInput>): Promise<LegacyOfficeConversionResult> {
      assertSameFamily(input);
      assertResourceLimits(input);
      if (input.signal.aborted) {
        throw conversionError('aborted', input);
      }
      try {
        await initialize();
      } catch {
        throw conversionError('failed', input);
      }
      if (input.signal.aborted) {
        throw conversionError('aborted', input);
      }
      let output: ReturnType<typeof convert_legacy_office> | undefined;
      try {
        output = convert_legacy_office(input.bytes, input.from, input.maxOutputBytes);
        const warnings = output.warnings().split('\n').filter(Boolean);
        const bytes = output.take_bytes();
        return {
          bytes,
          engine: LEGACY_OFFICE_WASM_ENGINE,
          engineVersion: LEGACY_OFFICE_WASM_ENGINE_VERSION,
          ...(warnings.length === 0 ? {} : { warnings }),
        };
      } catch (error) {
        const message = wasmErrorText(error);
        if (message.includes('OUTPUT_TOO_LARGE')) {
          throw conversionError('output-too-large', input);
        }
        if (message.includes('UNSUPPORTED:')) {
          throw conversionError('unsupported-input', input);
        }
        throw conversionError('failed', input);
      } finally {
        output?.free();
      }
    },
  };
}

function assertResourceLimits(input: Readonly<LegacyOfficeConversionInput>): void {
  if (input.bytes.byteLength > HARD_MAX_LEGACY_CONVERSION_BYTES) {
    throw conversionError('source-too-large', input);
  }
  if (
    !Number.isSafeInteger(input.maxOutputBytes)
    || input.maxOutputBytes <= 0
    || input.maxOutputBytes > HARD_MAX_LEGACY_CONVERSION_BYTES
  ) {
    throw conversionError('failed', input);
  }
}

function assertSameFamily(input: Readonly<LegacyOfficeConversionInput>): void {
  const valid = (input.from === 'doc' && input.to === 'docx')
    || (input.from === 'xls' && input.to === 'xlsx')
    || (input.from === 'ppt' && input.to === 'pptx');
  if (!valid) throw conversionError('unsupported-input', input);
}

function wasmErrorText(error: unknown): string {
  if (typeof error === 'string') return error;
  if (error instanceof Error) return error.message;
  return '';
}

function conversionError(
  reason: ConstructorParameters<typeof LegacyOfficeConversionError>[0],
  input: Pick<LegacyOfficeConversionInput, 'from' | 'to'>,
): LegacyOfficeConversionError {
  return new LegacyOfficeConversionError(reason, input.from, input.to);
}

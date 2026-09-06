import { describe, expect, it, vi } from 'vitest';
import { deflateRawSync } from 'node:zlib';
import { buildCfbFixture, buildStoredZip, buildZipFixture } from '../testing';
import {
  LegacyOfficeConversionError,
  normalizeOfficeInput,
  validateConvertedOoxml,
  type LegacyOfficeConversionOptions,
  type LegacyOfficeConverter,
  type LegacyOfficeFormatConversionOptions,
} from './legacy-office.js';
import type { LegacyOfficeFormat } from './legacy-office-error.js';
import {
  createDisposableWorkerLegacyOfficeConverter,
  type LegacyOfficeConversionWorker,
  type LegacyOfficeWorkerRequest,
} from './worker-converter.js';

const contentTypes: Record<'docx' | 'xlsx' | 'pptx', string> = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
};

function packageFor(format: keyof typeof contentTypes, extra = ''): Uint8Array {
  const mainParts = {
    docx: '/word/document.xml',
    xlsx: '/xl/workbook.xml',
    pptx: '/ppt/presentation.xml',
  } as const;
  return buildStoredZip({
    '[Content_Types].xml':
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
      + `<Override PartName="${mainParts[format]}" ContentType="${contentTypes[format]}"/>${extra}</Types>`,
    [mainParts[format].slice(1)]: '<root/>',
  });
}

function converterReturning(format: keyof typeof contentTypes): LegacyOfficeConverter {
  return {
    convert: vi.fn(async () => ({
      bytes: packageFor(format),
      engine: 'test-converter',
      engineVersion: '1.2.3',
      outputSha256: 'a'.repeat(64),
      warnings: ['drawing omitted'],
    })),
  };
}

function conversionFor(
  format: LegacyOfficeFormat,
  options: LegacyOfficeFormatConversionOptions,
): LegacyOfficeConversionOptions {
  return { [format]: options };
}

describe('normalizeOfficeInput', () => {
  it('leaves the existing OOXML/decryption path untouched when input is not legacy CFB', async () => {
    const bytes = packageFor('docx');
    const converter = converterReturning('docx');

    const result = await normalizeOfficeInput(bytes, 'docx', conversionFor('doc', { converter }));

    expect(result.bytes).toBe(bytes);
    expect(result.conversion).toBeUndefined();
    expect(converter.convert).not.toHaveBeenCalled();
  });

  it('retains the fail-closed legacy error when no converter is configured', async () => {
    const input = buildCfbFixture(['Root Entry', 'WordDocument']);

    await expect(normalizeOfficeInput(input, 'docx')).rejects.toMatchObject({
      code: 'legacy-binary-format',
    });
  });

  it('does not offer encrypted CFB input to a legacy converter', async () => {
    const converter = converterReturning('docx');
    const input = buildCfbFixture([
      'Root Entry',
      'EncryptionInfo',
      'EncryptedPackage',
      'WordDocument',
    ]);

    await expect(normalizeOfficeInput(
      input,
      'docx',
      conversionFor('doc', { converter }),
    )).rejects.toMatchObject({
      code: 'encrypted',
    });
    expect(converter.convert).not.toHaveBeenCalled();
  });

  it.each([
    ['WordDocument', 'doc', 'docx'],
    ['Workbook', 'xls', 'xlsx'],
    ['PowerPoint Document', 'ppt', 'pptx'],
  ] as const)('maps %s to the fixed %s -> %s conversion family', async (stream, from, to) => {
    const input = buildCfbFixture(['Root Entry', stream]);
    const converter = converterReturning(to);
    const onResult = vi.fn();

    const result = await normalizeOfficeInput(
      input,
      to,
      conversionFor(from, { converter, onResult }),
    );

    expect(converter.convert).toHaveBeenCalledOnce();
    expect(converter.convert).toHaveBeenCalledWith(expect.objectContaining({
      from,
      to,
      bytes: expect.any(Uint8Array),
      signal: expect.any(AbortSignal),
    }));
    expect(result.bytes).toEqual(packageFor(to));
    expect(result.conversion).toEqual({
      from,
      to,
      inputBytes: input.byteLength,
      outputBytes: packageFor(to).byteLength,
      engine: 'test-converter',
      engineVersion: '1.2.3',
      outputSha256: 'a'.repeat(64),
      warnings: ['drawing omitted'],
    });
    expect(onResult).toHaveBeenCalledWith(result.conversion);
  });

  it.each([
    ['doc', 'xls', 'Workbook', 'xlsx'],
    ['doc', 'ppt', 'PowerPoint Document', 'pptx'],
    ['xls', 'doc', 'WordDocument', 'docx'],
    ['xls', 'ppt', 'PowerPoint Document', 'pptx'],
    ['ppt', 'doc', 'WordDocument', 'docx'],
    ['ppt', 'xls', 'Workbook', 'xlsx'],
  ] as const)('configuring %s does not enable omitted %s input', async (
    configured,
    _inputFormat,
    stream,
    target,
  ) => {
    const configuredTarget = ({ doc: 'docx', xls: 'xlsx', ppt: 'pptx' } as const)[configured];
    const converter = converterReturning(configuredTarget);
    await expect(normalizeOfficeInput(
      buildCfbFixture(['Root Entry', stream]),
      target,
      conversionFor(configured, { converter }),
    )).rejects.toMatchObject({ code: 'legacy-binary-format' });
    expect(converter.convert).not.toHaveBeenCalled();
  });

  it('rejects an unambiguous cross-family CFB without invoking the selected converter', async () => {
    const converter = converterReturning('docx');
    await expect(normalizeOfficeInput(
      buildCfbFixture(['Root Entry', 'Workbook']),
      'docx',
      conversionFor('doc', { converter }),
    )).rejects.toMatchObject({
      reason: 'unsupported-input',
      from: 'doc',
      to: 'docx',
    });
    expect(converter.convert).not.toHaveBeenCalled();
  });

  it('isolates a diagnostics observer failure from the successful load', async () => {
    const input = buildCfbFixture(['Root Entry', 'WordDocument']);
    const result = await normalizeOfficeInput(input, 'docx', conversionFor('doc', {
      converter: converterReturning('docx'),
      onResult: () => { throw new Error('observer failed'); },
    }));
    expect(result.conversion?.to).toBe('docx');
  });

  it('records the original size after a worker converter transfers and detaches its input', async () => {
    const input = buildCfbFixture(['Root Entry', 'WordDocument']);
    const inputBytes = input.byteLength;
    const output = packageFor('docx');
    let messageListener: EventListener | undefined;
    const worker: LegacyOfficeConversionWorker = {
      postMessage: (message: LegacyOfficeWorkerRequest, transfer: Transferable[]) => {
        structuredClone(message, { transfer });
        queueMicrotask(() => messageListener?.(new MessageEvent('message', { data: {
          type: 'converted',
          requestId: 1,
          bytes: output.buffer,
        } })));
      },
      addEventListener: (type, listener) => {
        if (type === 'message') messageListener = listener;
      },
      removeEventListener: (type, listener) => {
        if (type === 'message' && messageListener === listener) messageListener = undefined;
      },
      terminate: vi.fn(),
    };

    const result = await normalizeOfficeInput(input, 'docx', conversionFor('doc', {
      converter: createDisposableWorkerLegacyOfficeConverter(() => worker),
    }));

    expect(input.byteLength).toBe(0);
    expect(result.conversion?.inputBytes).toBe(inputBytes);
    expect(result.bytes).toEqual(output);
  });

  it('rejects before conversion when the configured input budget is exceeded', async () => {
    const input = buildCfbFixture(['Root Entry', 'WordDocument']);
    const converter = converterReturning('docx');

    await expect(normalizeOfficeInput(input, 'docx', conversionFor('doc', {
      converter,
      maxInputBytes: input.byteLength - 1,
    }))).rejects.toMatchObject({
      code: 'legacy-office-conversion',
      reason: 'source-too-large',
    });
    expect(converter.convert).not.toHaveBeenCalled();
  });

  it('rejects converter output before ZIP parsing when its byte budget is exceeded', async () => {
    const input = buildCfbFixture(['Root Entry', 'WordDocument']);
    const output = packageFor('docx');

    await expect(normalizeOfficeInput(input, 'docx', conversionFor('doc', {
      converter: converterReturning('docx'),
      maxOutputBytes: output.byteLength - 1,
    }))).rejects.toMatchObject({
      code: 'legacy-office-conversion',
      reason: 'output-too-large',
    });
  });

  it('rejects timeout values that the host timer would clamp', async () => {
    const converter = converterReturning('docx');
    await expect(normalizeOfficeInput(
      buildCfbFixture(['Root Entry', 'WordDocument']),
      'docx',
      conversionFor('doc', { converter, timeoutMs: 0x80000000 }),
    )).rejects.toThrow(TypeError);
    expect(converter.convert).not.toHaveBeenCalled();
  });

  it('normalizes untrusted converter exceptions without exposing their message', async () => {
    const input = buildCfbFixture(['Root Entry', 'WordDocument']);
    const secret = 'private document name.doc';

    await expect(normalizeOfficeInput(input, 'docx', conversionFor('doc', {
      converter: { convert: async () => { throw new Error(secret); } },
    }))).rejects.toSatisfy((error: unknown) => {
      expect(error).toBeInstanceOf(LegacyOfficeConversionError);
      expect(error).toMatchObject({ reason: 'failed', from: 'doc', to: 'docx' });
      expect(String(error)).not.toContain(secret);
      return true;
    });
  });

  it('rejects malformed converter-supplied output digests', async () => {
    const input = buildCfbFixture(['Root Entry', 'WordDocument']);
    await expect(normalizeOfficeInput(input, 'docx', conversionFor('doc', {
      converter: {
        convert: async () => ({
          bytes: packageFor('docx'),
          outputSha256: 'not-a-sha256',
        }),
      },
    }))).rejects.toMatchObject({ reason: 'invalid-output' });
  });

  it('reconstructs converter-supplied typed errors so custom messages and formats cannot escape', async () => {
    const input = buildCfbFixture(['Root Entry', 'WordDocument']);
    const secret = 'customer-contract.doc';

    await expect(normalizeOfficeInput(input, 'docx', conversionFor('doc', {
      converter: {
        convert: async () => {
          throw new LegacyOfficeConversionError('unsupported-input', 'ppt', 'pptx', secret);
        },
      },
    }))).rejects.toSatisfy((error: unknown) => {
      expect(error).toMatchObject({ reason: 'unsupported-input', from: 'doc', to: 'docx' });
      expect(String(error)).not.toContain(secret);
      return true;
    });
  });

  it('rejects an already-aborted request without invoking the converter', async () => {
    const controller = new AbortController();
    controller.abort();
    const converter = converterReturning('docx');

    await expect(normalizeOfficeInput(
      buildCfbFixture(['Root Entry', 'WordDocument']),
      'docx',
      conversionFor('doc', { converter, signal: controller.signal }),
    )).rejects.toMatchObject({ reason: 'aborted' });
    expect(converter.convert).not.toHaveBeenCalled();
  });

  it('aborts a converter that exceeds its execution budget', async () => {
    const converter: LegacyOfficeConverter = {
      convert: ({ signal }) => new Promise((_resolve, reject) => {
        signal.addEventListener('abort', () => reject(new DOMException('Aborted', 'AbortError')), {
          once: true,
        });
      }),
    };

    await expect(normalizeOfficeInput(
      buildCfbFixture(['Root Entry', 'WordDocument']),
      'docx',
      conversionFor('doc', { converter, timeoutMs: 5 }),
    )).rejects.toMatchObject({ reason: 'timeout' });
  });
});

describe('validateConvertedOoxml', () => {
  it.each(['docx', 'xlsx', 'pptx'] as const)('accepts a macro-free %s package', async (format) => {
    await expect(validateConvertedOoxml(packageFor(format), format)).resolves.toBeUndefined();
  });

  it('validates the DEFLATE-compressed content-types part emitted by real ZIP writers', async () => {
    const content = new TextEncoder().encode(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
      + `<Override PartName="/word/document.xml" ContentType="${contentTypes.docx}"/>`
      + '</Types>',
    );
    const compressed = buildZipFixture({
      '[Content_Types].xml': content,
      'word/document.xml': '<document/>',
    }, (_name, bytes) => ({
      method: 8,
      bytes: new Uint8Array(deflateRawSync(bytes)),
    }));
    await expect(validateConvertedOoxml(compressed, 'docx')).resolves.toBeUndefined();
  });

  it.each([
    ['docx', 'xlsx'],
    ['docx', 'pptx'],
    ['xlsx', 'docx'],
    ['xlsx', 'pptx'],
    ['pptx', 'docx'],
    ['pptx', 'xlsx'],
  ] as const)('rejects valid macro-free %s converter output requested as %s', async (
    output,
    requested,
  ) => {
    const source = { docx: 'doc', xlsx: 'xls', pptx: 'ppt' }[requested];
    await expect(validateConvertedOoxml(packageFor(output), requested)).rejects.toMatchObject({
      reason: 'invalid-output',
      from: source,
      to: requested,
    });
  });

  it('requires the declared main part to exist in the ZIP', async () => {
    const missingMain = buildStoredZip({
      '[Content_Types].xml':
        `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
        + `<Override PartName="/word/document.xml" ContentType="${contentTypes.docx}"/></Types>`,
    });
    await expect(validateConvertedOoxml(missingMain, 'docx')).rejects.toMatchObject({
      reason: 'invalid-output',
    });
  });

  it('does not accept a target content type hidden in comments or a DTD', async () => {
    const fake = buildStoredZip({
      '[Content_Types].xml':
        `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
        + `<!-- <Override PartName="/word/document.xml" ContentType="${contentTypes.docx}"/> -->`
        + `<Override PartName="/xl/workbook.xml" ContentType="${contentTypes.xlsx}"/></Types>`,
    });
    await expect(validateConvertedOoxml(fake, 'docx')).rejects.toMatchObject({
      reason: 'invalid-output',
    });

    const dtd = buildStoredZip({
      '[Content_Types].xml':
        `<!DOCTYPE Types [<!ENTITY target "${contentTypes.docx}">]>`
        + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        + '<Override PartName="/word/document.xml" ContentType="&target;"/></Types>',
    });
    await expect(validateConvertedOoxml(dtd, 'docx')).rejects.toMatchObject({
      reason: 'invalid-output',
    });
  });

  it('rejects macro-bearing output even when its main part is macro-free', async () => {
    const macro = '<Override PartName="/word/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>';
    await expect(validateConvertedOoxml(packageFor('docx', macro), 'docx')).rejects.toMatchObject({
      reason: 'invalid-output',
    });
  });

  it('rejects auxiliary VBA data even without a vbaProject content type', async () => {
    const vbaData = '<Override PartName="/word/vbaData.xml" ContentType="application/vnd.ms-word.vbaData+xml"/>';
    await expect(validateConvertedOoxml(packageFor('docx', vbaData), 'docx')).rejects.toMatchObject({
      reason: 'invalid-output',
    });
  });

  it('rejects known active-content entry names even when content types are disguised', async () => {
    const withDisguisedVba = buildStoredZip({
      '[Content_Types].xml':
        `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
        + `<Default Extension="bin" ContentType="application/octet-stream"/>`
        + `<Override PartName="/word/document.xml" ContentType="${contentTypes.docx}"/></Types>`,
      'word/document.xml': '<document/>',
      'word/vbaProject.bin': new Uint8Array([1, 2, 3]),
    });
    await expect(validateConvertedOoxml(withDisguisedVba, 'docx')).rejects.toMatchObject({
      reason: 'invalid-output',
    });
  });

  it('decodes XML character references before checking active content types', async () => {
    const escapedMacro = '<Override PartName="/word/vbaProject.bin" ContentType="application/vnd.ms-office.vba&#80;roject"/>';
    await expect(validateConvertedOoxml(packageFor('docx', escapedMacro), 'docx')).rejects.toMatchObject({
      reason: 'invalid-output',
    });
  });

  it('rejects structurally malformed content-types XML and illegal character references', async () => {
    const unclosed = buildStoredZip({
      '[Content_Types].xml':
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
        + `<Override PartName="/word/document.xml" ContentType="${contentTypes.docx}"/>`,
      'word/document.xml': '<document/>',
    });
    await expect(validateConvertedOoxml(unclosed, 'docx')).rejects.toMatchObject({
      reason: 'invalid-output',
    });

    const illegalCharacter = packageFor(
      'docx',
      '<Default Extension="bin" ContentType="application/octet-stream&#1;"/>',
    );
    await expect(validateConvertedOoxml(illegalCharacter, 'docx')).rejects.toMatchObject({
      reason: 'invalid-output',
    });
  });

  it('rejects malformed and ambiguous ZIP packages', async () => {
    await expect(validateConvertedOoxml(new Uint8Array([0x50, 0x4b]), 'docx')).rejects.toMatchObject({
      reason: 'invalid-output',
    });
    const duplicate = buildStoredZip({
      '[Content_Types].xml': '<Types/>',
      './[Content_Types].xml': '<Types/>',
    });
    await expect(validateConvertedOoxml(duplicate, 'docx')).rejects.toMatchObject({
      reason: 'invalid-output',
    });
  });

  it('rejects OPC-equivalent duplicate part names and content-type mappings', async () => {
    const duplicatePart = buildStoredZip({
      '[Content_Types].xml':
        `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
        + `<Override PartName="/word/document.xml" ContentType="${contentTypes.docx}"/></Types>`,
      'word/document.xml': '<document/>',
      'WORD/DOCUMENT.XML': '<document/>',
    });
    await expect(validateConvertedOoxml(duplicatePart, 'docx')).rejects.toMatchObject({
      reason: 'invalid-output',
    });

    const duplicateMapping = packageFor(
      'docx',
      '<Default Extension="XML" ContentType="application/xml"/>'
        + '<Default Extension="xml" ContentType="application/xml"/>',
    );
    await expect(validateConvertedOoxml(duplicateMapping, 'docx')).rejects.toMatchObject({
      reason: 'invalid-output',
    });
  });
});

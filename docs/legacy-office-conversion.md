# Opt-in legacy Office conversion

The viewer can normalize legacy binary Office bytes before its existing OOXML
parser runs:

- `.doc` to macro-free `.docx`
- `.xls` to macro-free `.xlsx`
- `.ppt` to macro-free `.pptx`

This is an adapter API, not a bundled conversion engine. Ordinary DOCX, XLSX,
and PPTX loads do not import, fetch, initialize, or retain a converter engine or
its WASM. If no converter is supplied, legacy input continues to reject with
`OoxmlError.code === 'legacy-binary-format'`.

The renderer remains OOXML-only. A successful conversion enters exactly the
same parser, model, layout, and Canvas renderer as a native OOXML package.

## Converter contract

```typescript
import { DocxDocument, type LegacyOfficeConverter } from '@silurus/ooxml/docx';

const converter: LegacyOfficeConverter = {
  async convert({ bytes, from, to, maxOutputBytes, signal }) {
    // Run an application-owned local engine or explicitly configured service.
    // The library never supplies a remote endpoint or uploads these bytes.
    const result = await convertLegacyOffice(bytes, {
      from,
      to,
      maxOutputBytes,
      signal,
    });
    return {
      bytes: result.bytes,
      engine: 'example-engine',
      engineVersion: '1.0.0',
      outputSha256: result.outputSha256,
      warnings: result.warnings,
    };
  },
};

const document = await DocxDocument.load(input, {
  legacyConversion: {
    converter,
    timeoutMs: 120_000,
    maxInputBytes: 256 * 1024 * 1024,
    maxOutputBytes: 512 * 1024 * 1024,
    onResult(record) {
      // Content-free provenance: formats, sizes, engine, version, digest, warnings.
      conversionAuditLog.push(record);
    },
  },
});
```

The mapping is selected by the receiving API and cannot cross families:
`DocxDocument` requests `doc -> docx`, `XlsxWorkbook` requests `xls -> xlsx`,
and `PptxPresentation` requests `ppt -> pptx`. A converter must still verify the
binary structures it receives and return `unsupported-input` for an unsupported
version or feature set.

The same `legacyConversion` option is available on the browser viewers and the
Node `open*` / `materialize*` APIs. Node resolves conversion before it lazily
initializes parser WASM.

## Disposable Worker adapter

CPU-heavy browser conversion should run in a dedicated Worker. The shared
adapter transfers the source `ArrayBuffer` into one disposable Worker, transfers
the generated package back, and terminates that Worker on success, failure,
cancellation, or timeout:

```typescript
import {
  createDisposableWorkerLegacyOfficeConverter,
} from '@silurus/ooxml/legacy-conversion';

const converter = createDisposableWorkerLegacyOfficeConverter(
  () => new Worker(new URL('./legacy-office.worker.js', import.meta.url), {
    type: 'module',
  }),
  {
    maxConcurrency: 1,
    maxQueuedConversions: 4,
  },
);
```

The Worker installs the matching one-shot host around the application-owned
converter or WASM wrapper:

```typescript
import {
  installLegacyOfficeConversionWorkerHandler,
  type LegacyOfficeConverter,
} from '@silurus/ooxml/legacy-conversion';

const wasmConverter: LegacyOfficeConverter = {
  async convert(input) {
    await initializeConverterWasm();
    return convertWithWasm(input);
  },
};

installLegacyOfficeConversionWorkerHandler(self, wasmConverter);
```

Converter WASM and parser WASM own separate linear memories. The generated
OOXML package must therefore materialize as a standalone buffer once. Transfer
lists prevent additional JavaScript-realm clones, but they cannot eliminate the
copy from converter memory into that buffer or the parser's later copy into its
own memory.

The converter owns its request bytes and may detach their backing buffer. After
resolution, ownership of the returned bytes belongs to the host. Neither side
may retain or mutate bytes after ownership has moved.

## Validation and failure behavior

Converter output is rejected before parser handoff unless it has a bounded,
consistent ZIP central directory, the requested main document part, and a
readable `[Content_Types].xml` that declares that part as DOCX, XLSX, or PPTX.
The preflight also rejects ZIP encryption, unsupported ZIP compression,
duplicate entries and content-type declarations, macro-capable content types,
known VBA and ActiveX part names, malformed content-types XML, and output beyond
the configured limit. These checks are defense in depth; the converter remains
responsible for removing macros and embedded executable content, and the viewer
never executes those features. The ordinary OOXML path is intentionally
unchanged by this converter-only preflight.

Conversion failures are `LegacyOfficeConversionError` instances with stable
`code === 'legacy-office-conversion'`, `stage === 'conversion'`, formats, and one
of these reasons:

- `aborted`
- `timeout`
- `source-too-large`
- `output-too-large`
- `unsupported-input`
- `failed`
- `invalid-output`

Free-form converter exception messages are not propagated. Converter identity,
version, optional lowercase SHA-256, and warnings are bounded metadata supplied
by the converter and must never include document text, filenames, source URLs,
or passwords. The host checks the digest syntax but deliberately does not make a
second full pass over potentially large output. Verify it independently when the
converter is outside the ingestion system's trust boundary.

The defaults admit a 256 MiB source, a 512 MiB output, and two minutes of
conversion. Both byte limits have a non-configurable 1 GiB hard ceiling. A
custom in-process converter must honor the supplied `AbortSignal`; the disposable
Worker adapter enforces cancellation by terminating the Worker even if its WASM
code cannot cooperatively yield. The converter also receives `maxOutputBytes`
so it can stop before materializing an inadmissible package. Viewer reload,
supersession, and destruction
are combined with an application-supplied conversion signal, so an in-flight
disposable converter Worker is also terminated when its owning view no longer
needs the result. One disposable-adapter instance defaults to one live Worker
and four queued conversions; share that instance wherever an application needs
one common concurrency boundary. A full queue rejects with
`capacity-exceeded`.

Encrypted legacy binaries are outside this initial contract. They continue to
fail through the existing encryption path. Macros, external-link updates, and
embedded code must never be executed by a converter.

## Current implementation boundary

This repository currently provides the opt-in contract, browser/Node
normalization, converter-output preflight, typed errors, and disposable Worker
transport. A purpose-built legacy-format WASM engine is still under investigation
in [issue #1472](https://github.com/yukiyokotani/office-open-xml-viewer/issues/1472).
Applications can supply their own local adapter now, but should treat converted
OOXML as a derived representation and preserve the original binary as the
authoritative source.

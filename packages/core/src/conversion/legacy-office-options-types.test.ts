import { expectTypeOf, it } from 'vitest';
import type {
  LegacyOfficeConversionOptions,
  LegacyOfficeConverter,
} from './legacy-office.js';

const converter: LegacyOfficeConverter = {
  async convert() { return { bytes: new Uint8Array() }; },
};

const sources = {
  doc: { protocol: 'ooxml-legacy-doc-source/v1', builtin: 'doc', wasmUrl: 'https://example.test/doc.wasm' },
  ppt: { protocol: 'ooxml-legacy-ppt-source/v1', builtin: 'ppt', wasmUrl: 'https://example.test/ppt.wasm' },
  xls: { protocol: 'ooxml-legacy-xls-source/v1', builtin: 'xls', wasmUrl: 'https://example.test/xls.wasm' },
} as const;

it('accepts converter-only and each same-format source-only option', () => {
  expectTypeOf({ doc: { converter } }).toMatchTypeOf<LegacyOfficeConversionOptions>();
  expectTypeOf({ xls: { converter } }).toMatchTypeOf<LegacyOfficeConversionOptions>();
  expectTypeOf({ ppt: { converter } }).toMatchTypeOf<LegacyOfficeConversionOptions>();
  expectTypeOf({ doc: { source: sources.doc } }).toMatchTypeOf<LegacyOfficeConversionOptions>();
  expectTypeOf({ xls: { source: sources.xls } }).toMatchTypeOf<LegacyOfficeConversionOptions>();
  expectTypeOf({ ppt: { source: sources.ppt } }).toMatchTypeOf<LegacyOfficeConversionOptions>();
});

it('rejects ambiguous source plus converter options', () => {
  // @ts-expect-error A native source and byte converter are mutually exclusive.
  const doc: LegacyOfficeConversionOptions = { doc: { source: sources.doc, converter } };
  // @ts-expect-error A native source and byte converter are mutually exclusive.
  const xls: LegacyOfficeConversionOptions = { xls: { source: sources.xls, converter } };
  // @ts-expect-error A native source and byte converter are mutually exclusive.
  const ppt: LegacyOfficeConversionOptions = { ppt: { source: sources.ppt, converter } };
  void [doc, xls, ppt];
});

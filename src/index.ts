export * as pptx from './pptx';
export * as xlsx from './xlsx';
export * as docx from './docx';
// Font sources for the shared `fonts` load option. The format entries export
// the same objects; they are flat here so `import { googleFonts } from
// '@silurus/ooxml'` works as the `useGoogleFonts` deprecation message says.
export {
  fontCssApi,
  fontCss,
  fontFiles,
  fontResolver,
  googleFonts,
  googleFontsSubstitutions,
  officeFontSubstitutions,
  type FontCssApiOptions,
  type FontDemand,
  type FontCssOptions,
  type FontFaceDeclaration,
  type FontFilesOptions,
  type FontResolveContext,
  type FontResolveFn,
  type FontResolverOptions,
  type FontSource,
  type FontSourceCredentials,
  type FontSubstituteScript,
  type FontSubstitution,
  type GoogleFontsOptions,
} from './docx';

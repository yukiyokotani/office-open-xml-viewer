#!/usr/bin/env node

import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';

const SOURCE = 'packages/ooxml-common/resource-policy.json';
const OUTPUTS = Object.freeze({
  typescript: 'packages/core/src/worker/resource-policy.generated.ts',
  rust: 'packages/ooxml-common/src/resource-policy.generated.rs',
});

function parseArgs(argv) {
  const options = { root: process.cwd(), write: false };
  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index];
    if (argument === '--root') options.root = path.resolve(argv[++index]);
    else if (argument === '--write') options.write = true;
    else throw new Error(`Unknown argument: ${argument}`);
  }
  return options;
}

function requireExactObject(value, location, expectedKeys) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`${location} must contain one object.`);
  }
  const actualKeys = Object.keys(value).sort();
  if (actualKeys.join('\n') !== [...expectedKeys].sort().join('\n')) {
    throw new Error(`${location} must contain exactly: ${expectedKeys.join(', ')}.`);
  }
}

function requirePositiveSafeInteger(value, location) {
  if (!Number.isSafeInteger(value) || value <= 0) {
    throw new Error(`${location} must be a positive safe integer.`);
  }
}

function readPolicy(root) {
  const sourcePath = path.join(root, SOURCE);
  const value = JSON.parse(readFileSync(sourcePath, 'utf8'));
  const defaultKeys = [
    'maxArchiveEntryBytes',
    'maxTotalInflatedBytes',
    'maxArchiveEntries',
  ];
  const hardKeys = [
    ...defaultKeys,
    'maxCentralDirectoryBytes',
    'maxDocxBodyBlockXmlBytes',
    'maxDocxBodyChunkJsonBytes',
    'maxDocxBootstrapJsonBytes',
    'maxDocxRetainedModelJsonBytes',
    'maxPptxSlideXmlBytes',
    'maxPptxSlideJsonBytes',
    'maxPptxSharedDependencyXmlBytes',
    'maxXmlDomComplexity',
    'maxPptxSharedDependencyProjectionBytes',
    'maxPptxSharedCacheEntries',
    'maxPptxSharedCacheProjectionBytes',
    'maxPptxBootstrapSlides',
    'maxPptxBootstrapProjectionBytes',
    'maxPptxBootstrapJsonBytes',
    'maxPptxPreflightProjectionBytes',
    'maxPptxMaterializedSlideJsonBytes',
    'maxPptxCachedSlides',
    'maxPptxCachedSlideProjectionBytes',
    'maxRawPartCacheEntries',
    'maxRawPartCacheBytes',
    'maxEmbeddedFontBytes',
    'maxPptxMarkdownBytes',
    'maxWorksheetRows',
    'maxWorksheetCells',
    'maxWorksheetCellContentUtf8Bytes',
    'maxWorksheetJsonBytes',
    'maxWorkbookCachedRows',
    'maxWorkbookCachedCells',
    'maxWorkbookCachedCellContentUtf8Bytes',
    'maxWorkbookCachedJsonBytes',
    'maxRendererCoordinateIndexEntries',
  ];
  requireExactObject(value, SOURCE, ['defaults', 'hardCeilings']);
  requireExactObject(value.defaults, `${SOURCE}.defaults`, defaultKeys);
  requireExactObject(value.hardCeilings, `${SOURCE}.hardCeilings`, hardKeys);
  for (const key of defaultKeys) {
    requirePositiveSafeInteger(value.defaults[key], `${SOURCE}.defaults.${key}`);
  }
  for (const key of hardKeys) {
    requirePositiveSafeInteger(value.hardCeilings[key], `${SOURCE}.hardCeilings.${key}`);
  }
  if (value.defaults.maxTotalInflatedBytes < value.defaults.maxArchiveEntryBytes) {
    throw new Error(`${SOURCE}.defaults.maxTotalInflatedBytes must not be smaller than maxArchiveEntryBytes.`);
  }
  if (value.hardCeilings.maxTotalInflatedBytes < value.hardCeilings.maxArchiveEntryBytes) {
    throw new Error(`${SOURCE}.hardCeilings.maxTotalInflatedBytes must not be smaller than maxArchiveEntryBytes.`);
  }
  for (const key of defaultKeys) {
    if (value.defaults[key] > value.hardCeilings[key]) {
      throw new Error(`${SOURCE}.defaults.${key} must not exceed its hard ceiling.`);
    }
  }
  if (
    value.hardCeilings.maxDocxRetainedModelJsonBytes
      < Math.max(
        value.hardCeilings.maxDocxBodyChunkJsonBytes,
        value.hardCeilings.maxDocxBootstrapJsonBytes,
      )
  ) {
    throw new Error(
      `${SOURCE}.hardCeilings.maxDocxRetainedModelJsonBytes must not be smaller than a DOCX document unit.`,
    );
  }
  if (
    value.hardCeilings.maxPptxSharedCacheProjectionBytes
      < value.hardCeilings.maxPptxSharedDependencyProjectionBytes
  ) {
    throw new Error(
      `${SOURCE}.hardCeilings.maxPptxSharedCacheProjectionBytes must not be smaller than maxPptxSharedDependencyProjectionBytes.`,
    );
  }
  if (
    value.hardCeilings.maxPptxMaterializedSlideJsonBytes
      < value.hardCeilings.maxPptxSlideJsonBytes
  ) {
    throw new Error(
      `${SOURCE}.hardCeilings.maxPptxMaterializedSlideJsonBytes must not be smaller than maxPptxSlideJsonBytes.`,
    );
  }
  if (
    value.hardCeilings.maxPptxCachedSlideProjectionBytes
      < value.hardCeilings.maxPptxSlideJsonBytes
  ) {
    throw new Error(
      `${SOURCE}.hardCeilings.maxPptxCachedSlideProjectionBytes must not be smaller than maxPptxSlideJsonBytes.`,
    );
  }
  for (const [cacheKey, worksheetKey] of [
    ['maxWorkbookCachedRows', 'maxWorksheetRows'],
    ['maxWorkbookCachedCells', 'maxWorksheetCells'],
    ['maxWorkbookCachedCellContentUtf8Bytes', 'maxWorksheetCellContentUtf8Bytes'],
    ['maxWorkbookCachedJsonBytes', 'maxWorksheetJsonBytes'],
  ]) {
    if (value.hardCeilings[cacheKey] < value.hardCeilings[worksheetKey]) {
      throw new Error(
        `${SOURCE}.hardCeilings.${cacheKey} must not be smaller than ${worksheetKey}.`,
      );
    }
  }
  return value;
}

function render(policy) {
  const header = '// Generated by scripts/generate-ooxml-resource-policy.mjs. Do not edit.\n\n';
  const defaults = policy.defaults;
  const hard = policy.hardCeilings;
  const shared = `${header}export const STANDARD_MAX_ARCHIVE_ENTRY_BYTES = ${defaults.maxArchiveEntryBytes};\nexport const STANDARD_MAX_TOTAL_INFLATED_BYTES = ${defaults.maxTotalInflatedBytes};\nexport const STANDARD_MAX_ARCHIVE_ENTRIES = ${defaults.maxArchiveEntries};\nexport const HARD_MAX_ARCHIVE_ENTRIES = ${hard.maxArchiveEntries};\n`;
  const outputs = {
    typescript: `${shared}export const HARD_MAX_XML_DOM_COMPLEXITY = ${hard.maxXmlDomComplexity};\nexport const HARD_MAX_DOCX_BODY_BLOCK_XML_BYTES = ${hard.maxDocxBodyBlockXmlBytes};\nexport const HARD_MAX_DOCX_BODY_CHUNK_JSON_BYTES = ${hard.maxDocxBodyChunkJsonBytes};\nexport const HARD_MAX_DOCX_BOOTSTRAP_JSON_BYTES = ${hard.maxDocxBootstrapJsonBytes};\nexport const HARD_MAX_DOCX_RETAINED_MODEL_JSON_BYTES = ${hard.maxDocxRetainedModelJsonBytes};\nexport const HARD_MAX_PPTX_SLIDE_XML_BYTES = ${hard.maxPptxSlideXmlBytes};\nexport const HARD_MAX_PPTX_SLIDE_JSON_BYTES = ${hard.maxPptxSlideJsonBytes};\nexport const HARD_MAX_PPTX_SHARED_DEPENDENCY_XML_BYTES = ${hard.maxPptxSharedDependencyXmlBytes};\nexport const HARD_MAX_PPTX_SHARED_DEPENDENCY_PROJECTION_BYTES = ${hard.maxPptxSharedDependencyProjectionBytes};\nexport const HARD_MAX_PPTX_SHARED_CACHE_ENTRIES = ${hard.maxPptxSharedCacheEntries};\nexport const HARD_MAX_PPTX_SHARED_CACHE_PROJECTION_BYTES = ${hard.maxPptxSharedCacheProjectionBytes};\nexport const HARD_MAX_PPTX_BOOTSTRAP_SLIDES = ${hard.maxPptxBootstrapSlides};\nexport const HARD_MAX_PPTX_BOOTSTRAP_PROJECTION_BYTES = ${hard.maxPptxBootstrapProjectionBytes};\nexport const HARD_MAX_PPTX_BOOTSTRAP_JSON_BYTES = ${hard.maxPptxBootstrapJsonBytes};\nexport const HARD_MAX_PPTX_PREFLIGHT_PROJECTION_BYTES = ${hard.maxPptxPreflightProjectionBytes};\nexport const HARD_MAX_PPTX_MATERIALIZED_SLIDE_JSON_BYTES = ${hard.maxPptxMaterializedSlideJsonBytes};\nexport const HARD_MAX_PPTX_CACHED_SLIDES = ${hard.maxPptxCachedSlides};\nexport const HARD_MAX_PPTX_CACHED_SLIDE_PROJECTION_BYTES = ${hard.maxPptxCachedSlideProjectionBytes};\nexport const HARD_MAX_RAW_PART_CACHE_ENTRIES = ${hard.maxRawPartCacheEntries};\nexport const HARD_MAX_RAW_PART_CACHE_BYTES = ${hard.maxRawPartCacheBytes};\nexport const HARD_MAX_PPTX_MARKDOWN_BYTES = ${hard.maxPptxMarkdownBytes};\nexport const HARD_MAX_XLSX_WORKSHEET_ROWS = ${hard.maxWorksheetRows};\nexport const HARD_MAX_XLSX_WORKSHEET_CELLS = ${hard.maxWorksheetCells};\nexport const HARD_MAX_XLSX_WORKSHEET_CELL_CONTENT_UTF8_BYTES = ${hard.maxWorksheetCellContentUtf8Bytes};\nexport const HARD_MAX_XLSX_WORKSHEET_JSON_BYTES = ${hard.maxWorksheetJsonBytes};\nexport const HARD_MAX_XLSX_WORKBOOK_CACHED_ROWS = ${hard.maxWorkbookCachedRows};\nexport const HARD_MAX_XLSX_WORKBOOK_CACHED_CELLS = ${hard.maxWorkbookCachedCells};\nexport const HARD_MAX_XLSX_WORKBOOK_CACHED_CELL_CONTENT_UTF8_BYTES = ${hard.maxWorkbookCachedCellContentUtf8Bytes};\nexport const HARD_MAX_XLSX_WORKBOOK_CACHED_JSON_BYTES = ${hard.maxWorkbookCachedJsonBytes};\nexport const HARD_MAX_XLSX_RENDERER_COORDINATE_INDEX_ENTRIES = ${hard.maxRendererCoordinateIndexEntries};\n`,
    rust: `${header}pub const STANDARD_MAX_ARCHIVE_ENTRY_BYTES: u64 = ${defaults.maxArchiveEntryBytes};\npub const STANDARD_MAX_TOTAL_INFLATED_BYTES: u64 = ${defaults.maxTotalInflatedBytes};\npub const STANDARD_MAX_ARCHIVE_ENTRIES: u64 = ${defaults.maxArchiveEntries};\npub const HARD_MAX_ARCHIVE_ENTRY_BYTES: u64 = ${hard.maxArchiveEntryBytes};\npub const HARD_MAX_TOTAL_INFLATED_BYTES: u64 = ${hard.maxTotalInflatedBytes};\npub const HARD_MAX_ARCHIVE_ENTRIES: u64 = ${hard.maxArchiveEntries};\npub const HARD_MAX_CENTRAL_DIRECTORY_BYTES: u64 = ${hard.maxCentralDirectoryBytes};\npub const HARD_MAX_XML_DOM_COMPLEXITY: u64 = ${hard.maxXmlDomComplexity};\npub const HARD_MAX_DOCX_BODY_BLOCK_XML_BYTES: u64 = ${hard.maxDocxBodyBlockXmlBytes};\npub const HARD_MAX_DOCX_BODY_CHUNK_JSON_BYTES: u64 = ${hard.maxDocxBodyChunkJsonBytes};\npub const HARD_MAX_DOCX_BOOTSTRAP_JSON_BYTES: u64 = ${hard.maxDocxBootstrapJsonBytes};\npub const HARD_MAX_DOCX_RETAINED_MODEL_JSON_BYTES: u64 = ${hard.maxDocxRetainedModelJsonBytes};\npub const HARD_MAX_PPTX_SLIDE_XML_BYTES: u64 = ${hard.maxPptxSlideXmlBytes};\npub const HARD_MAX_PPTX_SLIDE_JSON_BYTES: u64 = ${hard.maxPptxSlideJsonBytes};\npub const HARD_MAX_PPTX_SHARED_DEPENDENCY_XML_BYTES: u64 = ${hard.maxPptxSharedDependencyXmlBytes};\npub const HARD_MAX_PPTX_SHARED_DEPENDENCY_PROJECTION_BYTES: u64 = ${hard.maxPptxSharedDependencyProjectionBytes};\npub const HARD_MAX_PPTX_SHARED_CACHE_ENTRIES: u64 = ${hard.maxPptxSharedCacheEntries};\npub const HARD_MAX_PPTX_SHARED_CACHE_PROJECTION_BYTES: u64 = ${hard.maxPptxSharedCacheProjectionBytes};\npub const HARD_MAX_PPTX_BOOTSTRAP_SLIDES: u64 = ${hard.maxPptxBootstrapSlides};\npub const HARD_MAX_PPTX_BOOTSTRAP_PROJECTION_BYTES: u64 = ${hard.maxPptxBootstrapProjectionBytes};\npub const HARD_MAX_PPTX_BOOTSTRAP_JSON_BYTES: u64 = ${hard.maxPptxBootstrapJsonBytes};\npub const HARD_MAX_PPTX_PREFLIGHT_PROJECTION_BYTES: u64 = ${hard.maxPptxPreflightProjectionBytes};\npub const HARD_MAX_PPTX_MATERIALIZED_SLIDE_JSON_BYTES: u64 = ${hard.maxPptxMaterializedSlideJsonBytes};\npub const HARD_MAX_PPTX_CACHED_SLIDES: u64 = ${hard.maxPptxCachedSlides};\npub const HARD_MAX_PPTX_CACHED_SLIDE_PROJECTION_BYTES: u64 = ${hard.maxPptxCachedSlideProjectionBytes};\npub const HARD_MAX_RAW_PART_CACHE_ENTRIES: u64 = ${hard.maxRawPartCacheEntries};\npub const HARD_MAX_RAW_PART_CACHE_BYTES: u64 = ${hard.maxRawPartCacheBytes};\npub const HARD_MAX_PPTX_MARKDOWN_BYTES: u64 = ${hard.maxPptxMarkdownBytes};\npub const HARD_MAX_XLSX_WORKSHEET_ROWS: u64 = ${hard.maxWorksheetRows};\npub const HARD_MAX_XLSX_WORKSHEET_CELLS: u64 = ${hard.maxWorksheetCells};\npub const HARD_MAX_XLSX_WORKSHEET_CELL_CONTENT_UTF8_BYTES: u64 = ${hard.maxWorksheetCellContentUtf8Bytes};\npub const HARD_MAX_XLSX_WORKSHEET_JSON_BYTES: u64 = ${hard.maxWorksheetJsonBytes};\npub const HARD_MAX_XLSX_WORKBOOK_CACHED_ROWS: u64 = ${hard.maxWorkbookCachedRows};\npub const HARD_MAX_XLSX_WORKBOOK_CACHED_CELLS: u64 = ${hard.maxWorkbookCachedCells};\npub const HARD_MAX_XLSX_RENDERER_COORDINATE_INDEX_ENTRIES: u64 = ${hard.maxRendererCoordinateIndexEntries};\n`,
  };
  return {
    ...outputs,
    typescript: `${outputs.typescript}export const HARD_MAX_EMBEDDED_FONT_BYTES = ${hard.maxEmbeddedFontBytes};\n`,
    rust: `${outputs.rust}pub const HARD_MAX_XLSX_WORKBOOK_CACHED_CELL_CONTENT_UTF8_BYTES: u64 = ${hard.maxWorkbookCachedCellContentUtf8Bytes};\npub const HARD_MAX_XLSX_WORKBOOK_CACHED_JSON_BYTES: u64 = ${hard.maxWorkbookCachedJsonBytes};\npub const HARD_MAX_EMBEDDED_FONT_BYTES: u64 = ${hard.maxEmbeddedFontBytes};\n`,
  };
}

function normalizeEol(value) {
  return value.replace(/\r\n?/gu, '\n');
}

export function synchronizeResourcePolicy(options) {
  const rendered = render(readPolicy(options.root));
  const stale = [];
  for (const [language, relativePath] of Object.entries(OUTPUTS)) {
    const outputPath = path.join(options.root, relativePath);
    if (options.write) {
      writeFileSync(outputPath, rendered[language]);
    } else if (
      !existsSync(outputPath)
      || normalizeEol(readFileSync(outputPath, 'utf8')) !== rendered[language]
    ) {
      stale.push(relativePath);
    }
  }
  if (stale.length > 0) {
    throw new Error(
      `Generated OOXML resource policy files are stale:\n${stale.map((file) => `- ${file}`).join('\n')}\nRun pnpm generate:resource-policy.`,
    );
  }
  process.stdout.write(
    options.write
      ? 'Generated OOXML resource policy files.\n'
      : 'OOXML resource policy files are synchronized.\n',
  );
}

const isMain = process.argv[1]
  && path.resolve(process.argv[1]) === path.resolve(fileURLToPath(import.meta.url));
if (isMain) {
  try {
    synchronizeResourcePolicy(parseArgs(process.argv.slice(2)));
  } catch (error) {
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
    process.exitCode = 1;
  }
}

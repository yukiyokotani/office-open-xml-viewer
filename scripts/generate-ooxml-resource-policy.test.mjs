import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { synchronizeResourcePolicy } from './generate-ooxml-resource-policy.mjs';
import { readArchiveEntryCount } from './measure-ooxml-archive-entries.mjs';

function fixture(policy = {
  defaults: {
    maxArchiveEntryBytes: 128,
    maxTotalInflatedBytes: 256,
    maxArchiveEntries: 12,
  },
  hardCeilings: {
    maxArchiveEntryBytes: 512,
    maxTotalInflatedBytes: 1024,
    maxArchiveEntries: 20,
    maxCentralDirectoryBytes: 64,
    maxDocxBodyBlockXmlBytes: 320,
    maxDocxBodyChunkJsonBytes: 640,
    maxDocxBootstrapJsonBytes: 640,
    maxDocxRetainedModelJsonBytes: 1280,
    maxPptxSlideXmlBytes: 320,
    maxPptxSlideJsonBytes: 640,
    maxPptxSharedDependencyXmlBytes: 160,
    maxXmlDomComplexity: 20,
    maxPptxSharedDependencyProjectionBytes: 320,
    maxPptxSharedCacheEntries: 25,
    maxPptxSharedCacheProjectionBytes: 960,
    maxPptxBootstrapSlides: 10,
    maxPptxBootstrapProjectionBytes: 320,
    maxPptxBootstrapJsonBytes: 320,
    maxPptxPreflightProjectionBytes: 640,
    maxPptxMaterializedSlideJsonBytes: 1280,
    maxPptxCachedSlides: 4,
    maxPptxCachedSlideProjectionBytes: 1280,
    maxRawPartCacheEntries: 8,
    maxRawPartCacheBytes: 1280,
    maxEmbeddedFontBytes: 320,
    maxPptxMarkdownBytes: 640,
    maxWorksheetRows: 100,
    maxWorksheetCells: 250,
    maxWorksheetCellContentUtf8Bytes: 320,
    maxWorksheetJsonBytes: 640,
    maxWorkbookCachedRows: 200,
    maxWorkbookCachedCells: 500,
    maxWorkbookCachedCellContentUtf8Bytes: 640,
    maxWorkbookCachedJsonBytes: 1280,
    maxRendererCoordinateIndexEntries: 250,
  },
}) {
  const root = mkdtempSync(path.join(tmpdir(), 'ooxml-resource-policy-'));
  mkdirSync(path.join(root, 'packages/ooxml-common/src'), { recursive: true });
  mkdirSync(path.join(root, 'packages/core/src/worker'), { recursive: true });
  writeFileSync(
    path.join(root, 'packages/ooxml-common/resource-policy.json'),
    `${JSON.stringify(policy, null, 2)}\n`,
  );
  return root;
}

test('reads the exact classic ZIP central-directory entry count for calibration', () => {
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(12, 8);
  eocd.writeUInt16LE(12, 10);
  expectArchiveEntryCount(eocd, 12);
});

function expectArchiveEntryCount(bytes, expected) {
  assert.equal(readArchiveEntryCount(bytes), expected);
}

test('generates matching TypeScript and Rust constants from one policy source', (context) => {
  const root = fixture();
  context.after(() => rmSync(root, { recursive: true, force: true }));

  synchronizeResourcePolicy({ root, write: true });
  synchronizeResourcePolicy({ root, write: false });

  assert.match(
    readFileSync(path.join(root, 'packages/core/src/worker/resource-policy.generated.ts'), 'utf8'),
    /STANDARD_MAX_ARCHIVE_ENTRY_BYTES = 128/,
  );
  assert.match(
    readFileSync(path.join(root, 'packages/ooxml-common/src/resource-policy.generated.rs'), 'utf8'),
    /STANDARD_MAX_TOTAL_INFLATED_BYTES: u64 = 256/,
  );
  assert.match(
    readFileSync(path.join(root, 'packages/core/src/worker/resource-policy.generated.ts'), 'utf8'),
    /STANDARD_MAX_ARCHIVE_ENTRIES = 12/,
  );
  assert.match(
    readFileSync(path.join(root, 'packages/ooxml-common/src/resource-policy.generated.rs'), 'utf8'),
    /STANDARD_MAX_ARCHIVE_ENTRIES: u64 = 12/,
  );
  assert.match(
    readFileSync(path.join(root, 'packages/core/src/worker/resource-policy.generated.ts'), 'utf8'),
    /HARD_MAX_DOCX_BODY_BLOCK_XML_BYTES = 320/,
  );
  assert.match(
    readFileSync(path.join(root, 'packages/ooxml-common/src/resource-policy.generated.rs'), 'utf8'),
    /HARD_MAX_DOCX_BODY_CHUNK_JSON_BYTES: u64 = 640/,
  );
  assert.match(
    readFileSync(path.join(root, 'packages/core/src/worker/resource-policy.generated.ts'), 'utf8'),
    /HARD_MAX_XLSX_WORKSHEET_JSON_BYTES = 640/,
  );
  assert.match(
    readFileSync(path.join(root, 'packages/core/src/worker/resource-policy.generated.ts'), 'utf8'),
    /HARD_MAX_XLSX_WORKBOOK_CACHED_CELL_CONTENT_UTF8_BYTES = 640/,
  );
  assert.match(
    readFileSync(path.join(root, 'packages/core/src/worker/resource-policy.generated.ts'), 'utf8'),
    /HARD_MAX_XLSX_WORKBOOK_CACHED_JSON_BYTES = 1280/,
  );
  assert.match(
    readFileSync(path.join(root, 'packages/core/src/worker/resource-policy.generated.ts'), 'utf8'),
    /HARD_MAX_PPTX_SLIDE_XML_BYTES = 320/,
  );
  assert.match(
    readFileSync(path.join(root, 'packages/ooxml-common/src/resource-policy.generated.rs'), 'utf8'),
    /HARD_MAX_PPTX_SLIDE_JSON_BYTES: u64 = 640/,
  );
  assert.match(
    readFileSync(path.join(root, 'packages/ooxml-common/src/resource-policy.generated.rs'), 'utf8'),
    /HARD_MAX_PPTX_SHARED_CACHE_PROJECTION_BYTES: u64 = 960/,
  );
  assert.match(
    readFileSync(path.join(root, 'packages/core/src/worker/resource-policy.generated.ts'), 'utf8'),
    /HARD_MAX_PPTX_MATERIALIZED_SLIDE_JSON_BYTES = 1280/,
  );
  assert.match(
    readFileSync(path.join(root, 'packages/core/src/worker/resource-policy.generated.ts'), 'utf8'),
    /HARD_MAX_PPTX_PREFLIGHT_PROJECTION_BYTES = 640/,
  );
  assert.match(
    readFileSync(path.join(root, 'packages/core/src/worker/resource-policy.generated.ts'), 'utf8'),
    /HARD_MAX_PPTX_CACHED_SLIDES = 4/,
  );
  assert.match(
    readFileSync(path.join(root, 'packages/ooxml-common/src/resource-policy.generated.rs'), 'utf8'),
    /HARD_MAX_PPTX_CACHED_SLIDE_PROJECTION_BYTES: u64 = 1280/,
  );
  assert.match(
    readFileSync(path.join(root, 'packages/core/src/worker/resource-policy.generated.ts'), 'utf8'),
    /HARD_MAX_RAW_PART_CACHE_ENTRIES = 8/,
  );
  assert.match(
    readFileSync(path.join(root, 'packages/ooxml-common/src/resource-policy.generated.rs'), 'utf8'),
    /HARD_MAX_PPTX_MARKDOWN_BYTES: u64 = 640/,
  );
  assert.match(
    readFileSync(path.join(root, 'packages/core/src/worker/resource-policy.generated.ts'), 'utf8'),
    /HARD_MAX_EMBEDDED_FONT_BYTES = 320/,
  );
  assert.match(
    readFileSync(path.join(root, 'packages/ooxml-common/src/resource-policy.generated.rs'), 'utf8'),
    /HARD_MAX_XLSX_RENDERER_COORDINATE_INDEX_ENTRIES: u64 = 250/,
  );
});

test('fails closed when a generated language file drifts', (context) => {
  const root = fixture();
  context.after(() => rmSync(root, { recursive: true, force: true }));
  synchronizeResourcePolicy({ root, write: true });
  writeFileSync(
    path.join(root, 'packages/core/src/worker/resource-policy.generated.ts'),
    '// stale\n',
  );

  assert.throws(
    () => synchronizeResourcePolicy({ root, write: false }),
    /Generated OOXML resource policy files are stale/,
  );
});

test('rejects invalid or internally inconsistent policy values', (context) => {
  const root = fixture({
    defaults: {
      maxArchiveEntryBytes: 513,
      maxTotalInflatedBytes: 1024,
      maxArchiveEntries: 12,
    },
    hardCeilings: {
      maxArchiveEntryBytes: 512,
      maxTotalInflatedBytes: 1024,
      maxArchiveEntries: 20,
      maxCentralDirectoryBytes: 64,
      maxDocxBodyBlockXmlBytes: 320,
      maxDocxBodyChunkJsonBytes: 640,
      maxDocxBootstrapJsonBytes: 640,
      maxDocxRetainedModelJsonBytes: 1280,
      maxPptxSlideXmlBytes: 320,
      maxPptxSlideJsonBytes: 640,
      maxPptxSharedDependencyXmlBytes: 160,
      maxXmlDomComplexity: 20,
      maxPptxSharedDependencyProjectionBytes: 320,
      maxPptxSharedCacheEntries: 25,
      maxPptxSharedCacheProjectionBytes: 960,
      maxPptxBootstrapSlides: 10,
      maxPptxBootstrapProjectionBytes: 320,
      maxPptxBootstrapJsonBytes: 320,
      maxPptxPreflightProjectionBytes: 640,
      maxPptxMaterializedSlideJsonBytes: 1280,
      maxPptxCachedSlides: 4,
      maxPptxCachedSlideProjectionBytes: 1280,
      maxRawPartCacheEntries: 8,
      maxRawPartCacheBytes: 1280,
      maxEmbeddedFontBytes: 320,
      maxPptxMarkdownBytes: 640,
      maxWorksheetRows: 100,
      maxWorksheetCells: 250,
      maxWorksheetCellContentUtf8Bytes: 320,
      maxWorksheetJsonBytes: 640,
      maxWorkbookCachedRows: 200,
      maxWorkbookCachedCells: 500,
      maxWorkbookCachedCellContentUtf8Bytes: 640,
      maxWorkbookCachedJsonBytes: 1280,
      maxRendererCoordinateIndexEntries: 250,
    },
  });
  context.after(() => rmSync(root, { recursive: true, force: true }));

  assert.throws(
    () => synchronizeResourcePolicy({ root, write: true }),
    /must not exceed its hard ceiling/,
  );
});

for (const [name, smallerKey, largerKey] of [
  [
    'shared cache projection smaller than one dependency',
    'maxPptxSharedCacheProjectionBytes',
    'maxPptxSharedDependencyProjectionBytes',
  ],
  [
    'materialized slide projection smaller than one slide',
    'maxPptxMaterializedSlideJsonBytes',
    'maxPptxSlideJsonBytes',
  ],
  [
    'cached slide projection smaller than one slide',
    'maxPptxCachedSlideProjectionBytes',
    'maxPptxSlideJsonBytes',
  ],
  [
    'workbook cached rows smaller than one worksheet',
    'maxWorkbookCachedRows',
    'maxWorksheetRows',
  ],
  [
    'workbook cached cells smaller than one worksheet',
    'maxWorkbookCachedCells',
    'maxWorksheetCells',
  ],
  [
    'workbook cached cell content smaller than one worksheet',
    'maxWorkbookCachedCellContentUtf8Bytes',
    'maxWorksheetCellContentUtf8Bytes',
  ],
  [
    'workbook cached JSON smaller than one worksheet',
    'maxWorkbookCachedJsonBytes',
    'maxWorksheetJsonBytes',
  ],
]) {
  test(`rejects ${name}`, (context) => {
    const root = fixture();
    context.after(() => rmSync(root, { recursive: true, force: true }));
    const policyPath = path.join(root, 'packages/ooxml-common/resource-policy.json');
    const policy = JSON.parse(readFileSync(policyPath, 'utf8'));
    policy.hardCeilings[smallerKey] = policy.hardCeilings[largerKey] - 1;
    writeFileSync(policyPath, `${JSON.stringify(policy, null, 2)}\n`);

    assert.throws(
      () => synchronizeResourcePolicy({ root, write: true }),
      new RegExp(`${smallerKey} must not be smaller than ${largerKey}`),
    );
  });
}

test('rejects a retained DOCX projection smaller than one document unit', (context) => {
  const root = fixture();
  context.after(() => rmSync(root, { recursive: true, force: true }));
  const policyPath = path.join(root, 'packages/ooxml-common/resource-policy.json');
  const policy = JSON.parse(readFileSync(policyPath, 'utf8'));
  policy.hardCeilings.maxDocxRetainedModelJsonBytes =
    policy.hardCeilings.maxDocxBodyChunkJsonBytes - 1;
  writeFileSync(policyPath, `${JSON.stringify(policy, null, 2)}\n`);

  assert.throws(
    () => synchronizeResourcePolicy({ root, write: true }),
    /maxDocxRetainedModelJsonBytes must not be smaller than a DOCX document unit/,
  );
});

test('accepts CRLF-normalized generated files on Windows-style checkouts', (context) => {
  const root = fixture();
  context.after(() => rmSync(root, { recursive: true, force: true }));
  synchronizeResourcePolicy({ root, write: true });
  for (const relativePath of [
    'packages/core/src/worker/resource-policy.generated.ts',
    'packages/ooxml-common/src/resource-policy.generated.rs',
  ]) {
    const file = path.join(root, relativePath);
    writeFileSync(file, readFileSync(file, 'utf8').replace(/\n/gu, '\r\n'));
  }

  assert.doesNotThrow(() => synchronizeResourcePolicy({ root, write: false }));
});

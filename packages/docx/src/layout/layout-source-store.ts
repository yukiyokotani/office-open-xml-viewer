import type {
  BodyElement,
  DocNote,
  DocRun,
  EmbeddedFontRef,
  SectionProps,
  TblpPr,
} from '../types.js';
import type { BodyAcquisitionInputProjections } from './acquisition-input-projections.js';
import type { BodyLayoutInput } from './body-layout-input.js';
import type { DocumentLayoutSettings } from '../layout-context.js';
import type { ImageMetadataRecord, MathOccurrence } from './resources.js';
import type { PaintResourceDescriptor, PaintResourceRegistry, SourceRef } from './types.js';
import type { ParagraphAcquisitionInput } from './text.js';
import type { DeepReadonly, ParsedUnsupportedTextBoxBlock } from './types.js';
import { imageResourceKey, sourceKey } from './source-key.js';
import { indexSealedPaintResourceDescriptors } from './paint-resources.js';
import { sealPlainData } from './plain-data.js';
import {
  projectTableColumnLayoutInput,
  type TableLayoutSource,
  type TableSourceAcquisitionInput,
} from './table-source-acquisition.js';

export interface LayoutSourceAcquisition {
  readonly acquisitionInputs: BodyAcquisitionInputProjections;
  effectiveTablePositioning(table: TableLayoutSource): DeepReadonly<TblpPr> | null;
  publicAnchorBridge(
    source: SourceRef,
    runIndex: number,
  ): Readonly<{ occurrenceId: string; pageOwned: boolean }> | null;
}

export interface LayoutParagraphAcquisitionFact {
  readonly source: SourceRef;
  readonly publicAnchorBridges: readonly (Readonly<{ occurrenceId: string; pageOwned: boolean }> | null)[];
  readonly numberingMarkerFallbackFontSizePt: number | null;
}

export type LayoutParagraphBlock = ParagraphAcquisitionInput & Readonly<{ type: 'paragraph' }>;
export type LayoutTableBlock = TableLayoutSource & Readonly<{ type: 'table' }>;
export type LayoutFlowBlock = LayoutParagraphBlock | LayoutTableBlock;
export type LayoutCellBlock =
  | LayoutFlowBlock
  | DeepReadonly<Exclude<BodyElement, { type: 'paragraph' | 'table' }>>;
export type LayoutStoryBlock =
  | LayoutCellBlock
  | ParsedUnsupportedTextBoxBlock;
export type LayoutStoryNote = DeepReadonly<Omit<DocNote, 'content'>> & Readonly<{
  content: readonly LayoutCellBlock[];
}>;

export interface LayoutTableAcquisitionFact {
  readonly source: SourceRef;
  readonly input: TableSourceAcquisitionInput;
}

export interface LayoutSourceAcquisitionFacts {
  readonly paragraphs: readonly LayoutParagraphAcquisitionFact[];
  readonly tables: readonly LayoutTableAcquisitionFact[];
}

export interface LayoutSourceFontFacts {
  readonly familyClasses: Readonly<Record<string, string>>;
  readonly familyPitches: Readonly<Record<string, string>>;
  readonly majorFamily: string | null;
  readonly minorFamily: string | null;
  readonly embeddedFonts: readonly EmbeddedFontRef[];
  readonly renderedFamilies: readonly string[];
  readonly preloadNames: readonly (string | null | undefined)[];
  readonly defaultBodyFontSizePt: number;
}

export interface LayoutBlockRepository {
  readonly body: readonly LayoutStoryBlock[];
  readonly footnotes: readonly LayoutStoryNote[];
  readonly endnotes: readonly LayoutStoryNote[];
  readonly sources: readonly SourceRef[];
  resolve(source: SourceRef): LayoutFlowBlock;
  storyRoot(source: SourceRef): readonly LayoutStoryBlock[];
}

export interface FatalParseFact {
  readonly message: string;
  readonly pageSize: Readonly<{ widthPt: number; heightPt: number }>;
}

export interface LayoutSourceDocumentFacts {
  readonly kinsoku: Readonly<{
    enabled: boolean;
    lineStartForbidden: readonly number[];
    lineEndForbidden: readonly number[];
  }>;
  readonly defaultTabPt: number;
  readonly characterSpacingControl?: string;
  readonly mathDefJc?: string;
  readonly documentHasEastAsianText: boolean;
  readonly normalStyleFontSizePt: number;
  readonly compat: DocumentLayoutSettings['compat'];
}

/** Complete parser-independent input accepted by both model and stream adapters. */
export interface LayoutSourceStoreInput {
  readonly bodyLayoutInput: BodyLayoutInput;
  readonly blockRepository: LayoutBlockRepositoryInput;
  readonly acquisitionFacts: LayoutSourceAcquisitionFacts;
  readonly section: Readonly<SectionProps>;
  readonly documentLayoutFacts: LayoutSourceDocumentFacts;
  readonly fonts: LayoutSourceFontFacts;
  readonly fontFamilyCharsets: Readonly<Record<string, string>>;
  readonly mathOccurrences: readonly MathOccurrence[];
  readonly imageMetadata: readonly ImageMetadataRecord[];
  readonly paintDescriptors: readonly PaintResourceDescriptor[];
  readonly hasPaginationFields: boolean;
  readonly requiresDomVerticalGlyphLayout: boolean;
  readonly fatalParse: FatalParseFact | null;
}

/**
 * One sealed document source for acquisition, layout, and paint. It deliberately
 * contains no public `DocxDocumentModel`; compatibility models live outside this
 * graph. Every collection and projection supplied here belongs to one acquisition
 * snapshot, so later public-model mutation cannot alter layout or paint facts.
 */
export interface LayoutSourceStore extends Readonly<Omit<LayoutSourceStoreInput, 'blockRepository' | 'acquisitionFacts' | 'paintDescriptors' | 'documentLayoutFacts'>> {
  readonly blocks: LayoutBlockRepository;
  readonly acquisition: LayoutSourceAcquisition;
  readonly paintResources: PaintResourceRegistry;
  readonly documentLayoutSettings: DocumentLayoutSettings;
}

const layoutSourceStores = new WeakSet<object>();

export function isLayoutSourceStore(value: unknown): value is LayoutSourceStore {
  return value !== null && typeof value === 'object' && layoutSourceStores.has(value);
}

type LayoutSourceStoreBaseInput = Omit<LayoutSourceStoreInput, 'bodyLayoutInput'>;

function assertExactOwnKeys(
  value: object,
  expected: readonly string[],
  label: string,
): void {
  const actual = Object.keys(value).sort();
  const canonical = [...expected].sort();
  if (actual.length !== canonical.length
    || actual.some((key, index) => key !== canonical[index])) {
    throw new TypeError(`${label} has unexpected fields: ${actual.join(',')}`);
  }
}

function assertExactMembership(
  actual: ReadonlySet<string>,
  expected: ReadonlySet<string>,
  label: string,
): void {
  const missing = [...expected].filter((key) => !actual.has(key));
  const extra = [...actual].filter((key) => !expected.has(key));
  if (missing.length !== 0 || extra.length !== 0) {
    throw new TypeError(`${label} membership mismatch; missing=${missing.join(',')} extra=${extra.join(',')}`);
  }
}

function validateBodyLayoutSources(
  input: BodyLayoutInput,
  blocks: LayoutBlockRepository,
): void {
  const requireStory = (source: SourceRef | null): void => {
    if (!source) return;
    blocks.storyRoot(source);
  };
  const requireBodyOccurrence = (source: SourceRef): void => {
    if (source.story !== 'body' || source.storyInstance !== 'body'
      || source.path.length !== 1 || blocks.body[source.path[0]!] === undefined) {
      throw new TypeError(`Unknown body layout occurrence source: ${sourceKey(source)}`);
    }
  };
  const validateSection = (section: BodyLayoutInput['initialSection']): void => {
    for (const source of [
      section.headers.default, section.headers.first, section.headers.even,
      section.footers.default, section.footers.first, section.footers.even,
    ]) requireStory(source);
  };
  if (input.source.story !== 'body' || input.source.storyInstance !== 'body'
    || input.source.path.length !== 0) {
    throw new TypeError('Body layout input requires the canonical body root');
  }
  validateSection(input.initialSection);
  for (const entry of input.sequence) {
    requireBodyOccurrence(entry.kind === 'body-block' ? entry.block.source : entry.source);
    if (entry.kind === 'body-block') {
      const block = blocks.resolve(entry.block.source);
      if (block.type !== entry.block.kind) throw new TypeError('Body layout block source kind mismatch');
    } else if (entry.kind === 'adjacent-table-group') {
      for (const table of entry.tables) {
        if (blocks.resolve(table.source).type !== 'table') {
          throw new TypeError('Adjacent table source kind mismatch');
        }
      }
    } else if (entry.kind === 'begin-section') {
      validateSection(entry.section);
    }
  }
}

function validateResourceManifests(
  blocks: LayoutBlockRepository,
  facts: LayoutSourceAcquisitionFacts,
  mathOccurrences: readonly MathOccurrence[],
  imageMetadata: readonly ImageMetadataRecord[],
  paintDescriptors: readonly PaintResourceDescriptor[],
): void {
  const imageKeys = new Set<string>();
  const imageKindByKey = new Map<string, 'image' | 'picture-bullet'>();
  const chartKeys = new Set<string>();
  const mathKeys = new Set<string>();
  const mathSourceByKey = new Map<string, string>();
  const addImage = (key: string, kind: 'image' | 'picture-bullet' = 'image'): void => {
    if (imageKeys.has(key)) throw new TypeError(`Duplicate canonical image resource: ${key}`);
    imageKeys.add(key);
    imageKindByKey.set(key, kind);
  };
  for (const fact of facts.paragraphs) {
    const paragraph = blocks.resolve(fact.source);
    if (paragraph.type !== 'paragraph') continue;
    if (fact.publicAnchorBridges.length !== paragraph.runs.length) {
      throw new TypeError(`Paragraph anchor bridge cardinality mismatch: ${sourceKey(fact.source)}`);
    }
    if (paragraph.numbering?.picBulletImagePath) {
      addImage(imageResourceKey(fact.source, paragraph.numbering.picBulletImagePath), 'picture-bullet');
    }
    paragraph.runs.forEach((run, runIndex) => {
      const runSource: SourceRef = { ...fact.source, path: [...fact.source.path, runIndex] };
      if (run.type === 'image') addImage(imageResourceKey(runSource, run.imagePath));
      if (run.type === 'chart') {
        if (chartKeys.has(run.resourceKey)) throw new TypeError(`Duplicate canonical chart resource: ${run.resourceKey}`);
        chartKeys.add(run.resourceKey);
      }
      if (run.type === 'math') {
        if (mathKeys.has(run.resourceKey)) throw new TypeError(`Duplicate canonical math resource: ${run.resourceKey}`);
        mathKeys.add(run.resourceKey);
        mathSourceByKey.set(run.resourceKey, sourceKey(run.source));
      }
      if (run.type === 'shape' && run.textBoxInput?.kind === 'compatibility') {
        for (const textBoxParagraph of run.textBoxInput.paragraphs) {
          if (textBoxParagraph.image) {
            addImage(imageResourceKey({
              ...textBoxParagraph.source,
              path: [...textBoxParagraph.source.path, 0],
            }, textBoxParagraph.image.imagePath));
          }
        }
      }
      if (run.type === 'shape' && run.fill?.fillType === 'image') {
        addImage(imageResourceKey(runSource, run.fill.imagePath));
      }
    });
  }
  const metadataKeys = new Set(imageMetadata.map((record) => record.resourceKey));
  if (metadataKeys.size !== imageMetadata.length) throw new TypeError('Duplicate image metadata resource');
  assertExactMembership(metadataKeys, imageKeys, 'Image metadata');

  const occurrenceKeys = new Set<string>();
  for (const occurrence of mathOccurrences) {
    if (occurrenceKeys.has(occurrence.resourceKey)) throw new TypeError('Duplicate math occurrence resource');
    occurrenceKeys.add(occurrence.resourceKey);
    if (mathSourceByKey.get(occurrence.resourceKey) !== sourceKey(occurrence.source)) {
      throw new TypeError(`Math occurrence source mismatch: ${occurrence.resourceKey}`);
    }
  }
  assertExactMembership(occurrenceKeys, mathKeys, 'Math occurrence');

  const paintKeys = new Set<string>();
  const paintImageKeys = new Set<string>();
  const paintChartKeys = new Set<string>();
  const paintMathKeys = new Set<string>();
  for (const descriptor of paintDescriptors) {
    if (paintKeys.has(descriptor.resourceKey)) throw new TypeError('Duplicate paint resource descriptor');
    paintKeys.add(descriptor.resourceKey);
    if (descriptor.kind === 'image' || descriptor.kind === 'picture-bullet') {
      paintImageKeys.add(descriptor.resourceKey);
      if (imageKindByKey.get(descriptor.resourceKey) !== descriptor.kind) {
        throw new TypeError(`Image paint resource kind mismatch: ${descriptor.resourceKey}`);
      }
    } else if (descriptor.kind === 'chart') {
      paintChartKeys.add(descriptor.resourceKey);
    } else {
      paintMathKeys.add(descriptor.resourceKey);
    }
  }
  assertExactMembership(paintKeys, new Set([...imageKeys, ...chartKeys, ...mathKeys]), 'Paint resource');
  assertExactMembership(paintImageKeys, imageKeys, 'Image paint resource');
  assertExactMembership(paintChartKeys, chartKeys, 'Chart paint resource');
  assertExactMembership(paintMathKeys, mathKeys, 'Math paint resource');
}

function sealLayoutSourceStoreWithBody(
  input: LayoutSourceStoreBaseInput,
  bodyLayoutInput: BodyLayoutInput,
): LayoutSourceStore {
  const blocks = createLayoutBlockRepository(input.blockRepository);
  validateBodyLayoutSources(bodyLayoutInput, blocks);
  sealPlainData(input.acquisitionFacts, 'layout source acquisition facts');
  sealPlainData(input.section, 'layout source section');
  sealPlainData(input.documentLayoutFacts, 'layout source document facts');
  sealPlainData(input.fonts, 'layout source font facts');
  sealPlainData(input.fontFamilyCharsets, 'layout source font charsets');
  sealPlainData(input.mathOccurrences, 'layout source math facts');
  sealPlainData(input.imageMetadata, 'layout source image facts');
  sealPlainData(input.paintDescriptors, 'layout source paint descriptors');
  if (input.fatalParse) sealPlainData(input.fatalParse, 'layout source fatal parse fact');
  for (const [label, value] of [
    ['block repository', blocks],
    ['body blocks', blocks.body],
    ['footnotes', blocks.footnotes],
    ['endnotes', blocks.endnotes],
    ['acquisition facts', input.acquisitionFacts],
    ['paragraph facts', input.acquisitionFacts.paragraphs],
    ['table facts', input.acquisitionFacts.tables],
    ['section', input.section],
    ['document facts', input.documentLayoutFacts],
    ['font facts', input.fonts],
    ['math facts', input.mathOccurrences],
    ['image facts', input.imageMetadata],
    ['paint descriptors', input.paintDescriptors],
  ] as const) {
    if (!Object.isFrozen(value)) throw new TypeError(`Layout source ${label} must be sealed`);
  }
  const paragraphBySource = new Map<string, LayoutParagraphAcquisitionFact>();
  for (const fact of input.acquisitionFacts.paragraphs) {
    assertExactOwnKeys(fact, [
      'source', 'publicAnchorBridges', 'numberingMarkerFallbackFontSizePt',
    ], 'Paragraph acquisition fact');
    const key = sourceKey(fact.source);
    if (paragraphBySource.has(key)) throw new TypeError(`Duplicate paragraph acquisition source: ${key}`);
    paragraphBySource.set(key, fact);
  }
  const tableSourceKeys = new Set<string>();
  const tableByIdentity = new WeakMap<TableLayoutSource, LayoutTableAcquisitionFact>();
  for (const fact of input.acquisitionFacts.tables) {
    assertExactOwnKeys(fact, ['source', 'input'], 'Table acquisition fact');
    const key = sourceKey(fact.source);
    if (tableSourceKeys.has(key)) throw new TypeError(`Duplicate table acquisition source: ${key}`);
    tableSourceKeys.add(key);
    const table = blocks.resolve(fact.source);
    if (table.type !== 'table') throw new TypeError('Table acquisition fact must identify a table');
    tableByIdentity.set(table, fact);
  }
  for (const source of blocks.sources) {
    const block = blocks.resolve(source);
    const key = sourceKey(source);
    if (block.type === 'paragraph' && !paragraphBySource.has(key)) {
      throw new TypeError(`Missing paragraph acquisition source: ${key}`);
    }
    if (block.type === 'table' && !tableSourceKeys.has(key)) {
      throw new TypeError(`Missing table acquisition source: ${key}`);
    }
  }
  validateResourceManifests(
    blocks,
    input.acquisitionFacts,
    input.mathOccurrences,
    input.imageMetadata,
    input.paintDescriptors,
  );
  const numberingMarkers = new WeakMap<object, Map<number, NonNullable<ParagraphAcquisitionInput['numberingMarkerShapeInput']>>>();
  for (const fact of input.acquisitionFacts.paragraphs) {
    const paragraph = blocks.resolve(fact.source);
    if (paragraph.type !== 'paragraph') throw new TypeError('Paragraph acquisition fact must identify a paragraph');
    for (const run of paragraph.runs) {
      if (run.type !== 'shape' || run.textBoxInput?.kind !== 'complete') continue;
      let root: readonly LayoutStoryBlock[];
      try {
        root = blocks.storyRoot(run.textBoxInput.source);
      } catch (error) {
        throw new TypeError(`Missing complete text-box story source: ${sourceKey(run.textBoxInput.source)}`, { cause: error });
      }
      if (root.length !== run.textBoxInput.blockCount) {
        throw new TypeError(`Complete text-box block count mismatch: ${sourceKey(run.textBoxInput.source)}`);
      }
    }
    if (paragraph.numbering && paragraph.numberingMarkerShapeInput
      && fact.numberingMarkerFallbackFontSizePt !== null) {
      let byFallback = numberingMarkers.get(paragraph.numbering);
      if (!byFallback) {
        byFallback = new Map();
        numberingMarkers.set(paragraph.numbering, byFallback);
      }
      byFallback.set(fact.numberingMarkerFallbackFontSizePt, paragraph.numberingMarkerShapeInput);
    }
  }
  const projections = Object.freeze({
    numberingMarkerShapeInput(numbering, _fallbackFontSizePt) {
      const retained = numberingMarkers.get(numbering)?.get(_fallbackFontSizePt);
      if (retained) return retained;
      throw new Error('Unknown numbering marker acquisition input');
    },
    paragraphMarkShapeInput(paragraph) {
      return paragraph.paragraphMarkShapeInput;
    },
    tableFormatInput(table) {
      const fact = tableByIdentity.get(table);
      if (!fact) throw new Error('Unknown table acquisition input');
      return fact.input.format;
    },
    tableColumnLayoutInput(table, availableWidthPt, intrinsicWidths, maximumWidthPt) {
      const fact = tableByIdentity.get(table);
      if (!fact) throw new Error('Unknown table acquisition input');
      return projectTableColumnLayoutInput(
        fact.input,
        availableWidthPt,
        (rowIndex, cellIndex) => intrinsicWidths(table.rows[rowIndex]!.cells[cellIndex]!),
        maximumWidthPt,
      );
    },
    tableParticipatesInOrdinaryFlow(table) {
      const fact = tableByIdentity.get(table);
      if (!fact) throw new Error('Unknown table acquisition input');
      return fact.input.format.ordinaryFlow;
    },
    paragraphAcquisitionInput(_paragraph, source) {
      const fact = paragraphBySource.get(sourceKey(source));
      if (!fact) throw new Error(`Unknown paragraph acquisition source: ${sourceKey(source)}`);
      const paragraph = blocks.resolve(source);
      if (paragraph.type !== 'paragraph') throw new Error(`Paragraph source kind mismatch: ${sourceKey(source)}`);
      return paragraph;
    },
  } satisfies BodyAcquisitionInputProjections);
  const acquisition = Object.freeze({
    acquisitionInputs: projections,
    effectiveTablePositioning(table) {
      const fact = tableByIdentity.get(table);
      if (!fact) throw new Error('Unknown table acquisition input');
      return fact.input.format.positioning === null ? null : (table.tblpPr ?? null);
    },
    publicAnchorBridge(source, runIndex) {
      const fact = paragraphBySource.get(sourceKey(source));
      if (!fact) throw new Error(`Unknown paragraph acquisition source: ${sourceKey(source)}`);
      if (!Number.isSafeInteger(runIndex) || runIndex < 0
        || runIndex >= fact.publicAnchorBridges.length) {
        throw new RangeError(`Unknown paragraph anchor bridge index: ${runIndex}`);
      }
      return fact.publicAnchorBridges[runIndex] ?? null;
    },
  } satisfies LayoutSourceAcquisition);
  const immutableNumberSet = (values: readonly number[]): Set<number> => {
    const retained = new Set(values);
    const view = Object.create(null) as Record<PropertyKey, unknown>;
    Object.defineProperties(view, {
      size: { get: () => retained.size },
      has: { value: (value: number) => retained.has(value) },
      entries: { value: () => retained.entries() },
      keys: { value: () => retained.keys() },
      values: { value: () => retained.values() },
      forEach: {
        value: (callback: (value: number, key: number, set: Set<number>) => void, thisArg?: unknown) => {
          retained.forEach((value) => callback.call(thisArg, value, value, view as unknown as Set<number>));
        },
      },
      [Symbol.iterator]: { value: () => retained[Symbol.iterator]() },
      [Symbol.toStringTag]: { value: 'Set' },
    });
    return Object.freeze(view) as unknown as Set<number>;
  };
  const documentLayoutSettings: DocumentLayoutSettings = Object.freeze({
    ...input.documentLayoutFacts,
    kinsoku: Object.freeze({
      enabled: input.documentLayoutFacts.kinsoku.enabled,
      lineStartForbidden: immutableNumberSet(input.documentLayoutFacts.kinsoku.lineStartForbidden),
      lineEndForbidden: immutableNumberSet(input.documentLayoutFacts.kinsoku.lineEndForbidden),
    }),
  });
  const store = Object.freeze({
    blocks,
    bodyLayoutInput,
    section: input.section,
    documentLayoutSettings,
    fonts: input.fonts,
    fontFamilyCharsets: input.fontFamilyCharsets,
    mathOccurrences: input.mathOccurrences,
    imageMetadata: input.imageMetadata,
    hasPaginationFields: input.hasPaginationFields,
    requiresDomVerticalGlyphLayout: input.requiresDomVerticalGlyphLayout,
    fatalParse: input.fatalParse,
    acquisition,
    paintResources: indexSealedPaintResourceDescriptors(input.paintDescriptors),
  }) as LayoutSourceStore;
  layoutSourceStores.add(store);
  return store;
}

/** Model-free constructor: validates and seals builder-owned plain data in place. */
export function sealLayoutSourceStore(input: LayoutSourceStoreInput): LayoutSourceStore {
  const { bodyLayoutInput, ...base } = input;
  const retained = sealPlainData(bodyLayoutInput, 'layout source body input') as BodyLayoutInput;
  return sealLayoutSourceStoreWithBody(base, retained);
}

export interface LayoutBlockRepositoryInput {
  readonly body: readonly LayoutStoryBlock[];
  readonly stories: readonly Readonly<{
    source: SourceRef;
    body: readonly LayoutStoryBlock[];
  }>[];
  readonly footnotes: readonly LayoutStoryNote[];
  readonly endnotes: readonly LayoutStoryNote[];
}

function storyKey(source: Pick<SourceRef, 'story' | 'storyInstance'>): string {
  return `${source.story}:${source.storyInstance}`;
}

/** Seal one source-addressable logical block repository. */
export function createLayoutBlockRepository(
  input: LayoutBlockRepositoryInput,
): LayoutBlockRepository {
  sealPlainData(input.body, 'layout source body blocks');
  sealPlainData(input.stories, 'layout source story blocks');
  sealPlainData(input.footnotes, 'layout source footnotes');
  sealPlainData(input.endnotes, 'layout source endnotes');
  const stories = new Map<string, readonly LayoutStoryBlock[]>();
  for (const { source, body } of input.stories) {
    if (source.path.length !== 0) throw new TypeError('Story repository roots require an empty source path');
    if (source.story !== 'header' && source.story !== 'footer' && source.story !== 'textbox') {
      throw new TypeError(`Unsupported repository story kind: ${source.story}`);
    }
    const key = storyKey(source);
    if (stories.has(key)) throw new TypeError(`Duplicate story source: ${key}`);
    stories.set(key, body);
  }
  const noteMap = (notes: readonly LayoutStoryNote[], kind: string): Map<string, readonly LayoutStoryBlock[]> => {
    const result = new Map<string, readonly LayoutStoryBlock[]>();
    for (const note of notes) {
      if (result.has(note.id)) throw new TypeError(`Duplicate ${kind} story source: ${note.id}`);
      result.set(note.id, note.content);
    }
    return result;
  };
  const footnotes = noteMap(input.footnotes, 'footnote');
  const endnotes = noteMap(input.endnotes, 'endnote');
  const storyRoot = (source: SourceRef): readonly LayoutStoryBlock[] => {
    if (source.path.length !== 0) throw new Error('Story lookup requires a root-only source');
    if (source.story === 'body' && source.storyInstance === 'body') return input.body;
    if (source.story === 'footnote') {
      const body = footnotes.get(source.storyInstance);
      if (body) return body;
    }
    if (source.story === 'endnote') {
      const body = endnotes.get(source.storyInstance);
      if (body) return body;
    }
    const body = stories.get(storyKey(source));
    if (body) return body;
    throw new Error(`Unknown ${source.story} story source: ${source.storyInstance}`);
  };
  const blocks = new Map<string, LayoutFlowBlock>();
  const indexedSources: SourceRef[] = [];
  const indexBody = (body: readonly LayoutStoryBlock[], root: SourceRef, prefix: number[] = []): void => {
    body.forEach((element, elementIndex) => {
      const path = [...prefix, elementIndex];
      if (element.type !== 'paragraph' && element.type !== 'table') return;
      const source = { ...root, path };
      const key = sourceKey(source);
      if (blocks.has(key)) throw new TypeError(`Duplicate block source: ${key}`);
      blocks.set(key, element as LayoutFlowBlock);
      indexedSources.push(Object.freeze({ ...source, path: Object.freeze([...path]) }));
      if (element.type === 'table') element.rows.forEach((row, rowIndex) => row.cells.forEach((cell, cellIndex) => {
        indexBody(cell.content as readonly LayoutStoryBlock[], root, [...path, rowIndex, cellIndex]);
      }));
    });
  };
  indexBody(input.body, { story: 'body', storyInstance: 'body', path: [] });
  for (const { source, body } of input.stories) indexBody(body, source);
  for (const note of input.footnotes) indexBody(note.content, { story: 'footnote', storyInstance: note.id, path: [] });
  for (const note of input.endnotes) indexBody(note.content, { story: 'endnote', storyInstance: note.id, path: [] });
  return Object.freeze({
    body: input.body,
    footnotes: input.footnotes,
    endnotes: input.endnotes,
    sources: Object.freeze(indexedSources),
    resolve(source: SourceRef) {
      const block = blocks.get(sourceKey(source));
      if (!block) throw new Error(`Unknown block source: ${sourceKey(source)}`);
      return block;
    },
    storyRoot,
  });
}

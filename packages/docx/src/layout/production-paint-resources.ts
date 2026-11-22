import type {
  BodyElement,
  DocParagraph,
  DocRun,
  DocTable,
  DocxDocumentModel,
  HeadersFooters,
  ImageRun,
  ShapeRun,
} from '../types.js';
import type { ImageMetadataRecord, MathOccurrence, PictureBulletSizeResolver } from './resources.js';
import type { BodyAcquisitionInputProjections } from './acquisition-input-projections.js';
import { chartResourceKey, imageResourceKey, vmlPictureResourceKey, vmlStrokeResourceKey, vmlEdgeResourceKey } from './source-key.js';
import { documentPageBackgroundStory, documentPageObjects } from '../page-object-wire.js';
import { pageObjectStorySource } from './body-layout-input.js';
import { pageBackgroundStorySource, type PageBackgroundLayoutInput } from './page-background.js';
import {
  createPaintResourceRegistry,
} from './paint-resources.js';
import type {
  ImagePaintResourceDescriptor,
  PaintResourceDescriptor,
  PaintResourceRegistry,
} from './types.js';

export function chartPaintResourceKey(source: import('./types.js').SourceRef): string {
  return chartResourceKey(source);
}

type ImageDescriptorOptions = Partial<ImageRun> & Pick<ImagePaintResourceDescriptor, 'vmlFill' | 'vmlStrokePattern' | 'effectHost' | 'tile' | 'blipDpi'>;

type ImageDescriptorCandidate = Omit<ImagePaintResourceDescriptor, 'intrinsicSize' | 'mimeType'>;

function imageCandidate(
  kind: 'image' | 'picture-bullet',
  resourceKey: string,
  partPath: string,
  run: ImageDescriptorOptions = {},
): ImageDescriptorCandidate {
  return {
    kind,
    resourceKey,
    partPath,
    ...(run.svgImagePath === undefined ? {} : { svgImagePath: run.svgImagePath }),
    ...(run.srcRect == null ? {} : { srcRect: { ...run.srcRect } }),
    ...(run.rotation === undefined ? {} : { rotation: run.rotation }),
    ...(run.flipH === undefined ? {} : { flipH: run.flipH }),
    ...(run.flipV === undefined ? {} : { flipV: run.flipV }),
    ...(run.alpha === undefined ? {} : { alpha: run.alpha }),
    // Rust ImageRun's absent color replacement is serialized null. Preserve
    // every authored string; only None/omission has no decoded effect owner.
    ...(run.colorReplaceFrom == null ? {} : { colorReplaceFrom: run.colorReplaceFrom }),
    ...(run.duotone === undefined ? {} : { duotone: { ...run.duotone } }),
    ...(run.vmlStrokePattern === undefined ? {} : { vmlStrokePattern: run.vmlStrokePattern }),
    ...(run.vmlFill === undefined ? {} : { vmlFill: run.vmlFill }),
    ...(run.tile === undefined ? {} : { tile: { ...run.tile } }),
    ...(run.blipDpi === undefined ? {} : { blipDpi: run.blipDpi }),
    ...(run.blipEffects?.length ? { blipEffects: run.blipEffects.map(effect => ({ ...effect })) } : {}),
    ...(run.effectHost === undefined ? {} : { effectHost: { ...run.effectHost } }),
  };
}

/** Acquired paragraphs registered outside the document's own stories (the
 * selected note separator stories): their source and acquisition runs. */
export type ExtraResourceParagraph = Readonly<{
  source: import('./types.js').SourceRef;
  runs: readonly DocRun[];
}>;

function collectDescriptorCandidates(
  doc: DocxDocumentModel,
  acquisitionInputs: BodyAcquisitionInputProjections | undefined,
  mathOccurrences: readonly MathOccurrence[],
  resolvePictureBulletSize?: PictureBulletSizeResolver,
  pageBackground?: PageBackgroundLayoutInput,
  extraParagraphs: readonly ExtraResourceParagraph[] = [],
): Readonly<{ imageMetadata: ImageMetadataRecord[]; descriptors: PaintResourceDescriptor[] }> {
  const imageCandidates: ImageDescriptorCandidate[] = [];
  const imageMetadata: ImageMetadataRecord[] = [];
  const descriptors: PaintResourceDescriptor[] = [];
  const addImage = (
    kind: 'image' | 'picture-bullet',
    source: import('./types.js').SourceRef,
    partPath: string,
    mimeType: string,
    widthPt: number,
    heightPt: number,
    run: ImageDescriptorOptions = {},
    resourceKey = imageResourceKey(source, partPath),
  ): void => {
    imageCandidates.push(imageCandidate(kind, resourceKey, partPath, run));
    imageMetadata.push({ resourceKey, mimeType, widthPt, heightPt });
  };
  const visitRun = (run: DocRun, source: import('./types.js').SourceRef): void => {
    if (run.type === 'image') {
      addImage('image', source, run.imagePath, run.mimeType, run.widthPt, run.heightPt, run);
      return;
    }
    if (run.type === 'chart') {
      descriptors.push({
        kind: 'chart',
        resourceKey: chartPaintResourceKey(source),
        intrinsicSize: { widthPt: run.widthPt, heightPt: run.heightPt },
        model: run.chart,
      });
      return;
    }
    if (run.type !== 'shape') return;
    const shape = run as ShapeRun & import('./types.js').VmlShapeAcquisitionFacts & Readonly<{ textBoxContent?: BodyElement[] }>;
    if (shape.fill?.fillType === 'image') {
      // §20.1.8.14 authored DPI (private wire); 0 or absence defers to the blip.
      const authoredDpi = (shape.fill as { __dpi?: unknown }).__dpi;
      const blipDpi = typeof authoredDpi === 'number' && Number.isFinite(authoredDpi) && authoredDpi > 0
        ? authoredDpi : undefined;
      addImage(
        'image', source, shape.fill.imagePath, shape.fill.mimeType,
        shape.widthPt, shape.heightPt,
        {
          ...(shape.fill.svgImagePath === undefined ? {} : {
            svgImagePath: shape.fill.svgImagePath,
          }),
          ...(shape.fill.srcRect === undefined ? {} : {
            srcRect: { ...shape.fill.srcRect },
          }),
          ...(shape.fill.alpha === undefined ? {} : { alpha: shape.fill.alpha }),
          ...(shape.fill.duotone === undefined ? {} : { duotone: shape.fill.duotone }),
          ...(shape.fill.blipEffects?.length ? {
            blipEffects: shape.fill.blipEffects,
            effectHost: {
              rotation: shape.rotation, flipH: shape.flipH, flipV: shape.flipV,
              ...(shape.fill.fillRect === undefined ? {} : { fillRect: { ...shape.fill.fillRect } }),
              // Physical effects of a tile sample at its DPI; preparation
              // replaces this with the blip's own DPI when none is authored.
              ...(shape.fill.tile ? { tileDpi: blipDpi ?? 96 } : {}),
            },
          } : {}),
          ...(shape.fill.vml === undefined ? {} : { vmlFill: shape.fill.vml }),
          ...(shape.fill.tile === undefined ? {} : { tile: shape.fill.tile }),
          ...(shape.fill.tile === undefined || blipDpi === undefined ? {} : { blipDpi }),
        },
      );
    }
    if (shape.vmlStrokeFill) {
      const fill = shape.vmlStrokeFill;
      addImage('image', source, fill.imagePath, fill.mimeType, shape.widthPt + (shape.vmlInsetPen === true ? 0 : shape.strokeWidth ?? 0), shape.heightPt + (shape.vmlInsetPen === true ? 0 : shape.strokeWidth ?? 0),
        { alpha: fill.alpha, duotone: fill.duotone, vmlFill: fill.vml, vmlStrokePattern: shape.vmlStrokePattern },
        vmlStrokeResourceKey(source, fill.imagePath));
    }
    for (const edge of shape.vmlEdgeStrokes ?? []) {
      const fill = edge.image;
      if (fill) addImage('image', source, fill.imagePath, fill.mimeType,
        shape.widthPt + edge.widthPt, shape.heightPt + edge.widthPt,
        { alpha: fill.alpha, vmlFill: fill.vml, vmlStrokePattern: edge.pattern ?? undefined },
        vmlEdgeResourceKey(source, edge.edge));
    }
    if (shape.vmlPicture) {
      const picture = shape.vmlPicture;
      // Part 4 §§19.1.2.10/.11/.19: the host owns picture placement.
      // Project resource/crop/effect facts explicitly: Rust serializes a full
      // ImageRun, and a TypeScript Pick does not erase its runtime transforms.
      addImage('image', source, picture.imagePath, picture.mimeType,
        shape.widthPt, shape.heightPt, {
          srcRect: picture.srcRect,
          colorReplaceFrom: picture.colorReplaceFrom,
          duotone: picture.duotone,
          // gain/blacklevel resolve to a §20.1.8.42 luminance correction.
          blipEffects: picture.blipEffects,
        }, vmlPictureResourceKey(source, picture.imagePath));
    }
    const storyInstance = `${source.story}:${source.storyInstance}:${source.path.join('.')}`;
    if (shape.textBoxContent !== undefined) {
      visitBody(shape.textBoxContent, 'textbox', storyInstance);
      return;
    }
    shape.textBlocks?.forEach((block, blockIndex) => {
      if (!block.imagePath) return;
      if (!block.mimeType || block.imageWidthPt == null || block.imageHeightPt == null) {
        throw new Error('Text-box compatibility image requires complete metadata');
      }
      const textBoxSource = {
        story: 'textbox' as const,
        storyInstance,
        // normalizeTextBoxInput projects each compatibility block to a retained
        // paragraph and its optional image to run 0.
        path: [blockIndex, 0],
      };
      addImage(
        'image', textBoxSource, block.imagePath, block.mimeType,
        block.imageWidthPt, block.imageHeightPt,
        { svgImagePath: block.svgImagePath },
      );
    });
  };
  const visitTable = (
    table: DocTable,
    story: import('./types.js').SourceRef['story'],
    storyInstance: string,
    prefix: number[],
  ): void => {
    table.rows.forEach((row, rowIndex) => row.cells.forEach((cell, cellIndex) => {
      visitBody(
        cell.content as BodyElement[],
        story,
        storyInstance,
        [...prefix, rowIndex, cellIndex],
      );
    }));
  };
  const visitParts = (
    parts: HeadersFooters | undefined,
    story: 'header' | 'footer',
    instancePrefix?: string,
  ): void => {
    if (!parts) return;
    for (const kind of ['default', 'first', 'even'] as const) {
      const part = parts[kind];
      if (part) visitBody(part.body, story, instancePrefix ? `${instancePrefix}:${kind}` : kind);
    }
  };
  const visitParagraph = (
    paragraph: DocParagraph,
    source: import('./types.js').SourceRef,
  ): void => {
    const numbering = paragraph.numbering;
    if (numbering?.picBulletImagePath) {
      const size = resolvePictureBulletSize?.(paragraph);
      const widthPt = numbering.picBulletWidthPt ?? size?.widthPt;
      const heightPt = numbering.picBulletHeightPt ?? size?.heightPt;
      if (!numbering.picBulletMimeType || widthPt == null || heightPt == null) {
        throw new Error('Picture bullet requires complete metadata');
      }
      addImage(
        'picture-bullet', source, numbering.picBulletImagePath,
        numbering.picBulletMimeType, widthPt, heightPt,
        numbering.picBulletTransform,
      );
    }
    const acquiredRuns = acquisitionInputs?.paragraphAcquisitionInput(paragraph, source).runs
      ?? paragraph.runs;
    let publicRunIndex = 0;
    acquiredRuns.forEach((acquiredRun, authoredRunIndex) => {
      if (acquiredRun.type === 'unavailableDrawing') return;
      const run = paragraph.runs[publicRunIndex++];
      if (run && (run.type === 'image' || run.type === 'chart' || run.type === 'shape')) {
        visitRun(run, { ...source, path: [...source.path, authoredRunIndex] });
      }
    });
  };
  const visitBody = (
    body: BodyElement[],
    story: import('./types.js').SourceRef['story'],
    storyInstance: string,
    prefix: number[] = [],
  ): void => {
    body.forEach((element, elementIndex) => {
      const path = [...prefix, elementIndex];
      if (element.type === 'paragraph') {
        visitParagraph(element, { story, storyInstance, path });
      } else if (element.type === 'table') {
        visitTable(element, story, storyInstance, path);
      } else if (element.type === 'sectionBreak') {
        visitParts(element.headers, 'header', `section:${elementIndex}`);
        visitParts(element.footers, 'footer', `section:${elementIndex}`);
      }
    });
  };

  visitBody(doc.body, 'body', 'body');
  visitParts(doc.headers, 'header');
  visitParts(doc.footers, 'footer');
  documentPageObjects(doc).forEach((object, ordinal) => {
    visitBody(object.body, 'header', pageObjectStorySource(ordinal).storyInstance);
  });
  visitBody(documentPageBackgroundStory(doc), 'header', pageBackgroundStorySource().storyInstance);
  for (const note of doc.footnotes ?? []) visitBody(note.content, 'footnote', note.id);
  for (const note of doc.endnotes ?? []) visitBody(note.content, 'endnote', note.id);
  for (const { source, runs } of extraParagraphs) {
    runs.forEach((run, runIndex) => {
      if (run.type === 'image' || run.type === 'chart' || run.type === 'shape') {
        visitRun(run, { ...source, path: [...source.path, runIndex] });
      }
    });
  }
  // ECMA-376 §17.2.1 background picture fills: one resource per physical page
  // size, whose extent is the host box of the Part 4 §19.1.2.5 brush.
  for (const fill of pageBackground?.fills ?? []) {
    if (fill.kind !== 'image') continue;
    const image = fill.fill;
    for (const resource of fill.resources) {
      imageCandidates.push(imageCandidate('image', resource.resourceKey, image.imagePath, {
        ...(image.svgImagePath === undefined ? {} : { svgImagePath: image.svgImagePath }),
        ...(image.srcRect === undefined ? {} : { srcRect: { ...image.srcRect } }),
        ...(image.alpha === undefined ? {} : { alpha: image.alpha }),
        ...(image.duotone === undefined ? {} : { duotone: { ...image.duotone } }),
        ...(image.blipEffects?.length ? {
          blipEffects: structuredClone(image.blipEffects) as ImageRun['blipEffects'],
          effectHost: {
            rotation: 0, flipH: false, flipV: false,
            ...(image.fillRect === undefined ? {} : { fillRect: { ...image.fillRect } }),
          },
        } : {}),
        ...(image.vml === undefined ? {} : { vmlFill: image.vml }),
      }));
      imageMetadata.push({
        resourceKey: resource.resourceKey,
        mimeType: image.mimeType,
        widthPt: resource.widthPt,
        heightPt: resource.heightPt,
      });
    }
  }

  const metadataByKey = new Map(imageMetadata.map((record) => [record.resourceKey, record]));
  if (metadataByKey.size !== imageMetadata.length) throw new Error('Duplicate image resource key');
  for (const [documentOrder, candidate] of imageCandidates.entries()) {
    const record = metadataByKey.get(candidate.resourceKey);
    if (!record) throw new Error(`Missing layout image metadata: ${candidate.resourceKey}`);
    descriptors.push({
      ...candidate,
      documentOrder,
      mimeType: record.mimeType,
      intrinsicSize: { widthPt: record.widthPt, heightPt: record.heightPt },
    });
  }
  for (const occurrence of mathOccurrences) {
    descriptors.push({ kind: 'math', resourceKey: occurrence.resourceKey });
  }
  return { imageMetadata, descriptors };
}

export interface DocumentSnapshotResourceProjection {
  readonly imageMetadata: readonly ImageMetadataRecord[];
  readonly paintResources: PaintResourceRegistry;
}

/** One exact source-keyed traversal owns both layout metadata and paint
 * descriptors, including complete rich text-box stories. */
export function projectDocumentSnapshotResources(
  doc: DocxDocumentModel,
  acquisitionInputs: BodyAcquisitionInputProjections | undefined,
  mathOccurrences: readonly MathOccurrence[],
  resolvePictureBulletSize?: PictureBulletSizeResolver,
  pageBackground?: PageBackgroundLayoutInput,
  extraParagraphs: readonly ExtraResourceParagraph[] = [],
): DocumentSnapshotResourceProjection {
  const projected = collectDescriptorCandidates(
    doc, acquisitionInputs, mathOccurrences, resolvePictureBulletSize, pageBackground, extraParagraphs,
  );
  return Object.freeze({
    imageMetadata: Object.freeze(projected.imageMetadata.map((record) => Object.freeze({ ...record }))),
    paintResources: createPaintResourceRegistry(projected.descriptors),
  });
}

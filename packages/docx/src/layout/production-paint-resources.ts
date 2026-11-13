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
import { documentPageObjects } from '../page-object-wire.js';
import { pageObjectStorySource } from './body-layout-input.js';
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

type ImageDescriptorOptions = Partial<ImageRun> & Pick<ImagePaintResourceDescriptor, 'vmlFill' | 'vmlStrokePattern' | 'effectHost'>;

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
    ...(run.blipEffects?.length ? { blipEffects: run.blipEffects.map(effect => ({ ...effect })) } : {}),
    ...(run.effectHost === undefined ? {} : { effectHost: { ...run.effectHost } }),
  };
}

function collectDescriptorCandidates(
  doc: DocxDocumentModel,
  acquisitionInputs: BodyAcquisitionInputProjections | undefined,
  mathOccurrences: readonly MathOccurrence[],
  resolvePictureBulletSize?: PictureBulletSizeResolver,
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
              ...(shape.fill.tile ? { tileDpi: 96 } : {}),
            },
          } : {}),
          ...(shape.fill.vml === undefined ? {} : { vmlFill: shape.fill.vml }),
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
  for (const note of doc.footnotes ?? []) visitBody(note.content, 'footnote', note.id);
  for (const note of doc.endnotes ?? []) visitBody(note.content, 'endnote', note.id);

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
): DocumentSnapshotResourceProjection {
  const projected = collectDescriptorCandidates(
    doc, acquisitionInputs, mathOccurrences, resolvePictureBulletSize,
  );
  return Object.freeze({
    imageMetadata: Object.freeze(projected.imageMetadata.map((record) => Object.freeze({ ...record }))),
    paintResources: createPaintResourceRegistry(projected.descriptors),
  });
}

import { resolveDrawingMLGeometry } from '@silurus/ooxml-core';
import type { DrawingMLShapePaintPlan } from '@silurus/ooxml-core';
import type { SectionLayoutContext } from '../layout-context.js';
import type { ShapeFill } from '../types.js';
import { resolveSectionContextForPage, type SectionPageLayoutPolicy } from './context.js';
import {
  sectionWritingMode,
  uprightPhysicalExtent,
  type PhysicalPageExtent,
} from './coordinate-space.js';
import type {
  DeepReadonly,
  LayoutPage,
  PageBackgroundLayout,
  SourceRef,
} from './types.js';

// ECMA-376 Part 1 §17.2.1 `w:background`: "the image or fill for the entire
// page surface, behind all other document content", for every page. The
// parser emits the private wire only while §17.15.1.26
// `w:displayBackgroundShape` is on and resolves theme colours there; layout
// fixes one full-page command list per physical page before paint.

type ImageFill = DeepReadonly<Extract<ShapeFill, { fillType: 'image' }>>;
type PaintFill = DeepReadonly<Exclude<ShapeFill, { fillType: 'image' }>>;

/** Parser facts: `color` is `RRGGBB`; fills layer above it in order, then
 * the drawings of the background story (`hasStory`). */
export interface PageBackgroundAcquisitionInput {
  readonly color: string | null;
  readonly fills: readonly DeepReadonly<ShapeFill>[];
  readonly hasStory: boolean;
}

/** One retained image resource of a background picture for one physical
 * page size: Part 4 §19.1.2.5 tile/frame sizes are host-relative, so the
 * resolved brush belongs to the page extent it fills. */
export interface PageBackgroundImageResourceInput {
  readonly widthPt: number;
  readonly heightPt: number;
  readonly resourceKey: string;
}

export type PageBackgroundFillInput =
  | Readonly<{ kind: 'paint'; fill: PaintFill }>
  | Readonly<{
      kind: 'image';
      fill: ImageFill;
      resources: readonly PageBackgroundImageResourceInput[];
    }>;

export interface PageBackgroundLayoutInput {
  readonly color: string | null;
  readonly fills: readonly PageBackgroundFillInput[];
  /** The background drawing story, laid out on every page in its
   * `background` layer above `color` and `fills`. */
  readonly story: SourceRef | null;
}

/** The story source of the §17.2.1 background drawing story: a header-class
 * page story laid out on every page in the page's `background` layer. */
export function pageBackgroundStorySource(): SourceRef {
  return Object.freeze({ story: 'header' as const, storyInstance: 'page-background', path: [] });
}

const HEX6 = /^[0-9A-Fa-f]{6}$/u;
const FILL_TYPES = new Set(['solid', 'gradient', 'pattern', 'image']);

function isRecord(value: unknown): value is Readonly<Record<string, unknown>> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

/** Validate the closed private wire. An unknown shape is a producer contract
 * error, never silently rendered as no background. */
export function pageBackgroundAcquisitionInput(
  wire: unknown,
): PageBackgroundAcquisitionInput | null {
  if (wire === undefined) return null;
  if (!isRecord(wire)) throw new Error('Unsupported page background');
  const { color, fills, body } = wire;
  if (body !== undefined && !Array.isArray(body)) throw new Error('Unsupported page background story');
  const hasStory = Array.isArray(body) && body.length > 0;
  if (color !== undefined && (typeof color !== 'string' || !HEX6.test(color))) {
    throw new Error('Unsupported page background color');
  }
  if (fills !== undefined && (!Array.isArray(fills) || !fills.every((fill) => (
    isRecord(fill) && typeof fill.fillType === 'string' && FILL_TYPES.has(fill.fillType)
    && (fill.fillType !== 'image' || (
      typeof fill.imagePath === 'string' && fill.imagePath.length > 0
      && typeof fill.mimeType === 'string'
      // Part 4 VML brushes; DrawingML content arrives as the story.
      && fill.tile === undefined
    ))
  )))) {
    throw new Error('Unsupported page background fill');
  }
  const acquiredFills = (fills ?? []) as readonly DeepReadonly<ShapeFill>[];
  if (color === undefined && acquiredFills.length === 0 && !hasStory) return null;
  return {
    color: color === undefined ? null : color.toUpperCase(),
    fills: acquiredFills,
    hasStory,
  };
}

/** The physical page extent `section` produces (the page geometry layout
 * gives the canvas). Shared by pagination and background resource keys. */
export function sectionPhysicalPageExtent(
  section: DeepReadonly<SectionLayoutContext>,
): PhysicalPageExtent {
  return uprightPhysicalExtent({
    widthPt: section.geometry.pageWidth,
    heightPt: section.geometry.pageHeight,
  }, sectionWritingMode(section));
}

/** Physical extent of every page a section owner produces. Page index only
 * moves margins (mirror margins), never the sheet size. */
export function sectionOwnerPhysicalPageExtent(
  context: DeepReadonly<SectionLayoutContext>,
  pageLayout: DeepReadonly<SectionPageLayoutPolicy>,
): PhysicalPageExtent {
  return sectionPhysicalPageExtent(resolveSectionContextForPage(
    context as SectionLayoutContext,
    pageLayout as SectionPageLayoutPolicy,
    0,
  ));
}

export function pageBackgroundResourceKey(
  fillIndex: number,
  extent: PhysicalPageExtent,
  partPath: string,
): string {
  return `page-background:${fillIndex}:${extent.widthPt}x${extent.heightPt}:${encodeURIComponent(partPath)}`;
}

/** Bind image fills to one retained resource per distinct physical page
 * size the document's sections produce. */
export function projectPageBackgroundLayoutInput(
  acquired: PageBackgroundAcquisitionInput | null | undefined,
  extents: readonly PhysicalPageExtent[],
): PageBackgroundLayoutInput | null {
  if (!acquired) return null;
  const sizes: PhysicalPageExtent[] = [];
  for (const extent of extents) {
    if (!sizes.some((size) => size.widthPt === extent.widthPt && size.heightPt === extent.heightPt)) {
      sizes.push({ widthPt: extent.widthPt, heightPt: extent.heightPt });
    }
  }
  return {
    color: acquired.color,
    story: acquired.hasStory ? pageBackgroundStorySource() : null,
    fills: acquired.fills.map((fill, fillIndex): PageBackgroundFillInput => (
      fill.fillType === 'image'
        ? {
            kind: 'image',
            fill,
            resources: sizes.map((size) => ({
              ...size,
              resourceKey: pageBackgroundResourceKey(fillIndex, size, fill.imagePath),
            })),
          }
        : { kind: 'paint', fill }
    )),
  };
}

function pagePlan(
  extent: PhysicalPageExtent,
  fill: PaintFill | null,
): DeepReadonly<DrawingMLShapePaintPlan & {
  readonly resolvedGeometry: NonNullable<DrawingMLShapePaintPlan['resolvedGeometry']>;
}> {
  const plan: DrawingMLShapePaintPlan = {
    rect: { x: 0, y: 0, w: extent.widthPt, h: extent.heightPt },
    geometry: { kind: 'preset', name: 'rect', adjustments: [] },
    fill: fill === null ? null : structuredClone(fill) as DrawingMLShapePaintPlan['fill'],
    stroke: null,
    transform: { rotationDeg: 0, flipH: false, flipV: false },
  };
  return { ...plan, resolvedGeometry: resolveDrawingMLGeometry(plan, 1) };
}

/** The page surface colour and full-page fill commands of one physical page,
 * painted before any page story, page border or anchored drawing (§17.2.1
 * "behind all other content"). */
export function materializePageBackground(
  input: PageBackgroundLayoutInput,
  extent: PhysicalPageExtent,
): PageBackgroundLayout {
  const commands: PageBackgroundLayout['commands'][number][] = [];
  for (const fill of input.fills) {
    if (fill.kind === 'paint') {
      commands.push({ kind: 'drawingml-shape', plan: pagePlan(extent, fill.fill) });
      continue;
    }
    const resource = fill.resources.find((candidate) => (
      candidate.widthPt === extent.widthPt && candidate.heightPt === extent.heightPt
    ));
    if (!resource) {
      throw new Error(`Page background has no retained image for a ${extent.widthPt}x${extent.heightPt}pt page`);
    }
    commands.push({
      kind: 'drawingml-image-fill',
      plan: pagePlan(extent, null),
      resourceKey: resource.resourceKey,
      ...(fill.fill.fillRect === undefined ? {} : {
        fillRect: {
          l: fill.fill.fillRect.l ?? 0,
          t: fill.fill.fillRect.t ?? 0,
          r: fill.fill.fillRect.r ?? 0,
          b: fill.fill.fillRect.b ?? 0,
        },
      }),
    });
  }
  return { color: input.color === null ? null : `#${input.color}`, commands };
}

/** Attach the document background to a completed page (no-op without one). */
export function withPageBackground(
  page: LayoutPage,
  input: PageBackgroundLayoutInput | null | undefined,
): LayoutPage {
  if (!input || (input.color === null && input.fills.length === 0)) return page;
  return Object.freeze({
    ...page,
    pageBackground: materializePageBackground(input, {
      widthPt: page.geometry.widthPt,
      heightPt: page.geometry.heightPt,
    }),
  });
}

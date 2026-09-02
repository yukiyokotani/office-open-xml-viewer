import {
  OoxmlResourceLimitError,
  type OoxmlResourceUsageSnapshot,
} from '@silurus/ooxml-core';
import {
  cappedAdd,
  measureStructuralJson,
} from '@silurus/ooxml-core/internal/resource-measurement';
import { HARD_MAX_PPTX_PREFLIGHT_PROJECTION_BYTES } from '@silurus/ooxml-core/worker';
import { PptxFontPreloadAccumulator } from './font-plan';
import type {
  MediaElement,
  PptxComment,
  PptxCommentAnchor,
  PptxCommentReply,
  Slide,
} from './types';
import type {
  PresentationBootstrap,
  PresentationBootstrapSlide,
  PptxEmbeddedFontRef,
} from './worker-protocol';

const ZERO_RESOURCE_USAGE: OoxmlResourceUsageSnapshot = Object.freeze({
  archiveEntryCount: 0,
  declaredInflatedBytes: 0,
  distinctInflatedBytes: 0,
  operationInflatedBytes: 0,
});

export const PPTX_MAX_PREFLIGHT_PROJECTION_BYTES =
  HARD_MAX_PPTX_PREFLIGHT_PROJECTION_BYTES;

export interface PresentationPreflightSlide {
  readonly index: number;
  readonly partName?: string;
  readonly notes: string | null;
  readonly hidden: boolean;
  readonly mediaElements: readonly Readonly<MediaElement>[];
  /** Compact slide comments retained for synchronous ScrollViewer UI in both
   * main and worker modes. Omitted for the common comment-free slide. */
  readonly comments?: readonly Readonly<PptxComment>[];
}

/**
 * Compact immutable facts retained for synchronous viewer behavior while full
 * slides are pulled and cached independently. Every field is a direct
 * projection of the canonical Rust Slide/PresentationML model: notes
 * (ECMA-376 Part 1 §13.3.5), hidden state (`p:sld@show`, §19.3.1.38),
 * media geometry/relationships, and `sldIdLst` OPC part identity.
 */
export interface PresentationPreflight {
  readonly slideCount: number;
  readonly slideWidth: number;
  readonly slideHeight: number;
  readonly defaultTextColor: string | null;
  readonly majorFont: string | null;
  readonly minorFont: string | null;
  readonly hlinkColor: string | null;
  readonly folHlinkColor: string | null;
  readonly embeddedFonts: readonly Readonly<PptxEmbeddedFontRef>[];
  readonly slides: readonly PresentationPreflightSlide[];
  readonly fontPreloadNames: readonly (string | null)[];
  readonly fontProviderNames?: readonly string[];
}

function assertNullableString(value: unknown, field: string): asserts value is string | null {
  if (value !== null && typeof value !== 'string') {
    throw new Error(`invalid PPTX presentation bootstrap ${field}`);
  }
}

function copyBootstrapSlide(
  value: unknown,
  expectedIndex: number,
): PresentationBootstrapSlide {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`invalid PPTX presentation bootstrap slide at ${expectedIndex}`);
  }
  const candidate = value as Partial<PresentationBootstrapSlide>;
  if (candidate.index !== expectedIndex) {
    throw new Error(`invalid PPTX presentation bootstrap slide index ${candidate.index}`);
  }
  if (candidate.partName !== undefined && typeof candidate.partName !== 'string') {
    throw new Error(`invalid PPTX presentation bootstrap slide partName at ${expectedIndex}`);
  }
  return Object.freeze({
    index: candidate.index,
    ...(candidate.partName === undefined ? {} : { partName: candidate.partName }),
  });
}

function copyEmbeddedFont(value: unknown, index: number): Readonly<PptxEmbeddedFontRef> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`invalid PPTX presentation bootstrap embedded font at ${index}`);
  }
  const candidate = value as Partial<PptxEmbeddedFontRef>;
  if (
    typeof candidate.fontName !== 'string' || candidate.fontName.length === 0 ||
    !['regular', 'bold', 'italic', 'boldItalic'].includes(candidate.style ?? '') ||
    typeof candidate.partPath !== 'string' || candidate.partPath.length === 0 ||
    candidate.partPath.startsWith('/') || candidate.partPath.split('/').includes('..') ||
    !['application/x-font-ttf', 'application/x-fontdata'].includes(candidate.contentType ?? '')
  ) {
    throw new Error(`invalid PPTX presentation bootstrap embedded font fields at ${index}`);
  }
  return Object.freeze({
    fontName: candidate.fontName,
    style: candidate.style as PptxEmbeddedFontRef['style'],
    partPath: candidate.partPath,
    contentType: candidate.contentType as PptxEmbeddedFontRef['contentType'],
  });
}

/** Validate and detach the JSON-decoded Rust bootstrap from mutable callers. */
export function normalizePresentationBootstrap(
  value: unknown,
): PresentationBootstrap {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error('invalid PPTX presentation bootstrap payload');
  }
  const candidate = value as Partial<PresentationBootstrap>;
  if (
    !Number.isSafeInteger(candidate.slideCount) || (candidate.slideCount ?? -1) < 0 ||
    !Number.isSafeInteger(candidate.slideWidth) || (candidate.slideWidth ?? 0) <= 0 ||
    !Number.isSafeInteger(candidate.slideHeight) || (candidate.slideHeight ?? 0) <= 0 ||
    !Array.isArray(candidate.embeddedFonts) ||
    !Array.isArray(candidate.slides) ||
    candidate.slides.length !== candidate.slideCount
  ) {
    throw new Error('invalid PPTX presentation bootstrap dimensions or slide count');
  }
  assertNullableString(candidate.defaultTextColor, 'defaultTextColor');
  assertNullableString(candidate.majorFont, 'majorFont');
  assertNullableString(candidate.minorFont, 'minorFont');
  assertNullableString(candidate.hlinkColor, 'hlinkColor');
  assertNullableString(candidate.folHlinkColor, 'folHlinkColor');
  return Object.freeze({
    slideCount: candidate.slideCount as number,
    slideWidth: candidate.slideWidth as number,
    slideHeight: candidate.slideHeight as number,
    defaultTextColor: candidate.defaultTextColor,
    majorFont: candidate.majorFont,
    minorFont: candidate.minorFont,
    hlinkColor: candidate.hlinkColor,
    folHlinkColor: candidate.folHlinkColor,
    embeddedFonts: Object.freeze(
      (candidate.embeddedFonts as readonly unknown[]).map(copyEmbeddedFont),
    ),
    slides: Object.freeze(candidate.slides.map(copyBootstrapSlide)),
  });
}

function copyMediaElement(element: MediaElement): Readonly<MediaElement> {
  return Object.freeze({
    type: 'media',
    x: element.x,
    y: element.y,
    width: element.width,
    height: element.height,
    rotation: element.rotation,
    flipH: element.flipH,
    flipV: element.flipV,
    mediaKind: element.mediaKind,
    posterPath: element.posterPath,
    posterMimeType: element.posterMimeType,
    mediaPath: element.mediaPath,
    mimeType: element.mimeType,
  });
}

function copyCommentReply(reply: PptxCommentReply): Readonly<PptxCommentReply> {
  return Object.freeze({
    ...(reply.id === undefined ? {} : { id: reply.id }),
    ...(reply.authorId === undefined ? {} : { authorId: reply.authorId }),
    ...(reply.author === undefined ? {} : { author: reply.author }),
    ...(reply.date === undefined ? {} : { date: reply.date }),
    ...(reply.status === undefined ? {} : { status: reply.status }),
    text: reply.text,
  });
}

function copyCommentAnchor(anchor: PptxCommentAnchor): Readonly<PptxCommentAnchor> {
  return Object.freeze({ ...anchor });
}

function copyComment(comment: PptxComment): Readonly<PptxComment> {
  return Object.freeze({
    ...(comment.authorId === undefined ? {} : { authorId: comment.authorId }),
    ...(comment.modernAuthorId === undefined ? {} : { modernAuthorId: comment.modernAuthorId }),
    ...(comment.id === undefined ? {} : { id: comment.id }),
    ...(comment.index === undefined ? {} : { index: comment.index }),
    ...(comment.author === undefined ? {} : { author: comment.author }),
    ...(comment.date === undefined ? {} : { date: comment.date }),
    ...(comment.x === undefined ? {} : { x: comment.x }),
    ...(comment.y === undefined ? {} : { y: comment.y }),
    ...(comment.anchors?.length
      ? { anchors: Object.freeze(comment.anchors.map(copyCommentAnchor)) }
      : {}),
    ...(comment.status === undefined ? {} : { status: comment.status }),
    text: comment.text,
    ...(comment.replies?.length
      ? { replies: Object.freeze(comment.replies.map(copyCommentReply)) }
      : {}),
  });
}

function assertOptionalString(value: unknown, field: string, slideIndex: number): void {
  if (value !== undefined && typeof value !== 'string') {
    throw new Error(`invalid PPTX presentation preflight comment ${field} at slide ${slideIndex}`);
  }
}

function normalizeCommentReply(value: unknown, slideIndex: number): Readonly<PptxCommentReply> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`invalid PPTX presentation preflight comment reply at slide ${slideIndex}`);
  }
  const reply = value as Partial<PptxCommentReply>;
  for (const field of ['id', 'authorId', 'author', 'date', 'status'] as const) {
    assertOptionalString(reply[field], field, slideIndex);
  }
  if (typeof reply.text !== 'string') {
    throw new Error(`invalid PPTX presentation preflight comment reply text at slide ${slideIndex}`);
  }
  if (reply.status !== undefined && !['active', 'resolved', 'closed'].includes(reply.status)) {
    throw new Error(`invalid PPTX presentation preflight comment reply status at slide ${slideIndex}`);
  }
  return copyCommentReply(reply as PptxCommentReply);
}

function normalizeComment(value: unknown, slideIndex: number): Readonly<PptxComment> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`invalid PPTX presentation preflight comment at slide ${slideIndex}`);
  }
  const comment = value as Partial<PptxComment>;
  for (const field of ['modernAuthorId', 'id', 'author', 'date', 'status'] as const) {
    assertOptionalString(comment[field], field, slideIndex);
  }
  for (const field of ['authorId', 'index', 'x', 'y'] as const) {
    const item = comment[field];
    if (item !== undefined && (typeof item !== 'number' || !Number.isSafeInteger(item))) {
      throw new Error(`invalid PPTX presentation preflight comment ${field} at slide ${slideIndex}`);
    }
  }
  if (typeof comment.text !== 'string' ||
      (comment.replies !== undefined && !Array.isArray(comment.replies)) ||
      (comment.anchors !== undefined && !Array.isArray(comment.anchors))) {
    throw new Error(`invalid PPTX presentation preflight comment fields at slide ${slideIndex}`);
  }
  if (comment.status !== undefined && !['active', 'resolved', 'closed'].includes(comment.status)) {
    throw new Error(`invalid PPTX presentation preflight comment status at slide ${slideIndex}`);
  }
  return copyComment({
    ...(comment as PptxComment),
    ...(comment.anchors?.length
      ? { anchors: comment.anchors.map((anchor) => normalizeCommentAnchor(anchor, slideIndex)) }
      : {}),
    ...(comment.replies?.length
      ? { replies: comment.replies.map((reply) => normalizeCommentReply(reply, slideIndex)) }
      : {}),
  });
}

function normalizeCommentAnchor(value: unknown, slideIndex: number): Readonly<PptxCommentAnchor> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`invalid PPTX presentation preflight comment anchor at slide ${slideIndex}`);
  }
  const anchor = value as Partial<PptxCommentAnchor>;
  if (anchor.type === 'slide' || anchor.type === 'unknown') return Object.freeze({ type: anchor.type });
  if (anchor.type === 'drawingElement') {
    assertOptionalString(anchor.elementId, 'anchor.elementId', slideIndex);
    assertOptionalString(anchor.creationId, 'anchor.creationId', slideIndex);
    return Object.freeze({
      type: 'drawingElement',
      ...(anchor.elementId === undefined ? {} : { elementId: anchor.elementId }),
      ...(anchor.creationId === undefined ? {} : { creationId: anchor.creationId }),
    });
  }
  if (anchor.type === 'textRange') {
    assertOptionalString(anchor.elementId, 'anchor.elementId', slideIndex);
    for (const field of ['start', 'length'] as const) {
      const item = anchor[field];
      if (item !== undefined && (typeof item !== 'number' || !Number.isSafeInteger(item))) {
        throw new Error(`invalid PPTX presentation preflight comment anchor.${field} at slide ${slideIndex}`);
      }
    }
    return Object.freeze({
      type: 'textRange',
      ...(anchor.elementId === undefined ? {} : { elementId: anchor.elementId }),
      ...(anchor.start === undefined ? {} : { start: anchor.start }),
      ...(anchor.length === undefined ? {} : { length: anchor.length }),
    });
  }
  throw new Error(`invalid PPTX presentation preflight comment anchor type at slide ${slideIndex}`);
}

function normalizeMediaElement(value: unknown, slideIndex: number): Readonly<MediaElement> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`invalid PPTX presentation preflight media at slide ${slideIndex}`);
  }
  const element = value as Partial<MediaElement>;
  for (const field of ['x', 'y', 'width', 'height', 'rotation'] as const) {
    if (typeof element[field] !== 'number' || !Number.isFinite(element[field])) {
      throw new Error(`invalid PPTX presentation preflight media ${field} at slide ${slideIndex}`);
    }
  }
  if (
    element.type !== 'media' ||
    typeof element.flipH !== 'boolean' ||
    typeof element.flipV !== 'boolean' ||
    (element.mediaKind !== 'audio' && element.mediaKind !== 'video') ||
    typeof element.posterPath !== 'string' ||
    typeof element.posterMimeType !== 'string' ||
    typeof element.mediaPath !== 'string' ||
    typeof element.mimeType !== 'string'
  ) {
    throw new Error(`invalid PPTX presentation preflight media fields at slide ${slideIndex}`);
  }
  return copyMediaElement(element as MediaElement);
}

function normalizePresentationPreflightValue(
  value: unknown,
  allowPartial: boolean,
): PresentationPreflight {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error('invalid PPTX presentation preflight payload');
  }
  const candidate = value as Partial<PresentationPreflight>;
  if (
    !Number.isSafeInteger(candidate.slideCount) || (candidate.slideCount ?? -1) < 0 ||
    !Number.isSafeInteger(candidate.slideWidth) || (candidate.slideWidth ?? 0) <= 0 ||
    !Number.isSafeInteger(candidate.slideHeight) || (candidate.slideHeight ?? 0) <= 0 ||
    !Array.isArray(candidate.embeddedFonts) ||
    !Array.isArray(candidate.slides) ||
    (allowPartial
      ? candidate.slides.length > (candidate.slideCount ?? -1)
      : candidate.slides.length !== candidate.slideCount) ||
    !Array.isArray(candidate.fontPreloadNames) ||
    (candidate.fontProviderNames !== undefined && !Array.isArray(candidate.fontProviderNames))
  ) {
    throw new Error('invalid PPTX presentation preflight dimensions or slide count');
  }
  assertNullableString(candidate.defaultTextColor, 'defaultTextColor');
  assertNullableString(candidate.majorFont, 'majorFont');
  assertNullableString(candidate.minorFont, 'minorFont');
  assertNullableString(candidate.hlinkColor, 'hlinkColor');
  assertNullableString(candidate.folHlinkColor, 'folHlinkColor');
  const slides = candidate.slides.map((value, index): PresentationPreflightSlide => {
    if (!value || typeof value !== 'object' || Array.isArray(value)) {
      throw new Error(`invalid PPTX presentation preflight slide at ${index}`);
    }
    const slide = value as Partial<PresentationPreflightSlide>;
    if (
      slide.index !== index ||
      (slide.partName !== undefined && typeof slide.partName !== 'string') ||
      (slide.notes !== null && typeof slide.notes !== 'string') ||
      typeof slide.hidden !== 'boolean' ||
      !Array.isArray(slide.mediaElements)
      || (slide.comments !== undefined && !Array.isArray(slide.comments))
    ) {
      throw new Error(`invalid PPTX presentation preflight slide fields at ${index}`);
    }
    return Object.freeze({
      index,
      ...(slide.partName === undefined ? {} : { partName: slide.partName }),
      notes: slide.notes,
      hidden: slide.hidden,
      mediaElements: Object.freeze(
        slide.mediaElements.map((media) => normalizeMediaElement(media, index)),
      ),
      ...(slide.comments?.length
        ? { comments: Object.freeze(slide.comments.map((comment) => normalizeComment(comment, index))) }
        : {}),
    });
  });
  const fontPreloadNames = candidate.fontPreloadNames.map((name, index) => {
    if (name !== null && typeof name !== 'string') {
      throw new Error(`invalid PPTX presentation preflight font at ${index}`);
    }
    return name;
  });
  const fontProviderNames = (candidate.fontProviderNames ?? []).map((name, index) => {
    if (typeof name !== 'string') {
      throw new Error(`invalid PPTX presentation provider font at ${index}`);
    }
    return name;
  });
  return Object.freeze({
    slideCount: candidate.slideCount as number,
    slideWidth: candidate.slideWidth as number,
    slideHeight: candidate.slideHeight as number,
    defaultTextColor: candidate.defaultTextColor,
    majorFont: candidate.majorFont,
    minorFont: candidate.minorFont,
    hlinkColor: candidate.hlinkColor,
    folHlinkColor: candidate.folHlinkColor,
    embeddedFonts: Object.freeze(
      (candidate.embeddedFonts as readonly unknown[]).map(copyEmbeddedFont),
    ),
    slides: Object.freeze(slides),
    fontPreloadNames: Object.freeze(fontPreloadNames),
    fontProviderNames: Object.freeze(fontProviderNames),
  });
}

/** Validate, detach, and freeze an authoritative compact preflight. */
export function normalizePresentationPreflight(value: unknown): PresentationPreflight {
  return normalizePresentationPreflightValue(value, false);
}

/** Validate a sequential compact prefix pushed by a progressive render worker. */
export function normalizePresentationPreflightPrefix(value: unknown): PresentationPreflight {
  return normalizePresentationPreflightValue(value, true);
}

export function findPreflightMimeType(
  preflight: PresentationPreflight,
  partPath: string,
): string {
  for (const slide of preflight.slides) {
    for (const media of slide.mediaElements) {
      if (media.mediaPath === partPath) return media.mimeType;
      if (media.posterPath === partPath) return media.posterMimeType;
    }
  }
  return '';
}

function projectSlide(
  slide: Slide,
  descriptor: PresentationBootstrapSlide,
): PresentationPreflightSlide {
  if (slide.index !== descriptor.index || slide.partName !== descriptor.partName) {
    throw new Error(`PPTX pulled slide identity does not match bootstrap index ${descriptor.index}`);
  }
  return Object.freeze({
    index: descriptor.index,
    ...(descriptor.partName === undefined ? {} : { partName: descriptor.partName }),
    notes: slide.notes ?? null,
    hidden: slide.hidden ?? false,
    mediaElements: Object.freeze(
      slide.elements
        .filter((element): element is MediaElement => element.type === 'media')
        .map(copyMediaElement),
    ),
    ...(slide.comments?.length
      ? { comments: Object.freeze(slide.comments.map(copyComment)) }
      : {}),
  });
}

export function assertPresentationPreflightProjectionBytes(
  observed: number,
  usage: OoxmlResourceUsageSnapshot = ZERO_RESOURCE_USAGE,
): void {
  assertProjectionBytes(observed, PPTX_MAX_PREFLIGHT_PROJECTION_BYTES, usage);
}

function assertProjectionBytes(
  observed: number,
  limit: number,
  usage: OoxmlResourceUsageSnapshot,
): void {
  if (observed <= limit) return;
  throw new OoxmlResourceLimitError(
    `PPTX presentation preflight exceeded its hard limit of ${limit} projected bytes`,
    {
      stage: 'parsing',
      violation: {
        format: 'pptx',
        operation: 'presentation-preflight',
        resource: 'presentation-preflight',
        metric: 'projected-bytes',
        limit,
        observed: Math.min(observed, limit + 1),
        configurable: false,
        usage,
      },
    },
  );
}

/** Transaction returned directly from a SlidePullWorker `acceptSlide` hook. */
export interface PreparedPresentationPreflightSlide {
  /** Conservative retained-state projection while the candidate awaits ACK. */
  readonly projectedBytes: number;
  rollback(): void;
  commit(): void;
}

interface PendingAcceptance {
  state: 'prepared' | 'committed' | 'rolled-back';
  readonly fact: PresentationPreflightSlide;
  readonly fonts: PptxFontPreloadAccumulator;
  readonly fontNames: readonly (string | null)[];
  readonly providerNames: readonly string[];
  readonly fontBytes: number;
  readonly committedBytes: number;
}

/** @internal Test-only lowering; production cannot raise or replace the hard ceiling. */
export interface PresentationPreflightBuilderOptions {
  readonly hardLimitForTesting?: number;
}

/**
 * Sequential admission builder. It never retains a source Slide: each accepted
 * unit contributes immutable compact facts and script flags, then can be ACKed
 * and released by the pull-session owner.
 */
export class PresentationPreflightBuilder {
  private readonly slideCountValue: number;
  private readonly slideWidthValue: number;
  private readonly slideHeightValue: number;
  private readonly defaultTextColorValue: string | null;
  private readonly majorFontValue: string | null;
  private readonly minorFontValue: string | null;
  private readonly hlinkColorValue: string | null;
  private readonly folHlinkColorValue: string | null;
  private readonly embeddedFontsValue: readonly Readonly<PptxEmbeddedFontRef>[];
  private descriptors: (PresentationBootstrapSlide | undefined)[];
  private slides: PresentationPreflightSlide[] = [];
  private fonts: PptxFontPreloadAccumulator;
  private fontPreloadNames: readonly (string | null)[];
  private fontProviderNames: readonly string[];
  private fontProjectionBytes: number;
  private projectionBytesValue: number;
  private readonly limit: number;
  private pending: PendingAcceptance | null = null;
  private finished: PresentationPreflight | null = null;

  constructor(
    bootstrap: PresentationBootstrap,
    options: PresentationPreflightBuilderOptions = {},
  ) {
    const normalized = normalizePresentationBootstrap(bootstrap);
    const requestedLimit = options.hardLimitForTesting ?? PPTX_MAX_PREFLIGHT_PROJECTION_BYTES;
    if (
      !Number.isSafeInteger(requestedLimit) || requestedLimit <= 0 ||
      requestedLimit > PPTX_MAX_PREFLIGHT_PROJECTION_BYTES
    ) {
      throw new Error('invalid PPTX presentation preflight test limit');
    }
    this.limit = requestedLimit;
    this.slideCountValue = normalized.slideCount;
    this.slideWidthValue = normalized.slideWidth;
    this.slideHeightValue = normalized.slideHeight;
    this.defaultTextColorValue = normalized.defaultTextColor;
    this.majorFontValue = normalized.majorFont;
    this.minorFontValue = normalized.minorFont;
    this.hlinkColorValue = normalized.hlinkColor;
    this.folHlinkColorValue = normalized.folHlinkColor;
    this.embeddedFontsValue = normalized.embeddedFonts;
    this.descriptors = [...normalized.slides];
    this.fonts = new PptxFontPreloadAccumulator(
      this.majorFontValue,
      this.minorFontValue,
    );
    this.fontPreloadNames = Object.freeze(this.fonts.names());
    this.fontProviderNames = Object.freeze(this.fonts.providerNames());
    this.fontProjectionBytes = measureStructuralJson(
      { fontPreloadNames: this.fontPreloadNames, fontProviderNames: this.fontProviderNames },
      this.limit,
    ).jsonBytes;
    this.projectionBytesValue = measureStructuralJson({
      slideCount: this.slideCountValue,
      slideWidth: this.slideWidthValue,
      slideHeight: this.slideHeightValue,
      defaultTextColor: this.defaultTextColorValue,
      majorFont: this.majorFontValue,
      minorFont: this.minorFontValue,
      hlinkColor: this.hlinkColorValue,
      folHlinkColor: this.folHlinkColorValue,
      embeddedFonts: this.embeddedFontsValue,
      remainingSlides: this.descriptors,
      slides: [],
      fontPreloadNames: this.fontPreloadNames,
      fontProviderNames: this.fontProviderNames,
    }, this.limit).jsonBytes;
    assertProjectionBytes(this.projectionBytesValue, this.limit, ZERO_RESOURCE_USAGE);
  }

  get acceptedSlideCount(): number {
    return this.finished?.slideCount ?? this.slides.length;
  }

  get projectedBytes(): number {
    return this.projectionBytesValue;
  }

  get remainingDescriptorCount(): number {
    return this.descriptors.reduce((count, descriptor) => count + Number(descriptor !== undefined), 0);
  }

  /** Latest immutable per-slide fact committed by the sequential cursor. */
  get latestSlide(): PresentationPreflightSlide | undefined {
    return this.slides[this.slides.length - 1];
  }

  /** Current cumulative font request set for the committed slide prefix. */
  get currentFontPreloadNames(): readonly (string | null)[] {
    return this.fontPreloadNames;
  }

  get currentFontProviderNames(): readonly string[] {
    return this.fontProviderNames;
  }

  /**
   * Read-only snapshot of the committed prefix while preflight is still open.
   * `slideCount` remains the final bootstrap count; `slides.length` is the
   * number currently paintable. The snapshot is created only for a consumer
   * that needs the current compact facts, not once per cursor step.
   */
  snapshot(): PresentationPreflight {
    if (this.finished) return this.finished;
    if (this.pending) throw new Error('PPTX presentation preflight has an uncommitted slide');
    return Object.freeze({
      slideCount: this.slideCountValue,
      slideWidth: this.slideWidthValue,
      slideHeight: this.slideHeightValue,
      defaultTextColor: this.defaultTextColorValue,
      majorFont: this.majorFontValue,
      minorFont: this.minorFontValue,
      hlinkColor: this.hlinkColorValue,
      folHlinkColor: this.folHlinkColorValue,
      embeddedFonts: this.embeddedFontsValue,
      slides: Object.freeze([...this.slides]),
      fontPreloadNames: this.fontPreloadNames,
      fontProviderNames: this.fontProviderNames,
    });
  }

  addSlide(
    slide: Slide,
    usage: OoxmlResourceUsageSnapshot = ZERO_RESOURCE_USAGE,
  ): void {
    this.prepareSlide(slide, usage).commit();
  }

  prepareSlide(
    slide: Slide,
    usage: OoxmlResourceUsageSnapshot = ZERO_RESOURCE_USAGE,
  ): PreparedPresentationPreflightSlide {
    if (this.finished) throw new Error('PPTX presentation preflight is already finished');
    if (this.pending) throw new Error('PPTX presentation preflight already has a prepared slide');
    const index = this.slides.length;
    const descriptor = this.descriptors[index];
    if (!descriptor) throw new Error('PPTX presentation preflight received an extra slide');
    const fact = projectSlide(slide, descriptor);
    const nextFonts = this.fonts.withSlide(slide);
    const nextFontNames = Object.freeze(nextFonts.names());
    const nextProviderNames = Object.freeze(nextFonts.providerNames());
    const nextFontBytes = measureStructuralJson(
      { fontPreloadNames: nextFontNames, fontProviderNames: nextProviderNames },
      this.limit,
    ).jsonBytes;
    const slideBytes = measureStructuralJson(
      fact,
      this.limit,
    ).jsonBytes;
    // Commit replaces the current descriptor with JSON `null`, releases its
    // strings, and adds one fact. This exact delta keeps the building-state
    // projection honest while descriptors and committed facts coexist.
    let committedBytes = this.projectionBytesValue
      - this.fontProjectionBytes
      - measureStructuralJson(descriptor, this.limit).jsonBytes
      + 4;
    committedBytes = cappedAdd(committedBytes, nextFontBytes, this.limit);
    committedBytes = cappedAdd(committedBytes, slideBytes, this.limit);
    if (this.slides.length !== 0) {
      committedBytes = cappedAdd(committedBytes, 1, this.limit);
    }
    // Before Rust ACK, both the unchanged builder and candidate acceptance
    // state are live. Charge the complete candidate projection in addition to
    // the existing retained state, then enforce the worse of prepare/commit.
    const candidateBytes = measureStructuralJson({
      slide: fact,
      fontPreloadNames: nextFontNames,
      fontProviderNames: nextProviderNames,
    }, this.limit).jsonBytes;
    const preparedBytes = cappedAdd(this.projectionBytesValue, candidateBytes, this.limit);
    const observed = Math.max(preparedBytes, committedBytes);
    assertProjectionBytes(observed, this.limit, usage);
    const pending: PendingAcceptance = {
      state: 'prepared',
      fact,
      fonts: nextFonts,
      fontNames: nextFontNames,
      providerNames: nextProviderNames,
      fontBytes: nextFontBytes,
      committedBytes,
    };
    this.pending = pending;
    return {
      projectedBytes: preparedBytes,
      commit: () => {
        if (pending.state === 'committed') return;
        if (pending.state === 'rolled-back') {
          throw new Error('PPTX presentation preflight cannot commit a rolled-back slide');
        }
        if (this.pending !== pending) {
          throw new Error('PPTX presentation preflight prepared slide is stale');
        }
        this.descriptors[index] = undefined;
        this.slides.push(pending.fact);
        this.fonts = pending.fonts;
        this.fontPreloadNames = pending.fontNames;
        this.fontProviderNames = pending.providerNames;
        this.fontProjectionBytes = pending.fontBytes;
        this.projectionBytesValue = pending.committedBytes;
        pending.state = 'committed';
        this.pending = null;
      },
      rollback: () => {
        if (pending.state === 'rolled-back') return;
        if (pending.state === 'committed') {
          throw new Error('PPTX presentation preflight cannot roll back a committed slide');
        }
        if (this.pending !== pending) {
          throw new Error('PPTX presentation preflight prepared slide is stale');
        }
        pending.state = 'rolled-back';
        this.pending = null;
      },
    };
  }

  finish(): PresentationPreflight {
    if (this.finished) return this.finished;
    if (this.pending) throw new Error('PPTX presentation preflight has an uncommitted slide');
    if (this.slides.length !== this.slideCountValue) {
      throw new Error(
        `PPTX presentation preflight is incomplete: ${this.slides.length}/${this.slideCountValue} slides`,
      );
    }
    this.finished = Object.freeze({
      slideCount: this.slideCountValue,
      slideWidth: this.slideWidthValue,
      slideHeight: this.slideHeightValue,
      defaultTextColor: this.defaultTextColorValue,
      majorFont: this.majorFontValue,
      minorFont: this.minorFontValue,
      hlinkColor: this.hlinkColorValue,
      folHlinkColor: this.folHlinkColorValue,
      embeddedFonts: this.embeddedFontsValue,
      slides: Object.freeze([...this.slides]),
      fontPreloadNames: this.fontPreloadNames,
      fontProviderNames: this.fontProviderNames,
    });
    // The frozen compact model owns its slide-array storage from here. The
    // builder releases both construction-only arrays rather than retaining a
    // second array of fact references or descriptor slots after finish.
    this.descriptors = [];
    this.slides = [];
    this.projectionBytesValue = measureStructuralJson(
      this.finished,
      this.limit,
    ).jsonBytes;
    return this.finished;
  }
}

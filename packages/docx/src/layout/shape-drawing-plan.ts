import { resolveVmlGeometryPass } from './vml-geometry-pass.js';
import { acquireVmlGroupVectors } from './vml-group-vectors.js';
import { acquireVmlTextContours } from './vml-text-outline.js';
import { acquireVmlSkew, acquireVmlSkewMatrix } from './vml-skew.js';
import { acquireVmlConcentricFill } from './vml-concentric-fill.js';
import { planVmlPathText } from './vml-text-path.js';
import { planVmlShadows } from './vml-shadow.js';
import { resolveDrawingMLGeometry, drawingMLGeometryBounds } from '@silurus/ooxml-core';
import type { ArrowEnd, DrawingMLShapePaintPlan, Stroke } from '@silurus/ooxml-core';
import type { ShapeRun } from '../types.js';
import type {
  DrawingPaintCommand,
  DeepReadonly,
  LayoutDiagnostic,
  LayoutRect,
  VmlTextPathAcquisitionInput,
  VmlShapeAcquisitionFacts,
} from './types.js';
import { snapshotPlainData } from './plain-data.js';
import type { TextLayoutService } from './text.js';

type ShapeDrawingCommand = Extract<
  DrawingPaintCommand,
  { kind: 'noop' | 'drawingml-shape' | 'drawingml-image-fill' | 'watermark-text' }
>;

export type ShapeDrawingPlanResult =
  | Readonly<{
      status: 'planned';
      command: ShapeDrawingCommand;
      before?: readonly Extract<DrawingPaintCommand, { kind: 'drawingml-shape' | 'drawingml-image-fill' | 'watermark-text' }>[];
      diagnostics?: readonly Readonly<Omit<LayoutDiagnostic, 'source'>>[];
    }>
  | Readonly<{
      status: 'unsupported';
      command: Extract<DrawingPaintCommand, { kind: 'noop' }>;
      diagnostics: readonly Readonly<Omit<LayoutDiagnostic, 'source'>>[];
    }>;

// When fitshape is active, every shaped dimension is multiplied by the same
// target/source ratio, so a 1pt internal reference cancels out exactly. This is
// used only when VML omitted font-size; an authored size is never replaced.
const SCALE_NEUTRAL_REFERENCE_FONT_SIZE_PT = 1;

function unsupportedTextPathPlan(
  message: string,
): Extract<ShapeDrawingPlanResult, { status: 'unsupported' }> {
  return Object.freeze({
    status: 'unsupported',
    command: Object.freeze({ kind: 'noop' }),
    diagnostics: Object.freeze([Object.freeze({
      code: 'UNSUPPORTED_FEATURE',
      severity: 'error',
      message,
    })]),
  });
}

function arrowEnd(end: ShapeRun['headEnd']): ArrowEnd | undefined {
  return end ? { type: end.type, w: end.w, len: end.len } : undefined;
}

function strokeFill(fill: DeepReadonly<ShapeRun['strokeFill']>): Stroke['fill'] | undefined {
  if (!fill) return undefined;
  if (fill.fillType === 'gradient') {
    return {
      fillType: 'gradient',
      stops: fill.stops.map((stop) => ({ position: stop.position, color: stop.color })),
      angle: fill.angle,
      gradType: fill.gradType,
      ...(fill.scaled === undefined ? {} : { scaled: fill.scaled }),
      ...(fill.path === undefined ? {} : { path: fill.path }),
      ...(fill.fillToRect === undefined ? {} : { fillToRect: { ...fill.fillToRect } }),
      ...(fill.tileRect === undefined ? {} : { tileRect: { ...fill.tileRect } }),
      ...(fill.flip === undefined ? {} : { flip: fill.flip }),
      ...(fill.rotWithShape === undefined ? {} : { rotWithShape: fill.rotWithShape }),
    };
  }
  return { fillType: 'pattern', fg: fill.fg, bg: fill.bg, preset: fill.preset };
}

function shapeStroke(shape: DeepReadonly<ShapeRun & VmlShapeAcquisitionFacts>): Stroke | null {
  if (!shape.stroke || !shape.strokeWidth || shape.strokeWidth <= 0) return null;
  const fill = strokeFill(shape.strokeFill);
  return {
    color: shape.stroke,
    width: shape.strokeWidth,
    ...(fill ? { fill } : {}),
    ...(shape.strokeDash ? { dashStyle: shape.strokeDash } : {}),
    ...(shape.strokeCustomDash?.length ? { customDash: shape.strokeCustomDash } : {}),
    ...(shape.strokeCap ? { lineCap: shape.strokeCap } : {}),
    ...(shape.strokeJoin ? { lineJoin: shape.strokeJoin } : {}),
    ...(shape.strokeMiterLimit !== undefined && shape.strokeMiterLimit !== null
      ? { miterLimit: shape.strokeMiterLimit }
      : {}),
    ...(shape.strokeAlignment ? { alignment: shape.strokeAlignment } : {}),
    ...(shape.strokeCompound ? { cmpd: shape.strokeCompound } : {}),
    ...(arrowEnd(shape.headEnd) ? { headEnd: arrowEnd(shape.headEnd) } : {}),
    ...(arrowEnd(shape.tailEnd) ? { tailEnd: arrowEnd(shape.tailEnd) } : {}),
  };
}

function planShapeDrawingBase(
  shape: DeepReadonly<ShapeRun & VmlShapeAcquisitionFacts>,
  bounds: LayoutRect,
  text?: TextLayoutService,
  textPath?: Readonly<VmlTextPathAcquisitionInput>,
  imageFillResourceKey?: string,
  maximumGeometryWork?: number,
  vmlPictureResourceKey?: string,
  vmlStrokeResourceKey?: string,
): ShapeDrawingPlanResult {
  const parserControlled = textPath !== undefined && (
    textPath.textPathOk !== undefined
    || textPath.on !== undefined
    || textPath.fitShape !== undefined
    || textPath.fitPath !== undefined
    || textPath.trim !== undefined
    || textPath.xScale !== undefined
  );
  // The stable public ShapeRun predates the private VML control projection.
  // Absence therefore keeps its historical visible+fitshape interpretation;
  // parser-created values always carry explicit false defaults and follow VML.
  const textPathEnabled = textPath !== undefined && (
    parserControlled
      ? textPath.textPathOk === true && textPath.on === true
      : true
  );
  if (textPathEnabled) {
    if (textPath.string.trim().length === 0) {
      return Object.freeze({ status: 'planned', command: Object.freeze({ kind: 'noop' }) });
    }
    if (!text) throw new Error('Shape textPath acquisition requires TextLayoutService');
    const fitShape = parserControlled ? textPath.fitShape === true : true;
    const fitPath = textPath.fitPath === true;
    if (textPath.fontSizePt !== undefined
      && (!Number.isFinite(textPath.fontSizePt) || textPath.fontSizePt < 0)) {
      throw new RangeError('VML textPath fontSizePt must be finite and non-negative');
    }
    if (!fitShape && !fitPath && textPath.fontSizePt === undefined) {
      return unsupportedTextPathPlan(
        'VML textPath fitShape=false requires an authored font-size',
      );
    }
    if (textPath.fontSizePt === 0) {
      return Object.freeze({ status: 'planned', command: Object.freeze({ kind: 'noop' }) });
    }
    const fontSizePt = textPath.fontSizePt ?? SCALE_NEUTRAL_REFERENCE_FONT_SIZE_PT;
    const family = textPath.fontFamily ?? undefined;
    const shaped = text.shape({
      text: textPath.string,
      fontSizePt,
      fonts: {
        ascii: family,
        highAnsi: family,
        eastAsia: family,
        complexScript: family,
      },
      weight: textPath.bold ? 700 : 400,
      style: textPath.italic ? 'italic' : 'normal',
      measure: true,
    });
    if (textPath.trim === true && !shaped.inkBounds) {
      return unsupportedTextPathPlan(
        'VML textPath trim=true requires glyph ink bounds',
      );
    }
    // The retained source rectangle must choose one coherent metric domain:
    // trim removes the font's reserved box and uses tight actual ink, while
    // non-trim keeps the typographic advance/ascent/descent rectangle.
    const xMinPt = textPath.trim === true ? shaped.inkBounds?.xMinPt ?? 0 : 0;
    const xMaxPt = textPath.trim === true
      ? shaped.inkBounds?.xMaxPt ?? 0
      : shaped.advancePt;
    const ascentPt = textPath.trim === true ? shaped.inkBounds?.ascentPt ?? 0 : shaped.ascentPt;
    const descentPt = textPath.trim === true ? shaped.inkBounds?.descentPt ?? 0 : shaped.descentPt;
    const sourceBounds = {
      xPt: xMinPt,
      yPt: -ascentPt,
      widthPt: xMaxPt - xMinPt,
      heightPt: ascentPt + descentPt,
    };
    if (
      !Number.isFinite(shaped.advancePt)
      || Object.values(sourceBounds).some((value) => !Number.isFinite(value))
      || shaped.spans.some((span) => !Number.isFinite(span.advancePt))
    ) {
      throw new Error('Shape textPath acquisition produced non-finite metrics');
    }
    if (
      sourceBounds.widthPt <= 0
      || sourceBounds.heightPt <= 0
      || shaped.spans.length === 0
    ) {
      return unsupportedTextPathPlan('VML textPath produced empty glyph metrics');
    }
    const command = snapshotPlainData({
        kind: 'watermark-text' as const,
        rect: { ...bounds },
        text: textPath.string,
        ...(shape.fill?.fillType === 'image' ? { imageResourceKey: imageFillResourceKey,
          imageRotationDeg: shape.fill.vml?.alignShape === false ? 0 : shape.rotation ?? 0,
          ...(shape.fill.vml?.alignShape === false ? { imageFlipH: shape.flipH ?? false, imageFlipV: shape.flipV ?? false } : {}),
        } : {}),
        fill: shape.fill?.fillType === 'image' ? { fillType: 'solid' as const, color: 'FFFFFF' } : shape.fill ? { ...shape.fill, ...(shape.fill.fillType === 'gradient'
          ? { stops: shape.fill.stops.map((stop) => ({ ...stop })) }
          : {}) } : null,
        opacity: Math.max(0, Math.min(1, shape.fillOpacity ?? 1)),
        rotationDeg: shape.rotation ?? 0,
        flipH: shape.flipH ?? false,
        flipV: shape.flipV ?? false,
        ...(shape.vmlSkew ? { glyphTransform: acquireVmlSkewMatrix({
          rect: { x: bounds.xPt, y: bounds.yPt, w: bounds.widthPt, h: bounds.heightPt },
          transform: { rotationDeg: shape.rotation ?? 0, flipH: shape.flipH ?? false, flipV: shape.flipV ?? false },
        }, shape.vmlSkew) } : {}),
        fitShape,
        fontSizePt,
        sourceBounds,
        spans: shaped.spans.map((span) => ({
          text: span.text,
          advancePt: span.advancePt,
          fontRoute: span.fontRoute,
          fontWeight: span.font.weight,
          fontStyle: span.font.style,
        })),
      }, 'VML textPath command');
    let acquiredCommand: Exclude<ShapeDrawingCommand, {kind: 'noop'}> =
      textPath.fitPath || textPath.xScale || shape.subpaths.length
        ? snapshotPlainData(
            {
              ...command,
              pathSlices: planVmlPathText(shape, bounds, textPath, command, maximumGeometryWork),
            },
            'VML fitted path text',
          )
        : command;
    const pen = shapeStroke(shape);
    if (acquiredCommand.imageResourceKey || pen || shape.vmlSkew || shape.vmlGroupTransform || shape.vmlShadow) {
      const ink = shaped.inkBounds
        ? {
            xPt: shaped.inkBounds.xMinPt,
            yPt: -shaped.inkBounds.ascentPt,
            widthPt: shaped.inkBounds.xMaxPt - shaped.inkBounds.xMinPt,
            heightPt: shaped.inkBounds.ascentPt + shaped.inkBounds.descentPt,
          }
        : sourceBounds;
      const points: Array<{ x: number; y: number }> = [];
      const corners = [
        [ink.xPt, ink.yPt],
        [ink.xPt + ink.widthPt, ink.yPt],
        [ink.xPt + ink.widthPt, ink.yPt + ink.heightPt],
        [ink.xPt, ink.yPt + ink.heightPt],
      ];
      if (acquiredCommand.pathSlices)
        for (const slice of acquiredCommand.pathSlices) {
          for (const x of [-slice.sourceWidth / 2, slice.sourceWidth / 2])
            for (const y of [sourceBounds.yPt, sourceBounds.yPt + sourceBounds.heightPt])
              points.push({
                x: bounds.xPt + slice.xPt + slice.a * x + slice.c * y,
                y: bounds.yPt + slice.yPt + slice.b * x + slice.d * y,
              });
        }
      else {
        const sx = fitShape ? bounds.widthPt / sourceBounds.widthPt : 1,
          sy = fitShape ? bounds.heightPt / sourceBounds.heightPt : 1;
        for (const [x, y] of corners)
          points.push({
            x: bounds.xPt + (x! - sourceBounds.xPt) * sx,
            y: bounds.yPt + (y! - sourceBounds.yPt) * sy,
          });
      }
      if (command.glyphTransform) {
        const m = command.glyphTransform;
        for (const p of points) {
          const x = p.x, y = p.y;
          p.x = m.a * x + m.c * y + m.e; p.y = m.b * x + m.d * y + m.f;
        }
      }
      if (points.length) {
        const left = Math.min(...points.map((p) => p.x)),
          top = Math.min(...points.map((p) => p.y));
        acquiredCommand = snapshotPlainData(
          {
            ...acquiredCommand,
            maskBounds: {
              xPt: left - (pen?.width ?? 0),
              yPt: top - (pen?.width ?? 0),
              widthPt: Math.max(...points.map((p) => p.x)) - left + 2 * (pen?.width ?? 0),
              heightPt: Math.max(...points.map((p) => p.y)) - top + 2 * (pen?.width ?? 0),
            },
          },
          'VML image text ink',
        );
      }
    }
    const {
      imageResourceKey: _imageKey,
      imageRotationDeg: _imageRotation,
      imageFlipH: _imageFlipH, imageFlipV: _imageFlipV,
      ...shadowText
    } = acquiredCommand;
    const shadow = shape.vmlShadow;
    const before: Exclude<ShapeDrawingCommand, {kind: 'noop'}>[] = shadow && !pen && shadow.kind !== 'perspective' && shadow.opacity > 0
      ? (shadow.kind === 'single' ? [[shadow.offset, shadow.color]] as const
        : [[shadow.offset, shadow.color], [shadow.offset2, shadow.color2]] as const).map(([offset, color]) => snapshotPlainData({
          ...shadowText,
          rect: { ...bounds, xPt: bounds.xPt + offset.xPt + offset.xFraction * bounds.widthPt,
            yPt: bounds.yPt + offset.yPt + offset.yFraction * bounds.heightPt },
          ...(shadowText.glyphTransform ? { glyphTransform: {
            ...shadowText.glyphTransform,
            e: shadowText.glyphTransform.e + (1 - shadowText.glyphTransform.a) * (offset.xPt + offset.xFraction * bounds.widthPt)
              - shadowText.glyphTransform.c * (offset.yPt + offset.yFraction * bounds.heightPt),
            f: shadowText.glyphTransform.f + (1 - shadowText.glyphTransform.d) * (offset.yPt + offset.yFraction * bounds.heightPt)
              - shadowText.glyphTransform.b * (offset.xPt + offset.xFraction * bounds.widthPt),
          } } : {}),
          ...(shadowText.maskBounds ? { maskBounds: { ...shadowText.maskBounds,
            xPt: shadowText.maskBounds.xPt + offset.xPt + offset.xFraction * bounds.widthPt,
            yPt: shadowText.maskBounds.yPt + offset.yPt + offset.yFraction * bounds.heightPt,
          } } : {}),
          fill: shape.fill || !shadow.obscured ? { fillType: 'solid' as const, color } : null, opacity: shadow.opacity,
        }, 'VML text shadow')) : [];
    const diagnostics: Array<Readonly<Omit<LayoutDiagnostic, 'source'>>> = [];
    // Part 4 §§19.1.2.23/.21: first deform glyph geometry, THEN apply the
    // physical point-space pen. Native strokeText would deform the pen too.
    // Font-service contours keep arbitrary affine fits and the scale-neutral
    // reference font honest, and share ordinary retained stroke machinery.
    const affinePerspective = shadow?.kind === 'perspective' && shadow.matrix[4] === 0 && shadow.matrix[5] === 0;
    if (pen || (affinePerspective && shadow.opacity > 0)) {
      const scales = acquiredCommand.pathSlices?.map(s => Math.hypot(s.a, s.b, s.c, s.d))
        ?? [fitShape ? Math.max(bounds.widthPt / sourceBounds.widthPt, bounds.heightPt / sourceBounds.heightPt) : 1];
      if (!text.glyphContours || !shaped.inkBounds) {
        diagnostics.push({ code: 'UNSUPPORTED_FEATURE', severity: 'error',
          message: 'VML text decoration requires native glyph contours' });
      } else {
        // Bound every later affine amplification when choosing native contour
        // precision. Ancestor skew and shadow scale must not magnify a coarse
        // source contour beyond the font service's output-point tolerance.
        const groupScale = shape.vmlGroupTransform
          ? Math.hypot(...shape.vmlGroupTransform.slice(0, 4)) : 1;
        const shadowScale = affinePerspective ? Math.hypot(...shadow.matrix.slice(0, 4)) : 1;
        const contours = text.glyphContours({ spans: shaped.spans, fontSizePt,
          ...(shaped.nativeBaseDirection !== undefined ? { baseDirection: shaped.nativeBaseDirection } : {}),
          ink: shaped.inkBounds, maximumScale: Math.max(...scales) * Math.max(1, groupScale)
            * Math.max(1, shadowScale) * (command.glyphTransform
            ? Math.hypot(command.glyphTransform.a, command.glyphTransform.b, command.glyphTransform.c, command.glyphTransform.d) : 1), maximumWork: maximumGeometryWork ?? Infinity });
        if (!contours?.length) {
          diagnostics.push({ code: 'UNSUPPORTED_FEATURE', severity: 'error',
            message: 'Native glyph contours are unavailable at the bounded outline precision' });
        } else {
          const fitted = acquireVmlTextContours(contours, acquiredCommand, bounds, maximumGeometryWork);
          const contourShape = { ...shape, presetGeometry: null, subpaths: fitted.subpaths,
            subpathPaint: [], fill: null, vmlPicture: undefined, vmlEdgeStrokes: undefined,
            // The fitting map already includes the host-local glyph skew.
            vmlSkew: undefined, vmlConcentricGradient: false, vmlShadow: undefined,
            headEnd: null, tailEnd: null,
          };
          let remaining = Math.min(maximumGeometryWork ?? Infinity, 1_048_576) - fitted.workUnits;
          if (shadow && (affinePerspective || shadow.kind !== 'perspective')) {
            // Part 4 §19.1.2.18: shadow matrices act on acquired glyph ink,
            // including fill-only text. Zero p factors are ordinary affine
            // geometry; they require no reciprocal-unit compatibility rule.
            const frame = fitted.bounds;
            const shadowPlan: DrawingMLShapePaintPlan = {
              rect: {x: frame.xPt, y: frame.yPt, w: frame.widthPt, h: frame.heightPt},
              geometry: {kind: 'custom', subpaths: fitted.subpaths},
              fill: shape.fill ? {fillType: 'solid', color: shadow.color} : null,
              stroke: pen, transform: {rotationDeg: shape.rotation ?? 0,
                flipH: shape.flipH ?? false, flipV: shape.flipV ?? false},
            };
            const resolvedGeometry = resolveDrawingMLGeometry(shadowPlan, 1, remaining);
            remaining -= resolvedGeometry.workUnits;
            const shadows = planVmlShadows(shape, {...shadowPlan, resolvedGeometry}, remaining,
              {x: bounds.xPt, y: bounds.yPt, w: bounds.widthPt, h: bounds.heightPt});
            before.push(...shadows);
            remaining -= shadows.reduce((sum, c) => sum + c.plan.resolvedGeometry.workUnits, 0);
          }
          if (pen) {
            const outline = planShapeDrawingPens(contourShape, fitted.bounds, undefined, undefined, undefined, remaining, undefined, vmlStrokeResourceKey);
            if (outline.status !== 'planned') throw new Error('VML text outline acquisition failed');
            before.push(...(outline.before ?? []));
            before.push(acquiredCommand);
            if (outline.command.kind === 'noop') throw new Error('VML text outline produced no command');
            acquiredCommand = outline.command;
          }
        }
      }
    }
    if (shadow?.kind === 'perspective' && !affinePerspective) diagnostics.push({ code: 'UNSUPPORTED_FEATURE',
      severity: 'error', message: 'VML projective text shadow requires resolved reciprocal-factor semantics' });
    return Object.freeze({ status: 'planned', command: acquiredCommand, ...(before.length ? { before } : {}),
      ...(diagnostics.length ? { diagnostics } : {}),
    });
  }
  let plan: DrawingMLShapePaintPlan = {
    rect: { x: bounds.xPt, y: bounds.yPt, w: bounds.widthPt, h: bounds.heightPt },
    geometry: shape.presetGeometry
      ? {
          kind: 'preset',
          name: shape.presetGeometry,
          adjustments: [...(shape.adjValues ?? [])],
        }
      : {
          kind: 'custom',
          subpaths: shape.subpaths.map((subpath) => subpath.map((command) => ({ ...command }))),
          // ECMA-376 §20.1.9.15 per-path fill/stroke, kept only when it
          // lines up with the subpaths.
          ...(shape.subpathPaint?.length && shape.subpathPaint.length === shape.subpaths.length
            ? { paint: shape.subpathPaint.map((paint) => ({ ...paint })) }
            : {}),
        },
    fill: shape.fill && shape.fill.fillType !== 'image' ? { ...shape.fill, ...(shape.fill.fillType === 'gradient'
      ? { stops: shape.fill.stops.map((stop) => ({ ...stop })) }
      : {}) } : null,
    stroke: shapeStroke(shape),
    transform: {
      rotationDeg: shape.rotation ?? 0,
      flipH: shape.flipH ?? false,
      flipV: shape.flipV ?? false,
    },
  };
  if (shape.vmlSkew) plan = acquireVmlSkew(plan, shape.vmlSkew, maximumGeometryWork);
  const retainedPlan = { ...plan, resolvedGeometry: plan.resolvedGeometry ?? resolveDrawingMLGeometry(plan, 1, maximumGeometryWork) };
  const shadows = planVmlShadows(shape, retainedPlan, (maximumGeometryWork ?? Infinity) - retainedPlan.resolvedGeometry.workUnits);
  let remainingWork = (maximumGeometryWork ?? Infinity) - retainedPlan.resolvedGeometry.workUnits
    - shadows.reduce((sum, command) => sum + (command.plan.resolvedGeometry?.workUnits ?? 0), 0);
  const acquirePass = (pass: DrawingMLShapePaintPlan) => {
    const resolvedGeometry = resolveVmlGeometryPass(pass, remainingWork);
    remainingWork -= resolvedGeometry.workUnits;
    return { ...pass, resolvedGeometry };
  };
  if (shape.vmlPicture) {
    if (!vmlPictureResourceKey) throw new Error('VML picture requires a retained image resource key');
    // Part 4 §§19.1.2.10/.11: shadow, backing brush, picture, outline.
    // Each pass binds geometry to its own fill/stroke contract during
    // acquisition. Paint only replays the retained commands.
    const backingPlan = acquirePass({ ...plan, stroke: null });
    const before: NonNullable<Extract<ShapeDrawingPlanResult, { status: 'planned' }>['before']>[number][] = [...shadows];
    if (shape.fill?.fillType === 'image') {
      if (!imageFillResourceKey) throw new Error('VML backing brush requires a retained image resource key');
      before.push({ kind: 'drawingml-image-fill', plan: backingPlan,
        resourceKey: imageFillResourceKey,
        ...(shape.fill.vml?.alignShape === false ? {
          imageRotationDeg: -(shape.rotation ?? 0), imageFlipH: shape.flipH ?? false, imageFlipV: shape.flipV ?? false,
        } : {}),
      });
    } else if (shape.fill) {
      before.push({ kind: 'drawingml-shape', plan: backingPlan,
        ...(shape.fillOpacity == null ? {} : { opacity: Math.max(0, Math.min(1, shape.fillOpacity)) }),
      });
    }
    return Object.freeze({ status: 'planned',
      before: snapshotPlainData(before, 'VML picture backing commands'),
      command: snapshotPlainData({ kind: 'drawingml-image-fill' as const,
        plan: acquirePass({ ...plan, fill: null }), resourceKey: vmlPictureResourceKey,
      }, 'VML picture and outline command'),
    });
  }
  if (shape.fill?.fillType === 'image') {
    if (shape.fill.tile !== undefined) {
      return Object.freeze({
        status: 'unsupported',
        command: Object.freeze({ kind: 'noop' }),
        diagnostics: Object.freeze([Object.freeze({
          code: 'UNSUPPORTED_FEATURE',
          severity: 'error',
          message: 'Tiled DrawingML shape image fills are not rendered',
        })]),
      });
    }
    if (!imageFillResourceKey) {
      throw new Error('DrawingML shape image fill requires a retained image resource key');
    }
    return Object.freeze({
      status: 'planned',
      ...(shadows.length ? { before: snapshotPlainData(shadows, 'VML shadows') } : {}),
      command: snapshotPlainData({
        kind: 'drawingml-image-fill' as const,
        plan: retainedPlan,
        resourceKey: imageFillResourceKey,
        ...(shape.fill.vml?.alignShape === false ? {
          imageRotationDeg: -(shape.rotation ?? 0), imageFlipH: shape.flipH ?? false, imageFlipV: shape.flipV ?? false,
        } : {}),
        ...(shape.fill.fillRect === undefined ? {} : {
          fillRect: {
            l: shape.fill.fillRect.l ?? 0,
            t: shape.fill.fillRect.t ?? 0,
            r: shape.fill.fillRect.r ?? 0,
            b: shape.fill.fillRect.b ?? 0,
          },
        }),
      }, 'DrawingML shape image-fill command'),
    });
  }
  // Part 4 §19.1.2.5: primary fill opacity does not attenuate the
  // independent stroke or shadow. Gradients/image brushes retain their alpha
  // in the brush; solid fills retain it separately. Split only translucent
  // fill+stroke commands, charging both geometries to the same allowance.
  if (shape.fill && shape.fillOpacity != null && shape.fillOpacity !== 1) {
    const opacity = Math.max(0, Math.min(1, shape.fillOpacity));
    const fillCommand = { kind: 'drawingml-shape' as const,
      plan: plan.stroke ? acquirePass({ ...plan, stroke: null }) : retainedPlan, opacity };
    return Object.freeze({
      status: 'planned',
      ...(shadows.length || plan.stroke ? { before: snapshotPlainData(
        [...shadows, ...(plan.stroke ? [fillCommand] : [])], 'VML fill and shadows',
      ) } : {}),
      command: snapshotPlainData(plan.stroke
        ? { kind: 'drawingml-shape' as const, plan: acquirePass({ ...plan, fill: null }) }
        : fillCommand, 'VML shape command'),
    });
  }
  return Object.freeze({
    status: 'planned',
    ...(shadows.length ? { before: snapshotPlainData(shadows, 'VML shadows') } : {}),
    command: snapshotPlainData(
      { kind: 'drawingml-shape', plan: retainedPlan } as const,
      'DrawingML shape command',
    ),
  });
}

/** Acquire pen passes independently: texture alpha must not multiply body fill
 * alpha, and an inset pen is a double-width stroke clipped to the silhouette
 * (Part 4 §§19.1.2.14/.21). All geometry is retained before paint. */
function planShapeDrawingPens(
  shape: DeepReadonly<ShapeRun & VmlShapeAcquisitionFacts>,
  bounds: LayoutRect,
  text?: TextLayoutService,
  textPath?: Readonly<VmlTextPathAcquisitionInput>,
  imageFillResourceKey?: string,
  maximumGeometryWork?: number,
  vmlPictureResourceKey?: string,
  vmlStrokeResourceKey?: string,
): ShapeDrawingPlanResult {
  const result = planShapeDrawingBase(
    shape,
    bounds,
    text,
    textPath,
    imageFillResourceKey,
    maximumGeometryWork,
    vmlPictureResourceKey,
    vmlStrokeResourceKey,
  );
  // Text acquisition has already routed its contour pen through this same
  // adapter. Reprocessing those passes would double inset width and acquire
  // image resources a second time.
  if (result.status === 'planned' && (result.command.kind === 'watermark-text'
    || result.before?.some(command => command.kind === 'watermark-text'))) return result;
  if (
    result.status !== 'planned' ||
    (!shape.vmlStrokeFill && !shape.vmlInsetPen && !shape.vmlConcentricGradient)
  )
    return result;
  const commands: ShapeDrawingCommand[] = [];
  let remaining = maximumGeometryWork ?? Infinity;
  for (const command of [...(result.before ?? []), result.command]) {
    // Shadow commands already own their independent tint/opacity and must
    // never acquire the source pen's image resource.
    const shadowPass = command.kind === 'drawingml-shape' && command.shadowPass;
    if (command.kind !== 'drawingml-shape' && command.kind !== 'drawingml-image-fill') {
      commands.push(command);
      continue;
    }
    remaining -= command.plan.resolvedGeometry.workUnits;
    const concentricFill =
      !shadowPass && shape.vmlConcentricGradient
        ? acquireVmlConcentricFill(command.plan)
        : undefined;
    if (shadowPass || (!command.plan.stroke && !concentricFill)) {
      commands.push(command);
      continue;
    }
    const insetStroke = shape.vmlInsetPen === true;
    const acquire = (plan: DrawingMLShapePaintPlan) => {
      const resolvedGeometry = resolveVmlGeometryPass(plan, remaining);
      remaining -= resolvedGeometry.workUnits;
      return { ...plan, resolvedGeometry };
    };
    if (concentricFill) {
      const work = concentricFill.rings.length * command.plan.resolvedGeometry.workUnits;
      if (work > 1_048_576)
        throw new RangeError('VML concentric fill geometry work budget exceeded');
      remaining -= work;
      if (remaining <= 0) throw new RangeError('VML concentric fill geometry work budget exceeded');
      commands.push({
        kind: 'drawingml-shape',
        plan: acquire({ ...command.plan, stroke: null }),
        concentricFill,
      });
      if (!command.plan.stroke) continue;
    }
    if (!command.plan.stroke) continue;
    const penPlan = acquire({
      ...command.plan,
      fill: null,
      stroke: {
        ...command.plan.stroke,
        ...(shape.vmlStrokeFill ? { color: 'FFFFFF' } : {}),
        width: command.plan.stroke.width * (insetStroke ? 2 : 1),
      },
    });
    if (!concentricFill)
      commands.push({ ...command, plan: acquire({ ...command.plan, stroke: null }) });
    if (shape.vmlStrokeFill) {
      if (!vmlStrokeResourceKey) throw new Error('VML textured pen requires a retained resource');
      const ink = drawingMLGeometryBounds({
        ...penPlan,
        transform: { rotationDeg: 0, flipH: false, flipV: false },
      });
      commands.push({
        kind: 'drawingml-image-fill',
        plan: penPlan,
        resourceKey: vmlStrokeResourceKey,
        strokeMask: true,
        insetStroke,
        ...(ink ? { maskBounds: { xPt: ink.x, yPt: ink.y, widthPt: ink.w, heightPt: ink.h } } : {}),
        ...(shape.vmlStrokeFill.vml?.alignShape === false
          ? {
              imageRotationDeg: -(shape.rotation ?? 0),
              imageFlipH: shape.flipH ?? false,
              imageFlipV: shape.flipV ?? false,
            }
          : {}),
      });
    } else
      commands.push({
        kind: 'drawingml-shape',
        plan: penPlan,
        insetStroke,
        ...(command.kind === 'drawingml-shape' && command.opacity !== undefined
          ? { opacity: command.opacity }
          : {}),
      });
  }
  return snapshotPlainData(
    {
      ...result,
      before: commands
        .slice(0, -1)
        .filter(
          (command): command is Exclude<ShapeDrawingCommand, { kind: 'noop' }> =>
            command.kind !== 'noop',
        ),
      command: commands[commands.length - 1]!,
    },
    'VML pen passes',
  );
}

/** Acquire independent border pens after the host's own fill/pen passes. */
export function planShapeDrawing(
  shape: DeepReadonly<ShapeRun & VmlShapeAcquisitionFacts>, bounds: LayoutRect,
  text?: TextLayoutService, textPath?: Readonly<VmlTextPathAcquisitionInput>,
  imageFillResourceKey?: string, maximumGeometryWork?: number,
  vmlPictureResourceKey?: string, vmlStrokeResourceKey?: string,
  vmlEdgeResourcePrefix?: string,
): ShapeDrawingPlanResult {
  let result = planShapeDrawingPens(shape, bounds, text, textPath, imageFillResourceKey,
    maximumGeometryWork, vmlPictureResourceKey, vmlStrokeResourceKey);
  if (result.status === 'planned' && shape.vmlEdgeStrokes?.length) {
    const edgeCommands: Exclude<ShapeDrawingCommand, {kind: 'noop'}>[] = [];
    let remaining = maximumGeometryWork ?? Infinity;
    for (const command of [...(result.before ?? []), result.command])
      if (command.kind === 'drawingml-shape' || command.kind === 'drawingml-image-fill')
        remaining -= command.plan.resolvedGeometry.workUnits;
    for (const edge of shape.vmlEdgeStrokes) {
      // Part 4 §19.2.2.6 explicitly permits ignoring the interior pen when
      // multi-column VML text boxes are not supported. No interior line exists
      // for this adapter's single-column text box; do not invent one.
      if (edge.edge === 'column') continue;
      const ends = { left: [0, 0, 0, 1], top: [0, 0, 1, 0],
        right: [1, 0, 1, 1], bottom: [0, 1, 1, 1] }[edge.edge];
      const penShape = {
        ...shape, presetGeometry: null, vmlEdgeStrokes: undefined, vmlGroupTransform: undefined,
        vmlShadow: undefined, vmlPicture: undefined, vmlConcentricGradient: false,
        fill: null, stroke: edge.color, strokeWidth: edge.widthPt,
        strokeDash: edge.dash, strokeCap: edge.cap, strokeJoin: edge.join,
        strokeMiterLimit: edge.miterLimit, strokeCompound: edge.compound,
        headEnd: null, tailEnd: null, vmlStrokeFill: edge.image ?? undefined,
        vmlStrokePattern: edge.pattern ?? undefined, vmlInsetPen: false,
        strokeFill: null, strokeCustomDash: undefined, strokeAlignment: undefined,
        subpaths: [[{ cmd: 'moveTo' as const, x: ends[0]!, y: ends[1]! },
          { cmd: 'lineTo' as const, x: ends[2]!, y: ends[3]! }]], subpathPaint: [],
      };
      if (edge.image && !vmlEdgeResourcePrefix) throw new Error('VML edge image pen requires a retained resource');
      const acquired = planShapeDrawing(penShape, bounds, undefined, undefined, undefined,
        remaining, undefined, edge.image ? `${vmlEdgeResourcePrefix}${edge.edge}` : undefined);
      if (acquired.status !== 'planned') throw new Error('VML edge pen acquisition failed');
      for (const command of [...(acquired.before ?? []), acquired.command]) {
        if (command.kind === 'noop') continue;
        edgeCommands.push(command);
        if (command.kind === 'drawingml-shape' || command.kind === 'drawingml-image-fill')
          remaining -= command.plan.resolvedGeometry.workUnits;
      }
    }
    if (edgeCommands.length) result = snapshotPlainData({ ...result,
      before: [...(result.before ?? []), ...(result.command.kind === 'noop' ? [] : [result.command]), ...edgeCommands.slice(0, -1)],
      command: edgeCommands.at(-1) ?? result.command,
    }, 'VML independent edge pens');
  }
  return shape.vmlGroupTransform
    ? acquireVmlGroupVectors(result, shape.vmlGroupTransform, bounds, maximumGeometryWork)
    : result;
}

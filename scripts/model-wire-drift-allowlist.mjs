// Known differences between the Rust model wire format and the hand-written
// TS model types, checked by scripts/check-model-wire-drift.mjs.
//
// Keys are `<RustType>.<field>: <category>` exactly as the checker prints them
// (`<RustType><tag=value>` names an inline enum variant). `*` matches any run
// of characters and is used only for one family with one shared reason.
// Every entry must still match a real difference: once the drift is fixed the
// checker reports the entry as stale and it must be deleted.
//
// Categories:
//   missing-in-ts          Rust emits the field; the TS type does not declare it.
//   ts-only                The TS type declares a field Rust never emits.
//   rust-optional          Rust may omit the field; the TS type requires it.
//   rust-null              Rust emits `null`; the TS type does not allow null.
//   literal-missing-in-ts  Rust emits a string value the TS literal union lacks.
//   variant-missing-in-ts  A tagged-union variant Rust emits is absent in TS.
//   variant-never-emitted  A TS tagged-union variant Rust never emits.
//   type-mismatch          Incompatible shapes.
//
// Entries marked DRIFT are genuine inaccuracies reported for follow-up (each
// changes a public TS type, so it is not fixed here); the rest are intentional.

export const MODEL_WIRE_DRIFT_ALLOWLIST = {
  docx: {
    // ---- Intentional: internal parser→layout channels ------------------------
    // `__`-prefixed fields are private acquisition facts that the layout layer
    // reads through internal wire types (packages/docx/src/layout/*-input.ts,
    // page-object-wire.ts, parser-model.ts, core math); they are deliberately
    // not part of the public model types.
    '*.__*: missing-in-ts': 'internal channel',

    // ---- Intentional: facts read through internal wire types -----------------
    // The public model type omits these; parser-model.ts / layout types read
    // them (typography acquisition, VML drawing plan, textbox content).
    'Document.diagnostics: missing-in-ts': 'read by parser-model.ts as parser diagnostics',
    'Document.fontFamilyCharsets: missing-in-ts': 'read by the layout source store / font services',
    'DocParagraph<type=paragraph>.paragraphMarkFontFacts: missing-in-ts': 'read by parser-model.ts',
    'NumberingInfo.fontFacts: missing-in-ts': 'read by parser-model.ts',
    'TextRun<type=text>.fontFamilyHighAnsi: missing-in-ts': 'typography acquisition input',
    'TextRun<type=text>.fontHint: missing-in-ts': 'typography acquisition input',
    'TextRun<type=text>.fontSlots: missing-in-ts': 'typography acquisition input',
    'TextRun<type=text>.langEastAsia: missing-in-ts': 'typography acquisition input',
    'FieldRun<type=field>.boldCs: missing-in-ts': 'typography acquisition input',
    'FieldRun<type=field>.cs: missing-in-ts': 'typography acquisition input',
    'FieldRun<type=field>.fontFamilyCs: missing-in-ts': 'typography acquisition input',
    'FieldRun<type=field>.fontFamilyEastAsia: missing-in-ts': 'typography acquisition input',
    'FieldRun<type=field>.fontFamilyHighAnsi: missing-in-ts': 'typography acquisition input',
    'FieldRun<type=field>.fontHint: missing-in-ts': 'typography acquisition input',
    'FieldRun<type=field>.fontSizeCs: missing-in-ts': 'typography acquisition input',
    'FieldRun<type=field>.fontSlots: missing-in-ts': 'typography acquisition input',
    'FieldRun<type=field>.italicCs: missing-in-ts': 'typography acquisition input',
    'FieldRun<type=field>.langBidi: missing-in-ts': 'typography acquisition input',
    'FieldRun<type=field>.langEastAsia: missing-in-ts': 'typography acquisition input',
    'FieldRun<type=field>.rtl: missing-in-ts': 'typography acquisition input',
    'ShapeRun<type=shape>.textBoxContent: missing-in-ts': 'textbox acquisition input',
    'ShapeRun<type=shape>.vmlConcentricGradient: missing-in-ts': 'VML drawing plan (layout/types.ts)',
    'ShapeRun<type=shape>.vmlEdgeStrokes: missing-in-ts': 'VML drawing plan (layout/types.ts)',
    'ShapeRun<type=shape>.vmlGroupTransform: missing-in-ts': 'VML drawing plan (layout/types.ts)',
    'ShapeRun<type=shape>.vmlInsetPen: missing-in-ts': 'VML drawing plan (layout/types.ts)',
    'ShapeRun<type=shape>.vmlSkew: missing-in-ts': 'VML drawing plan (layout/types.ts)',
    'ShapeRun<type=shape>.vmlStrokeFill: missing-in-ts': 'VML drawing plan (layout/types.ts)',
    'ShapeRun<type=shape>.vmlStrokePattern: missing-in-ts': 'VML drawing plan (layout/types.ts)',
    'ShapeRun<type=shape>.vmlTextBox: missing-in-ts': 'VML drawing plan (layout/types.ts)',
    'ShapeRun<type=shape>.vmlTextRects: missing-in-ts': 'VML drawing plan (layout/types.ts)',
    'TextPath.fitPath: missing-in-ts': 'VML text path acquisition (parser-model.ts)',
    'TextPath.fitShape: missing-in-ts': 'VML text path acquisition (parser-model.ts)',
    'TextPath.fontSizePt: missing-in-ts': 'VML text path acquisition (parser-model.ts)',
    'TextPath.on: missing-in-ts': 'VML text path acquisition (parser-model.ts)',
    'TextPath.textPathOk: missing-in-ts': 'VML text path acquisition (parser-model.ts)',
    'TextPath.trim: missing-in-ts': 'VML text path acquisition (parser-model.ts)',
    'TextPath.xScale: missing-in-ts': 'VML text path acquisition (parser-model.ts)',
    // `ShapeRun.vmlPicture` is a full Rust ImageRun; TS deliberately types it as
    // `Pick<ImageRun, ...>` with only the picture-fill fields it paints.
    'ImageRun.*: missing-in-ts': 'vmlPicture is a Pick<ImageRun> projection',
    'ImageRun.colorReplaceFrom: rust-null': 'vmlPicture: see ImageRun<type=image>.colorReplaceFrom',
    // Parser-private run kind recovered through a sidecar, not a public DocRun.
    'DocParagraph<type=paragraph>.runs<type=unavailableDrawing>: variant-missing-in-ts':
      'parser-private UnavailableDrawing (layout/text.ts UnavailableDrawingAcquisitionRun)',
    // The parser strips Word's rendered-page-break hints before serialization.
    'DocRun<type=break>.breakType: literal-missing-in-ts': '"renderedPage" runs are stripped by the parser',

    // ---- Intentional: TS type is a deliberate widening -----------------------
    // `DocRun = DocRunContent & { revision?: RunRevision }`: TS lets every run
    // carry a revision; Rust only emits it on text runs.
    '*.revision: ts-only': 'DocRun intersection widens revision to every run kind',
    // One TS MathScript interface covers both sub and sup nodes.
    'MathNode<kind=sub>.sup: ts-only': 'shared TS MathScript interface',
    'MathNode<kind=sup>.sub: ts-only': 'shared TS MathScript interface',
    // Shared TS ImageFill type also used for shape picture fills.
    'EffectFill<fillType=image>.alpha: ts-only': 'shared TS BlipEffectImageFill type',
    'EffectFill<fillType=image>.duotone: ts-only': 'shared TS BlipEffectImageFill type',
    // Chart fields derived on the TS side (packages/core/src/chart) after parse.
    'ChartLabelBox.effectFallbackIndex: ts-only': 'derived by core label-box style merge',
    'ChartLabelBox.effectFallbackStyle: ts-only': 'derived by core label-box style merge',
    'ChartLabelBox.effectStyleIndex: ts-only': 'derived by core label-box style merge',
    'ChartModel.fontSubstitutions: ts-only': 'supplied by the TS chart host',
    'ChartModel.linkedChartStyleRoles: ts-only': 'derived by core effective-style',
    'ChartModel.varyingPointChartStyleRoles: ts-only': 'derived by core effective-style',
    'ChartModel.varyingPointChartStyleRolesByGroup: ts-only': 'derived by core effective-style',
    // Rust ChartStyleFill has no `none` variant; TS fill fields use the shared
    // core `Fill` union (which includes NoFill) and so accept it.
    '*<fillType=none>: variant-never-emitted': 'shared core Fill union includes NoFill',
    // DrawingML line paints (`a:ln`, §20.1.2.2.24) cannot hold a blipFill, so
    // the TS line-paint types exclude the image variant of ChartStyleFill.
    'ChartDecorationLineStyle.fill<fillType=image>: variant-missing-in-ts': 'a:ln has no blipFill',
    'ChartExElementStyle.linePaints<fillType=image>: variant-missing-in-ts': 'a:ln has no blipFill',
    'ChartLabelBox.borderFill<fillType=image>: variant-missing-in-ts': 'a:ln has no blipFill',
    'ChartModel.chartBorderLineFill<fillType=image>: variant-missing-in-ts': 'a:ln has no blipFill',
    'ChartModel.legendLineFill<fillType=image>: variant-missing-in-ts': 'a:ln has no blipFill',
    'ChartModel.plotAreaLineFill<fillType=image>: variant-missing-in-ts': 'a:ln has no blipFill',

    // ---- DRIFT: genuine inaccuracies, reported for follow-up -----------------
    // serde camelCases `view_3d_present` to `view3dPresent`; the public TS type
    // declares `view3DPresent`, so the declared field is never populated (no TS
    // code reads it yet).
    'ChartThreeD.view3dPresent: missing-in-ts': 'DRIFT: TS name view3DPresent',
    'ChartThreeD.view3DPresent: ts-only': 'DRIFT: wire name is view3dPresent',
    // Emitted for every heading paragraph but declared and read by no TS code.
    'DocParagraph<type=paragraph>.outlineLevel: missing-in-ts': 'DRIFT: unread wire field',
    // Wire omits absent edges/fill (`undefined`); TS declares them required
    // `X | null`. Current readers use nullish/truthy checks.
    'ParagraphBorders.between: rust-optional': 'DRIFT: TS should be optional',
    'ParagraphBorders.bottom: rust-optional': 'DRIFT: TS should be optional',
    'ParagraphBorders.left: rust-optional': 'DRIFT: TS should be optional',
    'ParagraphBorders.right: rust-optional': 'DRIFT: TS should be optional',
    'ParagraphBorders.top: rust-optional': 'DRIFT: TS should be optional',
    'ShapeRun<type=shape>.fill: rust-optional': 'DRIFT: TS should be optional',
    // Not skipped when None, so the wire carries `null`; TS says `?: string`.
    // Readers use truthiness.
    'ImageRun<type=image>.colorReplaceFrom: rust-null': 'DRIFT: TS lacks null',
    // Area fills that DrawingML allows as blipFill (dTable / bandFmt / c14
    // inverted fill spPr): Rust can carry an image fill the TS type excludes.
    'ChartDataTable.fill<fillType=image>: variant-missing-in-ts': 'DRIFT: TS excludes image fill',
    'ChartSeries.invertedFill<fillType=image>: variant-missing-in-ts': 'DRIFT: TS excludes image fill',
    'ChartSurfaceBandFormat.fill<fillType=image>: variant-missing-in-ts': 'DRIFT: TS excludes image fill',
  },
};

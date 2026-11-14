// Known differences between the Rust model wire format and the hand-written
// TS model types, checked by scripts/check-model-wire-drift.mjs.
//
// Keys are exactly what `node scripts/check-model-wire-drift.mjs --list`
// prints: `<owner>.<field>: <category>[ <payload>]`. The owner is the Rust
// type (`<Type><tag=value>` for a tagged-union variant, `Rust(ts TsName)` when
// the TS type has a different name); the payload pins concrete values, so an
// allowlisted difference cannot silently grow. Matching is exact. Every entry
// must still match: once a difference is fixed the checker reports the entry
// as stale and it must be deleted.
//
// Categories:
//   missing-in-ts          Rust emits the field; the TS type does not declare it.
//   ts-only                The TS type declares a field Rust never emits.
//   rust-optional          Rust may omit the field; the TS type requires it.
//   rust-null              Rust emits `null`; the TS type does not allow null.
//   literal-missing-in-ts  Rust emits literal values the TS type does not accept.
//   literal-never-emitted  TS accepts literal values Rust never emits.
//   ts-narrower-literal    Rust emits any string/number/boolean; TS narrows it.
//   variant-missing-in-ts  A union alternative Rust emits is absent in TS.
//   variant-never-emitted  A TS union alternative Rust never emits.
//   rust-untyped/ts-untyped  One side is unconstrained (any), so unchecked.
//   type-mismatch          Incompatible shapes.
//
// Entries marked DRIFT are genuine inaccuracies reported for follow-up; the
// rest are intentional.

export const MODEL_WIRE_DRIFT_ALLOWLIST = {
  docx: {
    // ---- Intentional: internal parser→layout channels -----------------------
    // `__`-prefixed fields are private acquisition facts that the layout layer
    // reads through internal wire types (packages/docx/src/layout/*-input.ts,
    // page-object-wire.ts, parser-model.ts, core math); they are deliberately
    // not part of the public model types. Each one is listed, so a new channel
    // needs a reviewed entry.
    'AnchorHostMetrics<type=anchorHost>.__anchorOccurrenceId: missing-in-ts': 'internal channel',
    'AnchorHostMetrics<type=anchorHost>.__listLabel: missing-in-ts': 'internal channel',
    'BodyElement<type=sectionBreak>.__sectionPlacement: missing-in-ts': 'internal channel',
    'ChartRun<type=chart>.__anchorAcquisition: missing-in-ts': 'internal channel',
    'DocParagraph<type=paragraph>.__complexFieldBoundaries: missing-in-ts': 'internal channel',
    'DocParagraph<type=paragraph>.__indentsFromFloats: missing-in-ts': 'internal channel',
    'DocParagraph<type=paragraph>.__justifiedCompression: missing-in-ts': 'internal channel',
    'DocParagraph<type=paragraph>.__lineExtraAbove: missing-in-ts': 'internal channel',
    'DocParagraph<type=paragraph>.__noHangingIndentTab: missing-in-ts': 'internal channel',
    'DocParagraph<type=paragraph>.__paragraphTypographyAcquisition: missing-in-ts': 'internal channel',
    'DocParagraph<type=paragraph>.__runRevisions: missing-in-ts': 'internal channel',
    'DocTable<type=table>.__tableLayout: missing-in-ts': 'internal channel',
    'DocTableCell.__tableCellLayout: missing-in-ts': 'internal channel',
    'DocTableRow.__cutsInPlace: missing-in-ts': 'internal channel',
    'DocTableRow.__tableRowLayout: missing-in-ts': 'internal channel',
    'Document(ts DocxDocumentModel).__documentTypographySettings: missing-in-ts': 'internal channel',
    'Document(ts DocxDocumentModel).__noteLayoutSettings: missing-in-ts': 'internal channel',
    'Document(ts DocxDocumentModel).__pageLayoutSettings: missing-in-ts': 'internal channel',
    'Document(ts DocxDocumentModel).__pageObjects: missing-in-ts': 'internal channel',
    'DocumentSettings(ts DocSettings).__footnoteColumnBands: missing-in-ts': 'internal channel',
    'DocumentSettings(ts DocSettings).__plainFontLineBoxes: missing-in-ts': 'internal channel',
    'FieldRun<type=field>.__typographyAcquisition: missing-in-ts': 'internal channel',
    'FramePr.__anchorLock: missing-in-ts': 'internal channel',
    'ImageRun(ts Pick).__anchorAcquisition: missing-in-ts': 'internal channel',
    'ImageRun(ts Pick).__inlineObjectFontFacts: missing-in-ts': 'internal channel',
    'ImageRun(ts Pick).__inlineObjectOrigin: missing-in-ts': 'internal channel',
    'ImageRun(ts Pick).__listLabel: missing-in-ts': 'internal channel',
    'ImageRun<type=image>.__anchorAcquisition: missing-in-ts': 'internal channel',
    'ImageRun<type=image>.__inlineObjectFontFacts: missing-in-ts': 'internal channel',
    'ImageRun<type=image>.__inlineObjectOrigin: missing-in-ts': 'internal channel',
    'ImageRun<type=image>.__listLabel: missing-in-ts': 'internal channel',
    'MathNode<kind=run>.__script: missing-in-ts': 'internal channel',
    'SectionProps.__sectionPlacement: missing-in-ts': 'internal channel',
    'ShapeRun<type=shape>.__anchorAcquisition: missing-in-ts': 'internal channel',
    'ShapeRun<type=shape>.__inlineObjectFontFacts: missing-in-ts': 'internal channel',
    'ShapeRun<type=shape>.__inlineObjectOrigin: missing-in-ts': 'internal channel',
    'TextRun<type=text>.__listLabel: missing-in-ts': 'internal channel',
    'TextRun<type=text>.__noBreakAfter: missing-in-ts': 'internal channel',
    'TextRun<type=text>.__noBreakBefore: missing-in-ts': 'internal channel',
    'TextRun<type=text>.__noBreakHyphenOffsets: missing-in-ts': 'internal channel',
    'TextRun<type=text>.__optionalHyphen: missing-in-ts': 'internal channel',
    'TextRun<type=text>.__spaceAdvancePt: missing-in-ts': 'internal channel',
    'TextRun<type=text>.__typographyAcquisition: missing-in-ts': 'internal channel',

    // ---- Intentional: facts read through internal (non-public) wire types ---
    // The public model types omit these; parser-model.ts and the layout types read
    // them (typography acquisition, VML drawing plan, VML text path, textbox).
    'DocParagraph<type=paragraph>.paragraphMarkFontFacts: missing-in-ts': 'read by parser-model.ts',
    'Document(ts DocxDocumentModel).diagnostics: missing-in-ts': 'parser diagnostics read by parser-model.ts',
    'Document(ts DocxDocumentModel).fontFamilyCharsets: missing-in-ts': 'read by the layout source store and font services',
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
    'NumberingInfo.fontFacts: missing-in-ts': 'read by parser-model.ts',
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
    'TextRun<type=text>.fontFamilyHighAnsi: missing-in-ts': 'typography acquisition input',
    'TextRun<type=text>.fontHint: missing-in-ts': 'typography acquisition input',
    'TextRun<type=text>.fontSlots: missing-in-ts': 'typography acquisition input',
    'TextRun<type=text>.langEastAsia: missing-in-ts': 'typography acquisition input',

    // ---- Intentional: ShapeRun.vmlPicture projection ------------------------
    // `ShapeRun.vmlPicture` is a full Rust ImageRun; TS deliberately types it as
    // `Pick<ImageRun, ...>` with only the picture-fill fields it paints.
    'ImageRun(ts Pick).allowOverlap: missing-in-ts': 'vmlPicture Pick<ImageRun> projection',
    'ImageRun(ts Pick).alpha: missing-in-ts': 'vmlPicture Pick<ImageRun> projection',
    'ImageRun(ts Pick).anchor: missing-in-ts': 'vmlPicture Pick<ImageRun> projection',
    'ImageRun(ts Pick).anchorXAlign: missing-in-ts': 'vmlPicture Pick<ImageRun> projection',
    'ImageRun(ts Pick).anchorXFromMargin: missing-in-ts': 'vmlPicture Pick<ImageRun> projection',
    'ImageRun(ts Pick).anchorXPt: missing-in-ts': 'vmlPicture Pick<ImageRun> projection',
    'ImageRun(ts Pick).anchorXRelativeFrom: missing-in-ts': 'vmlPicture Pick<ImageRun> projection',
    'ImageRun(ts Pick).anchorYAlign: missing-in-ts': 'vmlPicture Pick<ImageRun> projection',
    'ImageRun(ts Pick).anchorYFromPara: missing-in-ts': 'vmlPicture Pick<ImageRun> projection',
    'ImageRun(ts Pick).anchorYPt: missing-in-ts': 'vmlPicture Pick<ImageRun> projection',
    'ImageRun(ts Pick).anchorYRelativeFrom: missing-in-ts': 'vmlPicture Pick<ImageRun> projection',
    'ImageRun(ts Pick).distBottom: missing-in-ts': 'vmlPicture Pick<ImageRun> projection',
    'ImageRun(ts Pick).distLeft: missing-in-ts': 'vmlPicture Pick<ImageRun> projection',
    'ImageRun(ts Pick).distRight: missing-in-ts': 'vmlPicture Pick<ImageRun> projection',
    'ImageRun(ts Pick).distTop: missing-in-ts': 'vmlPicture Pick<ImageRun> projection',
    'ImageRun(ts Pick).flipH: missing-in-ts': 'vmlPicture Pick<ImageRun> projection',
    'ImageRun(ts Pick).flipV: missing-in-ts': 'vmlPicture Pick<ImageRun> projection',
    'ImageRun(ts Pick).heightPt: missing-in-ts': 'vmlPicture Pick<ImageRun> projection',
    'ImageRun(ts Pick).inlineEffectExtent: missing-in-ts': 'vmlPicture Pick<ImageRun> projection',
    'ImageRun(ts Pick).rotation: missing-in-ts': 'vmlPicture Pick<ImageRun> projection',
    'ImageRun(ts Pick).svgImagePath: missing-in-ts': 'vmlPicture Pick<ImageRun> projection',
    'ImageRun(ts Pick).widthPt: missing-in-ts': 'vmlPicture Pick<ImageRun> projection',
    'ImageRun(ts Pick).wrapMode: missing-in-ts': 'vmlPicture Pick<ImageRun> projection',
    'ImageRun(ts Pick).wrapSide: missing-in-ts': 'vmlPicture Pick<ImageRun> projection',

    // ---- Intentional: parser-private or stripped values ---------------------
    // UnavailableDrawing runs are parser-private (recovered through a sidecar and
    // typed as layout/text.ts UnavailableDrawingAcquisitionRun); Word rendered-page
    // break hints are stripped by the parser before serialization.
    'DocParagraph<type=paragraph>.runs<type=unavailableDrawing>: variant-missing-in-ts': 'parser-private UnavailableDrawing',
    'DocRun<type=break>.breakType: literal-missing-in-ts ["renderedPage"]': 'renderedPage runs are stripped by the parser',

    // ---- Intentional: DocRun revision widening ------------------------------
    // `DocRun = DocRunContent & { revision?: RunRevision }` lets every run kind
    // carry a revision; Rust emits it on text runs only.
    'AnchorHostMetrics<type=anchorHost>.revision: ts-only': 'DocRun intersection',
    'ChartRun<type=chart>.revision: ts-only': 'DocRun intersection',
    'DocRun<type=break>.revision: ts-only': 'DocRun intersection',
    'DocRun<type=math>.revision: ts-only': 'DocRun intersection',
    'DocRun<type=ptab>.revision: ts-only': 'DocRun intersection',
    'FieldRun<type=field>.revision: ts-only': 'DocRun intersection',
    'ImageRun<type=image>.revision: ts-only': 'DocRun intersection',
    'ShapeRun<type=shape>.revision: ts-only': 'DocRun intersection',

    // ---- Intentional: shared TS interfaces ----------------------------------
    // One TS interface serves several Rust shapes (MathScript for sub and sup;
    // BlipEffectImageFill for every image effect fill).
    'EffectFill<fillType=image>.alpha: ts-only': 'shared TS interface',
    'EffectFill<fillType=image>.duotone: ts-only': 'shared TS interface',
    'MathNode<kind=sub>.sup: ts-only': 'shared TS interface',
    'MathNode<kind=sup>.sub: ts-only': 'shared TS interface',

    // ---- Intentional: chart fields derived on the TS side -------------------
    // Filled by packages/core/src/chart (label-box style merge, effective-style)
    // or by the chart host after parsing.
    'ChartLabelBox.effectFallbackIndex: ts-only': 'derived by core chart code',
    'ChartLabelBox.effectFallbackStyle: ts-only': 'derived by core chart code',
    'ChartLabelBox.effectStyleIndex: ts-only': 'derived by core chart code',
    'ChartModel.fontSubstitutions: ts-only': 'derived by core chart code',
    'ChartModel.linkedChartStyleRoles: ts-only': 'derived by core chart code',
    'ChartModel.varyingPointChartStyleRoles: ts-only': 'derived by core chart code',
    'ChartModel.varyingPointChartStyleRolesByGroup: ts-only': 'derived by core chart code',

    // ---- Intentional: shared core Fill union --------------------------------
    // Rust ChartStyleFill has no `none` variant; these TS fields use the shared
    // core `Fill` union, which includes NoFill.
    'ChartDataPointOverride.markerFillPaint<fillType=none>: variant-never-emitted': 'core Fill includes NoFill',
    'ChartExElementStyle.fillPaints<fillType=none>: variant-never-emitted': 'core Fill includes NoFill',
    'ChartLabelBox.fillPaint<fillType=none>: variant-never-emitted': 'core Fill includes NoFill',
    'ChartModel.chartFill<fillType=none>: variant-never-emitted': 'core Fill includes NoFill',
    'ChartModel.legendFill<fillType=none>: variant-never-emitted': 'core Fill includes NoFill',
    'ChartModel.plotAreaFill<fillType=none>: variant-never-emitted': 'core Fill includes NoFill',
    'ChartSeries.markerFillPaint<fillType=none>: variant-never-emitted': 'core Fill includes NoFill',
    'ChartStockBarPaint.fill<fillType=none>: variant-never-emitted': 'core Fill includes NoFill',

    // ---- Intentional: DrawingML line paints ---------------------------------
    // A DrawingML line (`a:ln`, ECMA-376 §20.1.2.2.24) cannot hold a blipFill, so
    // the TS line-paint types exclude the image variant of ChartStyleFill.
    'ChartDecorationLineStyle.fill<fillType=image>: variant-missing-in-ts': 'a:ln has no blipFill',
    'ChartExElementStyle.linePaints<fillType=image>: variant-missing-in-ts': 'a:ln has no blipFill',
    'ChartLabelBox.borderFill<fillType=image>: variant-missing-in-ts': 'a:ln has no blipFill',
    'ChartModel.chartBorderLineFill<fillType=image>: variant-missing-in-ts': 'a:ln has no blipFill',
    'ChartModel.legendLineFill<fillType=image>: variant-missing-in-ts': 'a:ln has no blipFill',
    'ChartModel.plotAreaLineFill<fillType=image>: variant-missing-in-ts': 'a:ln has no blipFill',

    // ---- Intentional: values the parser restricts ---------------------------
    // `PathPaint.stroke` is serialized only when false (skip_serializing_if on
    // true), which TS types as `stroke?: false`.
    'PathPaint.stroke: ts-narrower-literal [false]': 'serialized only when false',

    // ---- Intentional: vocabularies narrowed in TS ---------------------------
    // The Rust field is a `String` (or number) the parser fills from a closed
    // OOXML vocabulary; TS names that vocabulary. The payload pins the TS
    // literals, so changing them must update this entry. Making the Rust field
    // an enum would turn these into checked comparisons.
    'BlipEffect<type=blend>.blend: ts-narrower-literal ["darken","lighten","mult","over","screen"]': 'closed vocabulary',
    'BlipEffect<type=fillOverlay>.blend: ts-narrower-literal ["darken","lighten","mult","over","screen"]': 'closed vocabulary',
    'BlipReflection.algn: ts-narrower-literal ["b","bl","br","ctr","l","r","t","tl","tr"]': 'closed vocabulary',
    'BodyElement<type=pageBreak>.parity: ts-narrower-literal ["even","odd"]': 'closed vocabulary',
    'ChartexHistogramBinning.intervalClosed: ts-narrower-literal ["l","r"]': 'closed vocabulary',
    'ChartexRegionMap.regionLabelLayout: ts-narrower-literal ["bestFitOnly","none","showAll"]': 'closed vocabulary',
    'ChartexRegionMapColors.stopCount: ts-narrower-literal [2,3]': 'closed vocabulary',
    'ChartModel.bubbleSizeRepresents: ts-narrower-literal ["area","w"]': 'closed vocabulary',
    'ChartModel.catAxisTitleVerticalMode: ts-narrower-literal ["eaVert","horz","mongolianVert","vert","vert270","wordArtVert","wordArtVertRtl"]': 'closed vocabulary',
    'ChartModel.legendPos: ts-narrower-literal ["b","l","r","t","tr"]': 'closed vocabulary',
    'ChartModel.valAxisTitleVerticalMode: ts-narrower-literal ["eaVert","horz","mongolianVert","vert","vert270","wordArtVert","wordArtVertRtl"]': 'closed vocabulary',
    'ChartOfPie.splitType: ts-narrower-literal ["auto","cust","percent","pos","val"]': 'closed vocabulary',
    'ChartOfPie.type: ts-narrower-literal ["bar","pie"]': 'closed vocabulary',
    'ChartPatternFill(ts PatternFill).fillType: ts-narrower-literal ["pattern"]': 'closed vocabulary',
    'ChartPlotGroup.bubbleSizeRepresents: ts-narrower-literal ["area","w"]': 'closed vocabulary',
    'ChartPlotGroup.categoryAxis: ts-narrower-literal ["none","primary","secondary","unresolved"]': 'closed vocabulary',
    'ChartPlotGroup.kind: ts-narrower-literal ["area","area3D","bar","bar3D","bubble","doughnut","line","line3D","ofPie","pie","pie3D","radar","scatter","stock","surface","surface3D"]': 'closed vocabulary',
    'ChartPlotGroup.seriesAxis: ts-narrower-literal ["none","primary","secondary","unresolved"]': 'closed vocabulary',
    'ChartPlotGroup.valueAxis: ts-narrower-literal ["none","primary","secondary","unresolved"]': 'closed vocabulary',
    'ChartThreeDSeriesAxis.titleVerticalMode: ts-narrower-literal ["eaVert","horz","mongolianVert","vert","vert270","wordArtVert","wordArtVertRtl"]': 'closed vocabulary',
    'DocRun<type=ptab>.alignment: ts-narrower-literal ["center","left","right"]': 'closed vocabulary',
    'DocRun<type=ptab>.leader: ts-narrower-literal ["dot","hyphen","middleDot","none","underscore"]': 'closed vocabulary',
    'DocRun<type=ptab>.relativeTo: ts-narrower-literal ["indent","margin"]': 'closed vocabulary',
    'DocTableCell.vAlign: ts-narrower-literal ["bottom","center","top"]': 'closed vocabulary',
    'EffectContainer(ts BlipEffectContainer).kind: ts-narrower-literal ["sib","tree"]': 'closed vocabulary',
    'EmbeddedFont(ts EmbeddedFontRef).style: ts-narrower-literal ["bold","boldItalic","italic","regular"]': 'closed vocabulary',
    'FieldRun<type=field>.emphasisMark: ts-narrower-literal ["circle","comma","dot","underDot"]': 'closed vocabulary',
    'FieldRun<type=field>.vertAlign: ts-narrower-literal ["sub","super"]': 'closed vocabulary',
    'LineSpacing.rule: ts-narrower-literal ["atLeast","auto","exact"]': 'closed vocabulary',
    'MathNode<kind=array>.align: ts-narrower-literal ["center","eq","left"]': 'closed vocabulary',
    'MathNode<kind=bar>.pos: ts-narrower-literal ["bot","top"]': 'closed vocabulary',
    'MathNode<kind=groupChr>.pos: ts-narrower-literal ["bot","top"]': 'closed vocabulary',
    'MathNode<kind=run>.style: ts-narrower-literal ["bold","boldItalic","italic","roman"]': 'closed vocabulary',
    'PathPaint.fill: ts-narrower-literal ["darken","darkenLess","lighten","lightenLess","none"]': 'closed vocabulary',
    'PathPaint.fillRule: ts-narrower-literal ["evenodd","nonzero"]': 'closed vocabulary',
    'SecondaryValueAxis.labelAlignment: ts-narrower-literal ["ctr","l","r"]': 'closed vocabulary',
    'SecondaryValueAxis.titleVerticalMode: ts-narrower-literal ["eaVert","horz","mongolianVert","vert","vert270","wordArtVert","wordArtVertRtl"]': 'closed vocabulary',
    'Shadow.algn: ts-narrower-literal ["b","bl","br","ctr","l","r","t","tl","tr"]': 'closed vocabulary',
    'ShapeRun<type=shape>.strokeAlignment: ts-narrower-literal ["ctr","in"]': 'closed vocabulary',
    'ShapeRun<type=shape>.strokeCap: ts-narrower-literal ["butt","round","square"]': 'closed vocabulary',
    'ShapeRun<type=shape>.strokeJoin: ts-narrower-literal ["bevel","miter","round"]': 'closed vocabulary',
    'TabStop.alignment: ts-narrower-literal ["bar","center","clear","decimal","end","left","num","right","start"]': 'closed vocabulary',
    'TabStop.leader: ts-narrower-literal ["dot","heavy","hyphen","middleDot","none","underscore"]': 'closed vocabulary',
    'TextRun<type=text>.emphasisMark: ts-narrower-literal ["circle","comma","dot","underDot"]': 'closed vocabulary',
    'TextRun<type=text>.vertAlign: ts-narrower-literal ["sub","super"]': 'closed vocabulary',
    'VmlShadow.kind: ts-narrower-literal ["double","emboss","perspective","single"]': 'closed vocabulary',

    // ---- DRIFT: genuine inaccuracies, reported for follow-up ----------------
    // Each fix changes a public TS type (and API baseline), so none is fixed in
    // the PR that introduced this check.
    'ChartDataTable.fill<fillType=image>: variant-missing-in-ts': 'DRIFT: TS excludes an image fill Rust can carry',
    'ChartSeries.invertedFill<fillType=image>: variant-missing-in-ts': 'DRIFT: TS excludes an image fill Rust can carry',
    'ChartSurfaceBandFormat.fill<fillType=image>: variant-missing-in-ts': 'DRIFT: TS excludes an image fill Rust can carry',
    'ChartThreeD.view3DPresent: ts-only': 'DRIFT: the wire name is view3dPresent',
    'ChartThreeD.view3dPresent: missing-in-ts': 'DRIFT: TS declares view3DPresent',
    'DocParagraph<type=paragraph>.outlineLevel: missing-in-ts': 'DRIFT: emitted but read by no TS code',
    'ImageRun(ts Pick).colorReplaceFrom: rust-null': 'DRIFT: TS lacks null (readers use truthiness)',
    'ImageRun<type=image>.colorReplaceFrom: rust-null': 'DRIFT: TS lacks null (readers use truthiness)',
    'ParagraphBorders.between: rust-optional': 'DRIFT: TS should mark it optional',
    'ParagraphBorders.bottom: rust-optional': 'DRIFT: TS should mark it optional',
    'ParagraphBorders.left: rust-optional': 'DRIFT: TS should mark it optional',
    'ParagraphBorders.right: rust-optional': 'DRIFT: TS should mark it optional',
    'ParagraphBorders.top: rust-optional': 'DRIFT: TS should mark it optional',
    'ShapeRun<type=shape>.fill: rust-optional': 'DRIFT: TS should mark it optional',
    'TextRun<type=text>.fitTextId: variant-never-emitted (TS number)': 'DRIFT: the wire value is always a string',
  },
};

import type { ReferenceFontMetricProfile } from './reference-font-metrics.js';

/**
 * Published OFL font metadata, separate from the generated Office/macOS bundle
 * catalogs. These are metadata references, not installed-font detection: Canvas
 * may still select a different face for the authored family. Add a face only
 * from a pinned public font source after verifying the exact TTF hash, hhea,
 * unitsPerEm, OS/2 usWin metrics, and code-page bits. Keep the original font
 * bytes out of the runtime bundle.
 *
 * BIZ UDMincho: googlefonts/morisawa-biz-ud-mincho at
 * c30a6221b1f3d09afae9137ffe73c7cbec649947, fonts/ttf/.
 * Regular SHA-256: 468ee6d9b149ca144809e03841bf18740ecf014e055a00da6ecaf1aaf4165af2
 * Bold SHA-256: 1f077f8f84c1e09d5c4acdd6828048180c2f733ae5ae13271f48cf01bee4ae83
 * Both have OS/2 ulCodePageRange1=0x20020009 (Far East bit 17 set),
 * unitsPerEm=2048, hhea=(1802, -246, 0), and OS/2 usWin=(1802, 246).
 * The checked data covers only regular and bold normal styles; italic tuples deliberately remain absent.
 * The pinned Bold SHA-256 and usWin tuple were reverified on 2026-10-16;
 * installed macOS BIZ_UDMincho-regular.ttf confirms the Regular tuple.
 *
 * BIZ UDGothic: googlefonts/morisawa-biz-ud-gothic at
 * 18934af56b9c003ca58c54bffbf226848cb11032, fonts/ttf/.
 * Regular SHA-256: 7d2b48d84ef4e65c9f85bcf65fb1fb41c92b134be641fab6a7141d597c9a92b0
 * Bold SHA-256: c686d1b05b36d98473d98f67e74e96048a5ad6695de143111b8328b78900006e
 * Both have OS/2 ulCodePageRange1=0x20020009 (Far East bit 17 set),
 * unitsPerEm=2048, hhea=(1802, -246, 0), OS/2 usWin=(1802, 246),
 * and xAvgCharWidth=1718.
 * The same hhea/usWin and code-page metrics were confirmed in the installed
 * macOS BIZ_UDGothic.ttc regular/bold faces. These static metadata references
 * do not bundle font bytes or establish which face Canvas selected.
 *
 * Name identity (font metadata, not an Office layout rule): the pinned regular
 * and bold TTFs of both families record the Japanese (language 0x0411) name
 * ID 1 "BIZ UDゴシック" / "BIZ UD明朝" beside the English "BIZ UDGothic" /
 * "BIZ UDMincho", in the same files and therefore with the same hhea values.
 * Office documents can record either spelling as the requested family (for
 * example w:rFonts and the ECMA-376 §17.8.3.10 w:font@w:name), so both are
 * aliases of one profile. The localized name is a family
 * name shared by regular and bold; it never identifies a styled face by itself.
 * The installed macOS asset faces carry only the English names, so a local()
 * probe still needs the English/PostScript aliases listed first. The
 * proportional BIZ UDP* families are distinct fonts and are not aliased here.
 *
 * Noto Sans 2.015: notofonts/notofonts.github.io at
 * 28b15b4b43b7bed62b5cf6e6b0b5ff5846270535, fonts/NotoSans/hinted/ttf/.
 * Regular SHA-256: 478c558ea716033cd60c03438f628dfa75694dcf6b5f6d505a2f05fd2b4f3823
 * Bold SHA-256: 1df075a380fc7cb898acf64c1f7b3b4dd780de3caa860178bf929de35817a913
 * Italic SHA-256: 467e3f89eeca4108bb8710a2b9e0cf2281ac56d5b0609211a83776d0505eecb5
 * BoldItalic SHA-256: 1b602a9d6353be42c91df097a4857b69fa2696f26703d7a33b54a15d87c2622c
 * All four have OS/2 ulCodePageRange1=0x0000019f (no Far East bit),
 * unitsPerEm=1000, hhea=(1069, -293, 0), OS/2 usWin=(1124, 395), and
 * fsSelection USE_TYPO_METRICS with typo=(1069, -293, 0); the files' only
 * names are "Noto Sans" and the PostScript names below. The installed macOS
 * copies have the same hashes. Decoration and PANOSE are not recorded here.
 *
 * Noto Sans JP 2.004: notofonts/noto-cjk at
 * 165c01b46ea533872e002e0785ff17e44f6d97d8, Sans/SubsetOTF/JP/.
 * Regular SHA-256: dff723ba59d57d136764a04b9b2d03205544f7cd785a711442d6d2d085ac5073
 * Bold SHA-256: 1b0edfb500b73a4fa8a4fcaae1bbbd403994e08e73e3e0da37e70d3853f42c5f
 * Both have OS/2 ulCodePageRange1=0x60060107 (Far East bits 17 and 18 set),
 * unitsPerEm=1000, hhea=(1160, -288, 0) and OS/2 usWin=(1160, 288), without
 * USE_TYPO_METRICS; the only names are "Noto Sans JP" and the PostScript
 * names below. The installed macOS copies have the same hashes. Decoration,
 * PANOSE and GSUB vert coverage are not recorded here.
 *
 * Both Noto families were admitted with Word evidence only (DOCX line boxes
 * and grid cells), so `evidencedFormats` keeps them out of PowerPoint and
 * Excel font selection, requests and line boxes.
 */
export const OPEN_FONT_REFERENCE_PROFILES: readonly ReferenceFontMetricProfile[] = [
  {
    source: 'published-open-font', family: 'BIZ UDGothic',
    aliases: ['BIZ UDGothic', 'BIZUDGothic-Regular', 'BIZ UDゴシック'],
    weight: 400, style: 'normal', unitsPerEm: 2048,
    hhea: [1802, -246, 0], win: [1802, 246], farEastCodePage: true, xAvgCharWidth: 1718,
  },
  {
    source: 'published-open-font', family: 'BIZ UDGothic',
    aliases: ['BIZ UDGothic', 'BIZUDGothic-Bold', 'BIZ UDゴシック'],
    weight: 700, style: 'normal', unitsPerEm: 2048,
    hhea: [1802, -246, 0], win: [1802, 246], farEastCodePage: true, xAvgCharWidth: 1718,
  },
  {
    source: 'published-open-font', family: 'BIZ UDMincho',
    aliases: ['BIZ UDMincho', 'BIZUDMincho-Regular', 'BIZ UD明朝'],
    weight: 400, style: 'normal', unitsPerEm: 2048,
    hhea: [1802, -246, 0], win: [1802, 246], farEastCodePage: true, xAvgCharWidth: 1959,
  },
  {
    source: 'published-open-font', family: 'BIZ UDMincho',
    aliases: ['BIZ UDMincho', 'BIZUDMincho-Bold', 'BIZ UD明朝'],
    weight: 700, style: 'normal', unitsPerEm: 2048,
    hhea: [1802, -246, 0], win: [1802, 246], farEastCodePage: true, xAvgCharWidth: 1963,
  },
  ...([
    ['NotoSans-Regular', 400, 'normal', 577],
    ['NotoSans-Bold', 700, 'normal', 618],
    ['NotoSans-Italic', 400, 'italic', 550],
    ['NotoSans-BoldItalic', 700, 'italic', 592],
  ] as const).map(([postscript, weight, style, xAvgCharWidth]): ReferenceFontMetricProfile => ({
    source: 'published-open-font', family: 'Noto Sans', aliases: ['Noto Sans', postscript],
    weight, style, unitsPerEm: 1000, xAvgCharWidth,
    hhea: [1069, -293, 0], win: [1124, 395], typoMetrics: [1069, -293, 0], farEastCodePage: false,
    evidencedFormats: ['docx'],
  })),
  ...([
    ['NotoSansJP-Regular', 400, 982],
    ['NotoSansJP-Bold', 700, 983],
  ] as const).map(([postscript, weight, xAvgCharWidth]): ReferenceFontMetricProfile => ({
    source: 'published-open-font', family: 'Noto Sans JP', aliases: ['Noto Sans JP', postscript],
    weight, style: 'normal', unitsPerEm: 1000, xAvgCharWidth,
    hhea: [1160, -288, 0], win: [1160, 288], farEastCodePage: true,
    evidencedFormats: ['docx'],
  })),
];

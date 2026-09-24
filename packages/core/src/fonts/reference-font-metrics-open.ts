import type { ReferenceFontMetricProfile } from './reference-font-metrics.js';

/**
 * Published OFL font metadata, separate from the generated Office/macOS bundle
 * catalogs. These are metadata references, not installed-font detection: Canvas
 * may still select a different face for the authored family. Add a face only
 * from a pinned public font source after verifying the exact TTF hash, hhea,
 * unitsPerEm, and OS/2 code-page bits. Keep the original font bytes out of the
 * runtime bundle.
 *
 * BIZ UDMincho: googlefonts/morisawa-biz-ud-mincho at
 * c30a6221b1f3d09afae9137ffe73c7cbec649947, fonts/ttf/.
 * Regular SHA-256: 468ee6d9b149ca144809e03841bf18740ecf014e055a00da6ecaf1aaf4165af2
 * Bold SHA-256: 1f077f8f84c1e09d5c4acdd6828048180c2f733ae5ae13271f48cf01bee4ae83
 * Both have OS/2 ulCodePageRange1=0x20020009 (Far East bit 17 set),
 * unitsPerEm=2048, and hhea=(1802, -246, 0). The checked data covers only
 * regular and bold normal styles; italic tuples deliberately remain absent.
 *
 * BIZ UDGothic: googlefonts/morisawa-biz-ud-gothic at
 * 18934af56b9c003ca58c54bffbf226848cb11032, fonts/ttf/.
 * Regular SHA-256: 7d2b48d84ef4e65c9f85bcf65fb1fb41c92b134be641fab6a7141d597c9a92b0
 * Bold SHA-256: c686d1b05b36d98473d98f67e74e96048a5ad6695de143111b8328b78900006e
 * Both have OS/2 ulCodePageRange1=0x20020009 (Far East bit 17 set),
 * unitsPerEm=2048, hhea=(1802, -246, 0), and xAvgCharWidth=1718.
 * The same vertical and code-page metrics were confirmed in the installed
 * macOS BIZ_UDGothic.ttc regular/bold faces. These static metadata references
 * do not bundle font bytes or establish which face Canvas selected.
 */
export const OPEN_FONT_REFERENCE_PROFILES: readonly ReferenceFontMetricProfile[] = [
  {
    source: 'published-open-font', family: 'BIZ UDGothic',
    aliases: ['BIZ UDGothic', 'BIZUDGothic-Regular'],
    weight: 400, style: 'normal', unitsPerEm: 2048,
    hhea: [1802, -246, 0], farEastCodePage: true, xAvgCharWidth: 1718,
  },
  {
    source: 'published-open-font', family: 'BIZ UDGothic',
    aliases: ['BIZ UDGothic', 'BIZUDGothic-Bold'],
    weight: 700, style: 'normal', unitsPerEm: 2048,
    hhea: [1802, -246, 0], farEastCodePage: true, xAvgCharWidth: 1718,
  },
  {
    source: 'published-open-font', family: 'BIZ UDMincho',
    aliases: ['BIZ UDMincho', 'BIZUDMincho-Regular'],
    weight: 400, style: 'normal', unitsPerEm: 2048,
    hhea: [1802, -246, 0], farEastCodePage: true, xAvgCharWidth: 1959,
  },
  {
    source: 'published-open-font', family: 'BIZ UDMincho',
    aliases: ['BIZ UDMincho', 'BIZUDMincho-Bold'],
    weight: 700, style: 'normal', unitsPerEm: 2048,
    hhea: [1802, -246, 0], farEastCodePage: true, xAvgCharWidth: 1963,
  },
];

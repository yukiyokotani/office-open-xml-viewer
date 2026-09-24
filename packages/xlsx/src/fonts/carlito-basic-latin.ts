/*!
 * Carlito Basic Latin advances. Copyright 2013 The Carlito Project Authors.
 * Derived from Carlito-Bold.ttf (SHA-256 bb5d20f79b82599ec72983597437373a80f2d2085fa91fc144fd74e876a594db),
 * Google Fonts upstream commit 3a810cab78ebd6e2e4eed42af9e8453c4f9b850a.
 * SIL Open Font License 1.1: see ../CARLITO-OFL.txt.
 */
// Open-font, unshaped hmtx advances for U+0020–U+007E, indexed by code point - 0x20.
// Carlito is published as metric-compatible with Calibri. This is an XLSX
// compatibility reference for a missing authored face, not a shaped-text font.
const ADVANCES = [
  463, 667, 898, 1020, 1038, 1493, 1443, 478, 638, 638, 1020, 1020,
  528, 627, 547, 880, 1038, 1038, 1038, 1038, 1038, 1038, 1038, 1038,
  1038, 1038, 565, 565, 1020, 1020, 1020, 949, 1840, 1241, 1148, 1084,
  1291, 999, 940, 1305, 1292, 546, 678, 1120, 866, 1790, 1349, 1385,
  1090, 1405, 1153, 968, 1014, 1337, 1211, 1856, 1128, 1064, 979, 665,
  880, 665, 1020, 1020, 615, 1011, 1099, 857, 1099, 1031, 648, 971,
  1099, 503, 523, 983, 503, 1666, 1099, 1101, 1099, 1099, 728, 817,
  710, 1099, 969, 1526, 941, 970, 814, 704, 973, 704, 1020,
] as const;
const UNITS_PER_EM = 2048;

export function calibriCompatibleBasicLatinWidth(text: string, sizePx: number): number | undefined {
  if (!Number.isFinite(sizePx) || sizePx <= 0) return undefined;
  let units = 0;
  for (let i = 0; i < text.length; i++) {
    const index = text.charCodeAt(i) - 0x20;
    if (index < 0 || index >= ADVANCES.length) return undefined;
    units += ADVANCES[index];
  }
  return units / UNITS_PER_EM * sizePx;
}

export interface CalibriCompatibleWrapRoute {
  readonly family: string | null | undefined;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly checkedOfficeTuples: ReadonlySet<string>;
  readonly hasExactRoute: boolean;
  readonly hasDeclaredFace: boolean;
  readonly googleSubstitutes: boolean;
}

/** Library compatibility policy: an authored Calibri Bold tuple whose bounded
 * local() preflight completed without a route or timeout may use Carlito's
 * matching Basic Latin advances for XLSX plain-cell word wrapping. This is a
 * missing-resource policy, not proof that a system face is not installed.
 * Exact/local/application/opt-in web faces retain Canvas authority. hmtx does
 * not model GPOS kerning, GSUB ligatures, or non-Latin scripts; keep the gate
 * narrow until independently checked Office cases support expansion. This
 * boundary matches two Office-produced fixed-height wrapped-heading cases;
 * mixed-face runs, CJK text, and other OOXML formats were not established. */
export function shouldUseCalibriCompatibleWrap(route: CalibriCompatibleWrapRoute, text: string): boolean {
  return route.family?.trim().toLocaleLowerCase('en-US') === 'calibri'
    && route.bold && !route.italic && route.checkedOfficeTuples.has('calibri:700:normal')
    && !route.hasExactRoute && !route.hasDeclaredFace && !route.googleSubstitutes
    && /^[\x20-\x7e]+$/.test(text);
}

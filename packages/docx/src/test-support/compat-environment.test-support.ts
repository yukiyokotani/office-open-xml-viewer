import {
  buildDocxCompatProfile,
  type DocxCompatProfile,
  type DocxCompatSettingsInput,
} from '../layout/compat-profile.js';

const SETTINGS_KEYS = [
  'compatibilityMode', 'footnoteLayoutLikeWW8', 'adjustLineHeightInTable', 'useFeLayout',
  'balanceSingleByteDoubleByteWidth', 'lineWrapLikeWord6', 'enableOpenTypeFeatures',
  'ulTrailSpace', 'doNotExpandShiftReturn', 'doNotSuppressIndentation', 'ww11IndentRules',
  'doNotUseIndentAsNumberingTabStop', 'noTabHangInd',
] as const satisfies readonly (keyof DocxCompatSettingsInput)[];

type SettingsKey = typeof SETTINGS_KEYS[number];

/**
 * Test environments describe a document by its `w:settings` values. Move those
 * keys into the document compatibility profile that production measurement
 * environments carry (`compat`), keeping every other key. An existing
 * `compat` is replaced only when settings keys are present.
 */
export function withCompat<T extends object>(
  environment: T & Partial<Record<SettingsKey, unknown>>,
): Omit<T, SettingsKey> & { compat?: DocxCompatProfile } {
  const settings: Record<string, unknown> = {};
  const rest: Record<string, unknown> = {};
  let present = false;
  for (const [key, value] of Object.entries(environment)) {
    if ((SETTINGS_KEYS as readonly string[]).includes(key)) {
      present = true;
      settings[key] = value;
    } else {
      rest[key] = value;
    }
  }
  if (!present) return rest as Omit<T, SettingsKey>;
  return { ...rest, compat: buildDocxCompatProfile(settings as DocxCompatSettingsInput) } as
    Omit<T, SettingsKey> & { compat: DocxCompatProfile };
}

/** Profile of a document with these settings. */
export const compatProfile = (settings: DocxCompatSettingsInput = {}): DocxCompatProfile =>
  buildDocxCompatProfile(settings);

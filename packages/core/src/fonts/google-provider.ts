import { GOOGLE_FONT_SUBSTITUTES } from './google-fonts.js';
import {
  FontProvider,
  type FontAsset,
  type FontResolveOptions,
} from './provider.js';
import { loadGoogleFontRules, type FontPreloadEntry } from './google-css.js';
import { SCRIPT_GOOGLE_FONTS } from './scripts.js';

const GOOGLE_FONTS: Readonly<Record<string, FontPreloadEntry>> = {
  ...GOOGLE_FONT_SUBSTITUTES,
  ...SCRIPT_GOOGLE_FONTS,
};

function sourceUrl(source: string): string | undefined {
  return source.match(/url\(\s*(['"]?)(.*?)\1\s*\)/i)?.[2];
}

/** Built-in provider behind the backwards-compatible `useGoogleFonts` option. */
export class GoogleFontsProvider extends FontProvider {
  /** @internal */
  registrationFamily(family: string, isolatedFamily: string): string {
    const entry = GOOGLE_FONTS[family.trim().toLocaleLowerCase('en-US')];
    return entry?.loadFamily ?? (entry ? family.trim() : isolatedFamily);
  }

  /** @internal */
  registrationSource(
    family: string,
    registeredFamily: string,
  ): 'google' | 'substitute' {
    return family.trim().toLocaleLowerCase('en-US') === registeredFamily.toLocaleLowerCase('en-US')
      ? 'google'
      : 'substitute';
  }

  /** @internal */
  registrationKey(
    _family: string,
    registeredFamily: string,
    asset: FontAsset,
    isolatedKey: string,
  ): string {
    if (!('url' in asset.source)) return isolatedKey;
    const descriptors = asset.descriptors ?? {};
    return [
      'google-provider',
      registeredFamily.toLocaleLowerCase('en-US'),
      String(asset.source.url),
      descriptors.style ?? '',
      descriptors.weight ?? '',
      descriptors.stretch ?? '',
      descriptors.unicodeRange ?? '',
    ].join('|');
  }

  async resolve(
    families: readonly string[],
    { signal }: Readonly<FontResolveOptions>,
  ): Promise<readonly FontAsset[]> {
    const requests = new Map<string, { family: string; entry: FontPreloadEntry }>();
    for (const value of families) {
      const family = value.trim();
      const key = family.toLocaleLowerCase('en-US');
      const entry = GOOGLE_FONTS[key];
      if (family && entry) requests.set(key, { family, entry });
    }
    const rules = new Map<string, Awaited<ReturnType<typeof loadGoogleFontRules>>>();
    await Promise.all([...new Set([...requests.values()].map(({ entry }) => entry.url))]
      .map(async (url) => rules.set(url, await loadGoogleFontRules(url))));
    if (signal.aborted) throw signal.reason;

    const assets: FontAsset[] = [];
    for (const { family, entry } of requests.values()) {
      const target = (entry.loadFamily ?? family).toLocaleLowerCase('en-US');
      for (const rule of rules.get(entry.url) ?? []) {
        if (rule.family.toLocaleLowerCase('en-US') !== target) continue;
        const url = sourceUrl(rule.src);
        if (!url) continue;
        assets.push({ family, source: { url }, descriptors: rule.descriptors });
      }
    }
    return assets;
  }
}

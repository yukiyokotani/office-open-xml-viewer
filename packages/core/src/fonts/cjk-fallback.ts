import type { CjkLang } from './scripts.js';

/** Region used only when the document does not identify a CJK fallback. */
export type CjkFallback = 'auto' | CjkLang;

/** BCP 47 language → font region. These defaults are application policy,
 * not an OOXML font selection rule. Explicit scripts win over regions. */
export function cjkLangFromLanguage(language: string | null | undefined): CjkLang | null {
  if (!language?.trim()) return null;
  let locale: Intl.Locale;
  try { locale = new Intl.Locale(language.trim()); } catch { return null; }
  const { language: lang, script, region } = locale;
  if (lang === 'ja') return !script || ['Jpan', 'Hani', 'Hira', 'Kana'].includes(script) ? 'jp' : null;
  if (lang === 'ko') return !script || ['Kore', 'Hang', 'Hani'].includes(script) ? 'kr' : null;
  if (lang !== 'zh') return null;
  if (script === 'Hans') return 'sc';
  if (script === 'Hant') return 'tc';
  if (script) return null;
  if (region === 'HK' || region === 'MO') return 'hk';
  if (region === 'TW') return 'tc';
  return 'sc';
}

/** Snapshot on the loading realm before any asynchronous work. Workers receive
 * the concrete result; rendering never reads ambient language preferences. */
export function resolveCjkFallback(option: CjkFallback = 'auto'): CjkLang {
  if (option !== 'auto') {
    if (['sc', 'tc', 'hk', 'jp', 'kr'].includes(option)) return option;
    throw new TypeError('cjkFallback must be auto, sc, tc, hk, jp, or kr');
  }
  const html = typeof document === 'undefined'
    ? null : cjkLangFromLanguage(document.documentElement?.lang);
  if (html) return html;
  if (typeof navigator !== 'undefined') {
    for (const language of navigator.languages ?? []) {
      const region = cjkLangFromLanguage(language);
      if (region) return region;
    }
    const region = cjkLangFromLanguage(navigator.language);
    if (region) return region;
  }
  return 'jp';
}

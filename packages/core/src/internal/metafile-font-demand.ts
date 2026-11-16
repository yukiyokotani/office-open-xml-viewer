// Internal metafile text demand: imported only by the formats' lazily loaded
// font-source modules, so the record walk and the scalar tables stay out of
// the default (no font source) graph.
import { FontDemandScalarCollector, type FontDemandScalars } from '../fonts/scalar-demand.js';
import { metafilePaintTexts } from '../fonts/metafile-paint-texts.js';
import type { FontSessionDemand, LoadedFontSources } from '../fonts/source-session.js';

export { metafilePaintTexts } from '../fonts/metafile-paint-texts.js';

/** The Unicode scalars the metafile `bytes` paints or measures
 * (`metafilePaintTexts`, collected literally: the players apply no symbol-font
 * conversion beyond their own), or `'all'` when core's scalar rules cannot
 * prove them (unknown scalars, lone surrogates, the scalar ceiling). */
export function metafileDemandScalars(bytes: Uint8Array): FontDemandScalars {
  const collector = new FontDemandScalarCollector({ budget: Number.POSITIVE_INFINITY });
  for (const text of metafilePaintTexts(bytes)) if (!collector.addSymbolText(text, [], true)) break;
  return collector.result();
}

/** What {@link metafileFontDemand} needs of a font-source session. */
export interface MetafileFontDemandSession {
  ensure(demand: FontSessionDemand): Promise<LoadedFontSources>;
}

const demands = new WeakMap<MetafileFontDemandSession, (bytes: Uint8Array) => Promise<LoadedFontSources>>();

/**
 * The demand a document's image owner runs before a metafile picture plays
 * (core `bindMetafileFontDemand`): the scalars its text records paint
 * ({@link metafileDemandScalars}) are loaded from the families `session`
 * already serves, and the resulting snapshot is what the picture paints with.
 * Metafile text never adds a family: the players paint served faces only for
 * families the document's own demand loaded. One function per session.
 */
export function metafileFontDemand(session: MetafileFontDemandSession): (bytes: Uint8Array) => Promise<LoadedFontSources> {
  let demand = demands.get(session);
  if (!demand) {
    demand = (bytes) => session.ensure({ families: [], fallbackFamilies: [], scalars: metafileDemandScalars(bytes) });
    demands.set(session, demand);
  }
  return demand;
}

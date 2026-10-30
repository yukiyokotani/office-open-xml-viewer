/**
 * Installed authored families are never demanded from a font source.
 *
 * Library policy (font-source design §3.2, §8): the authored family names the
 * face to use, so an installed authored face is never displaced by a source
 * face. A source face registers under its real CSS family, and a document
 * font shadows a system font of the same name, so a source serving an
 * installed family (an application face named "Arial") would otherwise
 * replace it. Every authored family is probed (`local()`, batched past the
 * per-call cap), before any source face of the delta registers; only a family
 * whose probe completed and found no installed face is demanded and cached. A
 * family whose probe did not complete (deadline) is left out of this delta, so
 * the host keeps it, and is probed again with the next delta: an unproven
 * family never loads a source face that might shadow it. Library fallback
 * families stay demanded.
 *
 * Shared by the PPTX and XLSX loaders (DOCX folds the same probe into its
 * office-fallback preflight). Each painting realm owns one instance per
 * document: main and worker probe the same host, so both demand alike.
 */
import { probeFontFamilyPresence, type FontFamilyPresence } from './office-fallback.js';
import { normalizeLocalFontMetricFamily } from './local-metrics.js';

/** Presence of `families` (normalised keys): installed ones, and every one
 * whose probe completed. */
export type InstalledFamilyProbe = (families: readonly string[]) => Promise<FontFamilyPresence>;

/** The shared family presence probe (`local()`), batched so every family is
 * probed, against this realm's font set. */
const probeInstalledFamilies: InstalledFamilyProbe = (families) => probeFontFamilyPresence(families);

export class InstalledFamilyExclusion {
  private readonly installed = new Map<string, Promise<boolean | undefined>>();

  constructor(private readonly probe: InstalledFamilyProbe = probeInstalledFamilies) {}

  /** `demand` without the authored families proven installed or unproven. */
  async exclude<T extends { readonly families: readonly string[] }>(demand: T): Promise<T> {
    const fresh = [...new Set(demand.families.map(normalizeLocalFontMetricFamily))]
      .filter((key) => !this.installed.has(key));
    if (fresh.length) {
      const names = new Map(demand.families.map((family) => [normalizeLocalFontMetricFamily(family), family]));
      const probed = this.probe(fresh.map((key) => names.get(key) as string))
        .catch((): FontFamilyPresence => ({ installed: new Set(), checked: new Set() }));
      for (const key of fresh) {
        const outcome = probed.then(({ installed, checked }) => (checked.has(key) ? installed.has(key) : undefined));
        this.installed.set(key, outcome);
        // Unproven families are not cached: the next delta probes them again.
        void outcome.then((value) => { if (value === undefined && this.installed.get(key) === outcome) this.installed.delete(key); });
      }
    }
    // Demand only families proven missing (false); installed (true) and
    // unproven (undefined) families keep the host face.
    const keep = await Promise.all(demand.families.map(async (family) =>
      (await this.installed.get(normalizeLocalFontMetricFamily(family))) === false));
    const families = demand.families.filter((_, index) => keep[index]);
    return families.length === demand.families.length ? demand : Object.freeze({ ...demand, families: Object.freeze(families) });
  }
}

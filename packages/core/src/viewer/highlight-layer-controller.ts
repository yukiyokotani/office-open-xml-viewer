import type { FindMatch } from '../search/find-match';

export interface HighlightSlot { highlightLayer: HTMLDivElement; }

/** Owns mounted search overlays and the optional text measuring canvas. */
export class HighlightLayerController<Slot extends HighlightSlot, Run, Location> {
  private measureContext: CanvasRenderingContext2D | null | undefined;

  constructor(private readonly hooks: {
    slots(): ReadonlyMap<number, Slot>;
    active(): boolean;
    runs(unit: number): Run[] | undefined;
    setRuns(unit: number, runs: Run[]): void;
    paint(unit: number, slot: Slot, runs: Run[], measure: (font: string) => (text: string) => number): void;
    reveal(unit: number): void;
    unitOf(location: Location): number;
  }) {}

  redrawAll(): void {
    for (const [unit, slot] of this.hooks.slots()) this.redrawSlot(unit, slot);
  }

  redrawSlot(unit: number, slot: Slot): void {
    const runs = this.hooks.active() ? this.hooks.runs(unit) : undefined;
    if (!runs) {
      slot.highlightLayer.innerHTML = '';
      return;
    }
    this.hooks.paint(unit, slot, runs, (font) => this.measure(font));
  }

  refreshRuns(unit: number, runs: Run[]): void {
    if (this.hooks.active()) this.hooks.setRuns(unit, runs);
  }

  async activate<Match extends FindMatch<Location>>(match: Match | null): Promise<Match | null> {
    if (match) this.hooks.reveal(this.hooks.unitOf(match.location));
    this.redrawAll();
    return match;
  }

  measure(font: string): (text: string) => number {
    if (this.measureContext === undefined) {
      const canvas = document.createElement('canvas');
      this.measureContext = canvas.getContext('2d');
    }
    const context = this.measureContext;
    if (!context || typeof context.measureText !== 'function') return (text) => text.length;
    context.font = font;
    return (text) => context.measureText(text).width;
  }

  destroy(): void { this.measureContext = undefined; }
}

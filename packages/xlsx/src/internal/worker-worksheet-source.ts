import { isWasmTrap, type WasmParserHost } from '@silurus/ooxml-core';
import type { LegacyXlsDirectSourceDescriptor } from '@silurus/ooxml-core/internal/legacy-xls-source';
import type { WorksheetCursorArchive } from '../worksheet-pull-worker.js';
import type {
  LegacyXlsNativeArchive,
  OwnedLegacyXlsSource,
} from '@silurus/ooxml-legacy-converter/internal/direct-xls-engine';
import { requestXlsFontMeasurement } from '@silurus/ooxml-legacy-converter/internal/xls-font-worker';

export interface WorkerWorksheetArchive extends WorksheetCursorArchive {
  assert_healthy(): void;
  parse(): Uint8Array;
  extract_image(path: string): Uint8Array;
}

export interface OoxmlWorksheetArchive extends WorkerWorksheetArchive {
  free(): void;
  resource_usage(): Uint8Array;
  to_markdown(): string;
}

type OpenLegacyXlsSource = (
  bytes: Uint8Array,
  descriptor: LegacyXlsDirectSourceDescriptor,
) => Promise<OwnedLegacyXlsSource>;

const openLegacyXlsSource: OpenLegacyXlsSource = async (bytes, descriptor) => {
  const engine = await import('@silurus/ooxml-legacy-converter/internal/direct-xls-engine');
  return await engine.openLegacyXlsSource(bytes, descriptor);
};

/** Own exactly one worker-local XLSX or direct XLS source. */
export class WorkerWorksheetSourceOwner<TArchive extends OoxmlWorksheetArchive> {
  private legacy: OwnedLegacyXlsSource | undefined;
  private measurementController: AbortController | undefined;
  private measuredMaximumDigitWidth: number | undefined;

  constructor(
    private readonly ooxmlHost: WasmParserHost<TArchive>,
    private readonly openDirect: OpenLegacyXlsSource = openLegacyXlsSource,
  ) {}

  get kind(): 'ooxml' | 'legacy-xls' {
    return this.legacy ? 'legacy-xls' : 'ooxml';
  }

  async openLegacy(
    bytes: Uint8Array,
    descriptor: LegacyXlsDirectSourceDescriptor,
    measure = false,
  ): Promise<LegacyXlsNativeArchive> {
    if (this.legacy || this.ooxmlHost.archive) throw new Error('Workbook source already loaded');
    const owned = await this.openDirect(bytes, descriptor);
    this.legacy = owned;
    const controller = new AbortController();
    this.measurementController = controller;
    try {
      const { configureLegacyXlsMeasurement } = await import(
        '@silurus/ooxml-legacy-converter/internal/direct-xls-engine'
      );
      this.measuredMaximumDigitWidth = await configureLegacyXlsMeasurement(
        owned.archive,
        measure ? requestXlsFontMeasurement(self) : undefined,
        controller.signal,
      );
    } catch (error) {
      try { this.closeLegacy(); } catch {}
      throw error;
    }
    return owned.archive;
  }

  get maximumDigitWidth(): number | undefined {
    return this.measuredMaximumDigitWidth;
  }

  cursor(): WorkerWorksheetArchive | null {
    return this.legacy?.archive ?? this.ooxmlHost.archive;
  }

  execute<T>(operation: (archive: WorkerWorksheetArchive) => T): T {
    const archive = this.cursor();
    if (!archive) throw new Error('Workbook not loaded');
    if (!this.legacy) return this.ooxmlHost.run(() => operation(archive));
    try {
      return operation(archive);
    } catch (error) {
      if (isWasmTrap(error)) {
        try { this.closeLegacy(); } catch {}
      }
      throw error;
    }
  }

  ooxml(operation: string): TArchive {
    if (this.legacy) throw new Error(`${operation} is unsupported for direct legacy XLS sources`);
    const archive = this.ooxmlHost.archive;
    if (!archive) throw new Error('Workbook not loaded');
    return archive;
  }

  closeLegacy(): void {
    const owned = this.legacy;
    this.legacy = undefined;
    this.measurementController?.abort();
    this.measurementController = undefined;
    this.measuredMaximumDigitWidth = undefined;
    owned?.closeArchive();
  }
}

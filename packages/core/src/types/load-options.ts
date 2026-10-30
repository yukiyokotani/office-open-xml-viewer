import type { CjkFallback } from '../fonts/cjk-fallback.js';
import type { MathRenderer } from '../math/mathjax';
import type { ChartThreeDRenderer } from '../chart/three-d-contract';
import type { ChartRegionMapRenderer } from '../chart/region-map-contract';
import type { ChartExRenderer } from '../chart/chart-ex-contract';
import type { TiffRenderer } from '../image/tiff-contract';
import type { OoxmlResourceMetrics } from './resource-metrics.js';
import type { FontSource } from '../fonts/source.js';

/** A positive safe-integer byte count, or `null` to disable one public limit. */
export type OoxmlResourceLimit = number | null;

/** Admission limits for the inflated contents of one OOXML package session. */
export interface OoxmlResourceLimits {
  /**
   * Maximum permitted inflated size for any one archive entry, including
   * media. Enforced against both the ZIP declaration and actual output.
   */
  maxArchiveEntryBytes?: OoxmlResourceLimit;
  /** Maximum actual inflated bytes across distinct entries in the session. */
  maxTotalInflatedBytes?: OoxmlResourceLimit;
  /**
   * Maximum number of entries admitted from the archive central directory.
   * `null` disables this configurable limit only; the internal implementation
   * hard ceiling remains active.
   */
  maxArchiveEntries?: OoxmlResourceLimit;
}

/**
 * Per-load XLSX worksheet admission limits, validated before load effects and
 * snapshotted for that load. DOCX and PPTX ignore them. An omitted value, an
 * empty object, or an omitted field uses the default (100,000 rows / 250,000
 * cells / 32 MiB owned UTF-8 / 64 MiB JSON). Each supplied field must be a
 * positive safe integer no greater than `Number.MAX_SAFE_INTEGER - 1`; `null`
 * is invalid.
 *
 * These are logical budgets, not physical memory measurements. Raising them
 * accepts additional out-of-memory risk; catching allocation failures is not
 * guaranteed.
 * Each dimension of the aggregate derived worksheet cache is the larger of its
 * default (200,000 rows / 500,000 cells / 64 MiB owned / 128 MiB JSON) and the
 * corresponding limit. Each renderer coordinate index independently allows
 * `max(250000, maxCells)` entries, charged per unique key. ZIP, input, wire,
 * representation and copy limits remain separate. The non-retaining Node
 * `worksheetRows` path keeps streaming semantics: cumulative budgets apply
 * only to models it materializes.
 */
export interface XlsxWorksheetLimits {
  /** Maximum retained row records for one worksheet (not grid row coordinates). */
  maxRows?: number;
  /**
   * Maximum retained cell records for one worksheet. This counts records, not
   * grid coordinates or unique indexes. Delimited-text parsing counts every
   * parsed field, including blanks; its fixed 64 MiB input limit is separate.
   */
  maxCells?: number;
  /**
   * Maximum UTF-8 bytes of all strings owned by `Cell.value` across one
   * worksheet. This includes typed-value discriminators, rich-text and
   * phonetic strings, plus formula text outside `Cell.value`. Shared strings
   * are charged per cell after resolution.
   */
  maxOwnedUtf8Bytes?: number;
  /**
   * Maximum structural JSON UTF-8 bytes for one worksheet model, including
   * covered ancillary data. Browser and Node measure `JSON.stringify` of the
   * assembled model. The native Rust path measures complete `serde_json`
   * output, which is a separate boundary.
   */
  maxJsonBytes?: number;
}

/** Format-neutral progress reported while a progressive layout pass runs. */
export interface ProgressiveLayoutProgress {
  /**
   * Units committed by the current pass (pages for DOCX, slides for PPTX).
   * This is telemetry rather than a final count and may move backward when a
   * converging layout pass restarts.
   */
  committedUnits: number;
}

/** Format-neutral publication of a newly paintable progressive prefix. */
export interface ProgressiveLayoutPartial {
  /** Units currently available to paint (pages for DOCX, slides for PPTX). */
  availableUnits: number;
  /** Final unit count when the format knows it before layout completes. */
  totalUnits?: number;
  /** Whether the publication is authoritative rather than provisional. */
  exact: boolean;
}

/**
 * Common load-time options shared by the docx / pptx / xlsx
 * `Document.load` / `Presentation.load` / `Workbook.load` factories and their
 * viewer wrappers.
 *
 * This is the single source of truth — each package re-exports this exact type
 * as its `LoadOptions` so application code can pass one options object to any
 * of the three.
 */
export interface LoadOptions {
  /**
   * Ordered font sources consulted for authored and library fallback families
   * the host lacks. Default: none, so no font is requested from any network
   * host and the canvas falls back to locally available fonts.
   *
   * Build the entries with the factories exported next to this type
   * (`googleFonts()`, `fontCssApi()`, `fontCss()`, `fontFiles()`). The first
   * source that declares a face for a family claims it; later sources are not
   * consulted for that family, and a claimed family never retries another
   * source if its faces fail to load (the family then keeps its host fallback
   * with a `console.warn`). Faces register under their real CSS family in the
   * painting realm, and only the `unicode-range` subsets that intersect the
   * document's painted characters are loaded.
   *
   * A hosted source receives the end user's IP address and User-Agent, which
   * may have privacy / GDPR implications for your application.
   */
  fonts?: readonly FontSource[];
  /**
   * @deprecated Since 0.90; removed in 0.91. Replace `useGoogleFonts: true`
   * with `fonts: [googleFonts()]`. Every format normalises it to exactly that and
   * warns once per JavaScript realm; passing it together with `fonts` throws
   * a `TypeError`, and `false` is a no-op that still warns.
   *
   * Opt in to loading webfont substitutes from Google Fonts
   * (`fonts.googleapis.com`). Default `false` — the canvas falls back to
   * locally available fonts. When enabled, end-user IP / User-Agent is sent
   * to Google, which may have privacy / GDPR implications for your
   * application.
   */
  useGoogleFonts?: boolean;
  /**
   * Regional preference used when Han text reaches font fallback and the
   * document has not already identified a region. The requested font remains
   * first; recognized regional CJK font names, an East Asian run language where
   * available, and unambiguous Kana or Hangul take priority. Text without Han
   * keeps its existing font route.
   *
   * `auto` (also the omitted default) snapshots HTML lang, then
   * navigator.languages / navigator.language; without a usable CJK language it
   * uses `jp`. Bare `zh` uses `sc`. An explicit value makes the regional choice
   * independent of the host locale, but does not enable webfonts or by itself
   * guarantee pixel-identical output across different font environments.
   */
  cjkFallback?: CjkFallback;
  /**
   * Password for an encrypted OOXML file ([MS-OFFCRYPTO] Agile Encryption).
   *
   * Password-protected Office documents are CFB (OLE2) containers, not ZIPs.
   * When this is set and the input is Agile-encrypted, `load()` decrypts it on
   * the main thread (via WebCrypto) and parses the recovered plaintext ZIP.
   *
   * Errors (thrown as {@link import('../errors/ooxml-error').OoxmlError}):
   *   - no `password` on an encrypted file → code `'encrypted'`
   *   - wrong `password`                   → code `'invalid-password'`
   *   - a non-Agile scheme (Standard / Extensible / legacy) → code
   *     `'unsupported-encryption'`
   *
   * Note: Agile Encryption uses a high password-hash spin count (commonly
   * 100,000), so decryption of a protected file adds roughly a second of
   * WebCrypto work before parsing begins.
   *
   * Security notes:
   *   - This value is held as an ordinary JS `string` in memory for the
   *     duration of key derivation. The library does not zero it, and does
   *     not wrap it in a `SecureString`-equivalent — it becomes eligible for
   *     garbage collection like any other string once nothing references it,
   *     but no explicit wipe is performed. It is never logged or included in
   *     thrown errors.
   *   - Decryption recovers the plaintext but does not verify the file's HMAC
   *     data-integrity tag ([MS-OFFCRYPTO] §2.3.4.14), so ciphertext tampering
   *     is not detected — see "Security & Privacy" in the README.
   */
  password?: string;
  /**
   * Override the URL the parser worker fetches the WebAssembly module from.
   *
   * By default each format resolves the `.wasm` asset that ships next to its
   * bundle (relative to the module URL), so no configuration is needed. Set
   * this to serve the parser WASM from a CDN or a self-hosted path instead — a
   * relative value is resolved against the current document URL. The same
   * dependency-injection contract across docx / pptx / xlsx.
   *
   * The referenced file must be the matching format's `*_parser_bg.wasm`
   * artifact (the one wasm-bindgen emitted for that parser); pointing it at a
   * mismatched or missing file makes `load()` reject when the worker
   * instantiates it.
   */
  wasmUrl?: string | URL;
  /**
   * @deprecated Use `resourceLimits.maxArchiveEntryBytes`. Scheduled for
   * removal in a future breaking release.
   *
   * Existing positive safe-integer values remain an all-entry inflated-byte
   * limit. Zero, negative, and NaN values retain their historical fallback
   * behavior; other invalid positive values reject during `load()`.
   */
  maxZipEntryBytes?: number;
  /**
   * Inflated archive admission limits for one document session. Omitted fields
   * use the library defaults. A positive safe integer overrides a default;
   * `null` disables that configurable limit only. Limits are admission policy,
   * not guarantees of exact browser-process memory use.
   */
  resourceLimits?: OoxmlResourceLimits;
  /**
   * XLSX-only logical worksheet admission limits; ignored by DOCX and PPTX.
   * Omitted fields use the library defaults. See {@link XlsxWorksheetLimits}.
   */
  xlsxWorksheetLimits?: XlsxWorksheetLimits;
  /**
   * Emit one content-free resource-usage card after load succeeds or fails.
   * Includes observed archive counters and configured limits, but never source
   * URLs, part names, document text, passwords, or error messages.
   */
  debug?: boolean;
  /**
   * Receive the initial content-free, machine-readable report that powers the
   * debug card, without enabling console output. After resource options validate,
   * the callback runs once when the current load settles, including failed loads
   * for which no renderer instance is returned. The callback is not awaited;
   * synchronous exceptions and rejected promises are ignored and never change
   * load results.
   *
   * A browser report covers the underlying document/workbook/presentation
   * factory. It does not wait for a Viewer's first canvas paint; that paint and
   * later lazy worksheet, slide, image, or media access may increase counters or
   * surface a separate render error. Successfully opened packages include the
   * declared package total and source byte size in the report. On a successful
   * load, call `getResourceMetrics()` on the returned engine or Viewer for a fresh
   * snapshot that includes subsequently observed lazy package work.
   */
  onResourceMetrics?: (metrics: OoxmlResourceMetrics) => void;
  /**
   * Opt-in worker liveness limit in milliseconds. For an ordinary load it is
   * the response deadline for the parse request. When progressive DOCX/PPTX
   * preparation runs in `mode: 'worker'`, it becomes a silence interval that is restarted by each layout
   * progress or partial publication, so a busy worker is not timed out merely
   * because the complete document takes longer. Silence before the first
   * publication rejects `load()`; silence after `load()` has resolved leaves
   * `layoutComplete` false and makes `waitUntilLayoutComplete()` reject (and is
   * also delivered to the layout-completion/error callback when configured).
   * **Default: unlimited.** A worker that throws or fails to load rejects
   * immediately regardless of this value.
   */
  workerTimeoutMs?: number;
  /**
   * Opt-in OMML equation engine (MathJax + STIX Two Math, ~4 MB). Inject it
   * **once** here and every render of this document / presentation / workbook
   * uses it — the same dependency-injection contract across all three formats
   * and their viewers. Import it from the separate `@silurus/ooxml/math` entry
   * (`import { math } from '@silurus/ooxml/math'`). Omit it and equations are
   * skipped and the engine asset is not fetched or evaluated. The on-demand,
   * self-contained render-worker asset retains the small worker-side loader.
   */
  math?: MathRenderer;
  /**
   * Opt-in model-space 3-D chart renderer. Import `threeD` from the separate
   * `@silurus/ooxml/three-d` entry and inject it once. When omitted, authored
   * 3-D groups use their canonical 2-D chart family and the mesh/camera renderer
   * is not loaded or evaluated in main mode. The on-demand render-worker asset
   * is self-contained and includes its worker-side implementation.
   */
  threeD?: ChartThreeDRenderer;
  /**
   * Opt-in offline ChartEx Region Map renderer with fixed Natural Earth
   * geometry. Import `regionMap` from
   * `@silurus/ooxml/region-map` and inject it once. Omit it when geospatial
   * charts are not needed so the asset is not loaded or evaluated in main mode.
   * The on-demand render-worker asset includes its worker-side implementation.
   */
  regionMap?: ChartRegionMapRenderer;
  /**
   * Opt in to Microsoft ChartEx (`cx:*`) chart families. Import `chartEx`
   * from `@silurus/ooxml/chart-ex` and inject it once. Classic DrawingML 2-D
   * charts remain available without this module. In main mode the implementation
   * stays outside the format entry; the on-demand render-worker asset remains
   * self-contained and includes its worker-side implementation.
   */
  chartEx?: ChartExRenderer;
  /**
   * Opt in to bounded TIFF 6.0 image decoding. Import `tiff` from the separate
   * `@silurus/ooxml/tiff` entry and inject it once. The built-in codec accepts
   * stripped bilevel, 8-bit grayscale, RGB, RGBA and process-CMYK images, plus
   * CCITT Group 4 bilevel images. Without this option, recognized TIFF images
   * use a visible unavailable-image placeholder and the document keeps
   * rendering. Unsupported or malformed input makes standalone codec calls and
   * DOCX/PPTX rendering report `TiffDecodeError`; XLSX rendering, including
   * `XlsxViewer`, contains it at that picture and shows the placeholder. The
   * built-in codec works in main and worker modes.
   */
  tiff?: TiffRenderer;
}

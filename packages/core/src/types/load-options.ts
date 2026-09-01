import type { MathRenderer } from '../math/mathjax';
import type { ChartThreeDRenderer } from '../chart/three-d-contract';
import type { ChartRegionMapRenderer } from '../chart/region-map-contract';
import type { ChartExRenderer } from '../chart/chart-ex-contract';
import type { TiffRenderer } from '../image/tiff-contract';
import type { OoxmlResourceMetrics } from './resource-metrics.js';

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
   * Opt in to loading webfont substitutes from Google Fonts
   * (`fonts.googleapis.com`). Default `false` — the canvas falls back to
   * locally available fonts.
   *
   * When enabled, end-user IP / User-Agent is sent to Google, which may
   * have privacy / GDPR implications for your application. To avoid the
   * third-party request, host the substitutes yourself and reference them
   * via `@font-face` in your application CSS.
   */
  useGoogleFonts?: boolean;
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
   * Opt-in OMML equation engine (MathJax + STIX Two Math, ~3 MB). Inject it
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
   * uncompressed, 8-bit, chunky process-CMYK strips; other TIFF classes fail
   * closed. Omit it to keep the decoder out of ordinary format bundles; TIFF
   * parts are then skipped without aborting the surrounding document render.
   */
  tiff?: TiffRenderer;
}

import { decodeDataUrl, WasmParserHost } from '@silurus/ooxml-core';
import {
  decodeOoxmlResourceUsage,
  resourcePolicyForWasm,
  serializeWorkerError,
  type PullSessionCommand,
  type PullSessionResponse,
} from '@silurus/ooxml-core/worker';
import { PresentationPreflightBuilder } from './presentation-preflight.js';
import { isSlidePullCommand, SlidePullWorker } from './slide-pull-worker.js';
import type {
  PresentationBootstrap,
  PptxWorkerRequest,
  PptxWorkerResponse,
} from './worker-protocol.js';
import init, { PptxArchive, reinit } from './wasm/pptx_parser.js';
import { WorkerPresentationSourceOwner } from './internal/worker-presentation-source.js';

const host = new WasmParserHost<PptxArchive>(init, {
  freeArchive: (archive) => archive.free(),
  reinit,
});
const source = new WorkerPresentationSourceOwner(host);
let ooxmlWasmInput: Parameters<typeof host.setWasmInput>[0] | undefined;

let preflightBuilder: PresentationPreflightBuilder | null = null;
type PresentationLifecycleState = 'empty' | 'opening' | 'ready' | 'failed';
let presentationState: PresentationLifecycleState = 'empty';

function reservePresentationParse(): void {
  if (presentationState !== 'empty') {
    const error = new Error('this PPTX worker already owns a presentation parse');
    error.name = 'PptxWorkerStateError';
    throw Object.assign(error, { code: 'ooxml-pptx-parse-already-started' });
  }
  presentationState = 'opening';
}

const slidePull = new SlidePullWorker(
  () => source.cursor(),
  (slideIndex, slide, usage) => {
    if (!preflightBuilder) return;
    if (slideIndex !== preflightBuilder.acceptedSlideCount) {
      throw new Error(
        `PPTX preflight expected slide ${preflightBuilder.acceptedSlideCount}, received ${slideIndex}`,
      );
    }
    return preflightBuilder.prepareSlide(slide, usage);
  },
  (operation) => {
    return source.execute(operation);
  },
);

const post = (
  message: PptxWorkerResponse | PullSessionResponse<ArrayBuffer, number>,
  transfer?: Transferable[],
) => (self.postMessage as (value: unknown, transfer?: Transferable[]) => void)(message, transfer);

self.onmessage = async (
  event: MessageEvent<PptxWorkerRequest | PullSessionCommand<number>>,
) => {
  const request = event.data;

  if (isSlidePullCommand(request)) {
    await slidePull.dispatchSafely(request, post);
    return;
  }

  if (request.kind === 'init') {
    ooxmlWasmInput = decodeDataUrl(request.wasmUrl) ?? request.wasmUrl;
    return;
  }

  const id = request.id;
  let ownsParseReservation = false;
  try {
    // Reservation must happen before the first await, but still inside the
    // correlated error boundary so poison/identity failures cannot orphan a
    // main-side request indefinitely.
    if (request.kind === 'openSlideSession') slidePull.reserveOpen(request);
    if (request.kind === 'parse') {
      reservePresentationParse();
      ownsParseReservation = true;
    }
    if (request.kind === 'openSlideSession') {
      await slidePull.open(request.slideIndex, request);
      await slidePull.postOpenedSafely(
        request,
        () => post({
          kind: 'slideSessionOpened',
          id,
          sessionId: request.sessionId,
          operationId: request.operationId,
          generation: request.generation,
        }),
        (error) => post({ kind: 'error', id, ...serializeWorkerError(error) }),
      );
      return;
    }

    if (request.kind === 'parse') await slidePull.reset();
    await slidePull.run(async () => {
      if (request.kind !== 'parse' && source.cursor()) {
        source.execute((archive) => archive.assert_healthy());
      }

      if (request.kind === 'parse') {
        preflightBuilder = null;
        if (request.source) {
          await source.openModelSource(
            new Uint8Array(request.buffer),
            request.source,
            request.sourceTransfer,
          );
        } else {
          if (ooxmlWasmInput === undefined) throw new Error('PPTX WASM input was not configured');
          host.setWasmInput(ooxmlWasmInput);
          await host.ensureReady();
          const [maxEntry, maxTotal, maxEntries] = resourcePolicyForWasm(request.resourcePolicy);
          host.run(() => {
            const opened = new PptxArchive(
              new Uint8Array(request.buffer), maxEntry, maxTotal, maxEntries,
            );
            host.setArchive(opened);
          });
        }
        const bootstrap = JSON.parse(new TextDecoder().decode(
          source.execute((current) => current.presentation_bootstrap()),
        )) as PresentationBootstrap;
        // Ordinary loads retain compact facts in the worker and return them at
        // the end. Progressive main-mode loads decode each sequential slide in
        // Window so the presentation can publish the opening prefix itself;
        // keeping a second builder here would duplicate the bounded projection.
        preflightBuilder = request.progressiveLayout
          ? null
          : new PresentationPreflightBuilder(bootstrap, { cjkFallback: request.cjkFallback });
        post({ kind: 'presentationOpened', id, bootstrap });
        presentationState = 'ready';
        return;
      }

      const archive = source.cursor();
      if (!archive) throw new Error('No pptx loaded');

      if (request.kind === 'finishPresentationPreflight') {
        if (!preflightBuilder) throw new Error('PPTX presentation preflight is not active');
        const preflight = preflightBuilder.finish();
        preflightBuilder = null;
        post({ kind: 'presentationPreflightReady', id, preflight });
        return;
      }

      if (request.kind === 'extractMedia') {
        const bytes = source.extractMedia(request.path).buffer as ArrayBuffer;
        post({ kind: 'mediaExtracted', id, bytes }, [bytes]);
        return;
      }

      if (request.kind === 'extractImage') {
        const bytes = source.execute(
          (current) => current.extract_image(request.path).buffer as ArrayBuffer,
        );
        post({ kind: 'imageExtracted', id, bytes }, [bytes]);
        return;
      }

      if (request.kind === 'extractFont') {
        const bytes = source.extractFont(request.path).buffer as ArrayBuffer;
        post({ kind: 'fontExtracted', id, bytes }, [bytes]);
        return;
      }

      if (request.kind === 'resourceUsage') {
        const bytes = source.resourceUsage();
        const usage = bytes === undefined ? undefined : decodeOoxmlResourceUsage(bytes);
        post({ kind: 'resourceUsage', id, usage });
        return;
      }

      if (request.kind === 'toMarkdown') {
        post({ kind: 'markdownRendered', id, markdown: source.toMarkdown() });
      }
    });
  } catch (error) {
    if (ownsParseReservation) {
      presentationState = 'failed';
      try { source.closeModelSource(); } catch {}
    }
    if (request.kind === 'openSlideSession') slidePull.abandonOpen(request.sessionId);
    try {
      post({ kind: 'error', id, ...serializeWorkerError(error) });
    } catch {
      // Ownership cleanup already converged; the response channel is gone.
    }
  }
};

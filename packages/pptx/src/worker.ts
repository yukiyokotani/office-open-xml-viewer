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

const host = new WasmParserHost<PptxArchive>(init, {
  freeArchive: (archive) => archive.free(),
  reinit,
});

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
  () => host.archive,
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
    const archive = host.archive;
    if (!archive) throw new Error('Presentation not loaded');
    return host.run(() => operation(archive));
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
    host.setWasmInput(decodeDataUrl(request.wasmUrl) ?? request.wasmUrl);
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
      await host.ensureReady();
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
      await host.ensureReady();
      if (request.kind !== 'parse' && host.archive) {
        const retained = host.archive;
        host.run(() => retained.assert_healthy());
      }

      if (request.kind === 'parse') {
        preflightBuilder = null;
        const [maxEntry, maxTotal, maxEntries] = resourcePolicyForWasm(request.resourcePolicy);
        const bootstrap = host.run(() => {
          const archive = new PptxArchive(
            new Uint8Array(request.buffer),
            maxEntry,
            maxTotal,
            maxEntries,
          );
          host.setArchive(archive);
          return JSON.parse(
            new TextDecoder().decode(archive.presentation_bootstrap()),
          ) as PresentationBootstrap;
        });
        // Ordinary loads retain compact facts in the worker and return them at
        // the end. Progressive main-mode loads decode each sequential slide in
        // Window so the presentation can publish the opening prefix itself;
        // keeping a second builder here would duplicate the bounded projection.
        preflightBuilder = request.progressiveLayout
          ? null
          : new PresentationPreflightBuilder(bootstrap);
        post({ kind: 'presentationOpened', id, bootstrap });
        presentationState = 'ready';
        return;
      }

      const archive = host.archive;
      if (!archive) throw new Error('No pptx loaded');

      if (request.kind === 'finishPresentationPreflight') {
        if (!preflightBuilder) throw new Error('PPTX presentation preflight is not active');
        const preflight = preflightBuilder.finish();
        preflightBuilder = null;
        post({ kind: 'presentationPreflightReady', id, preflight });
        return;
      }

      if (request.kind === 'extractMedia') {
        const bytes = host.run(() => archive.extract_media(request.path).buffer as ArrayBuffer);
        post({ kind: 'mediaExtracted', id, bytes }, [bytes]);
        return;
      }

      if (request.kind === 'extractImage') {
        const bytes = host.run(() => archive.extract_image(request.path).buffer as ArrayBuffer);
        post({ kind: 'imageExtracted', id, bytes }, [bytes]);
        return;
      }

      if (request.kind === 'extractFont') {
        const bytes = host.run(() => archive.extract_font(request.path).buffer as ArrayBuffer);
        post({ kind: 'fontExtracted', id, bytes }, [bytes]);
        return;
      }

      if (request.kind === 'resourceUsage') {
        const usage = decodeOoxmlResourceUsage(host.run(() => archive.resource_usage()));
        post({ kind: 'resourceUsage', id, usage });
        return;
      }

      if (request.kind === 'toMarkdown') {
        post({ kind: 'markdownRendered', id, markdown: host.run(() => archive.to_markdown()) });
      }
    });
  } catch (error) {
    if (ownsParseReservation) presentationState = 'failed';
    if (request.kind === 'openSlideSession') slidePull.abandonOpen(request.sessionId);
    try {
      post({ kind: 'error', id, ...serializeWorkerError(error) });
    } catch {
      // Ownership cleanup already converged; the response channel is gone.
    }
  }
};

import { loadFontDemandCollector } from './font-demand-collector.js';
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
import type { WorkerPresentationSourceOwner } from './internal/worker-presentation-source.js';

const host = new WasmParserHost<PptxArchive>(init, {
  freeArchive: (archive) => archive.free(),
  reinit,
});
let source: WorkerPresentationSourceOwner<PptxArchive> | undefined;

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
  () => source?.cursor() ?? host.archive,
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
    if (source) return source.execute(operation);
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
      if (!source) await host.ensureReady();
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
      if (request.kind === 'parse' ? !request.source : !source) await host.ensureReady();
      if (request.kind !== 'parse' && (source?.cursor() ?? host.archive)) {
        if (source) source.execute((archive) => archive.assert_healthy());
        else host.run(() => host.archive?.assert_healthy());
      }

      if (request.kind === 'parse') {
        preflightBuilder = null;
        let bootstrap: PresentationBootstrap;
        if (request.source) {
          if (!request.sourceOwnerUrl) throw new TypeError('PPTX source owner URL is missing');
          const { WorkerPresentationSourceOwner } = await import(/* @vite-ignore */ request.sourceOwnerUrl) as typeof import('./internal/worker-presentation-source.js');
          source = new WorkerPresentationSourceOwner(host);
          await source.openModelSource(
            new Uint8Array(request.buffer),
            request.source,
            request.sourceTransfer,
          );
          bootstrap = JSON.parse(new TextDecoder().decode(
            source.execute((current) => current.presentation_bootstrap()),
          )) as PresentationBootstrap;
        } else {
          const [maxEntry, maxTotal, maxEntries] = resourcePolicyForWasm(request.resourcePolicy);
          bootstrap = host.run(() => {
            const archive = new PptxArchive(
              new Uint8Array(request.buffer), maxEntry, maxTotal, maxEntries,
            );
            host.setArchive(archive);
            return JSON.parse(new TextDecoder().decode(
              archive.presentation_bootstrap(),
            )) as PresentationBootstrap;
          });
        }
        // Ordinary loads retain compact facts in the worker and return them at
        // the end. Progressive main-mode loads decode each sequential slide in
        // Window so the presentation can publish the opening prefix itself;
        // keeping a second builder here would duplicate the bounded projection.
        preflightBuilder = request.progressiveLayout
          ? null
          : new PresentationPreflightBuilder(bootstrap, { cjkFallback: request.cjkFallback, collectFontDemand: request.collectFontDemand === true ? await loadFontDemandCollector(true) : undefined });
        post({ kind: 'presentationOpened', id, bootstrap });
        presentationState = 'ready';
        return;
      }

      const archive = source?.cursor() ?? host.archive;
      if (!archive) throw new Error('No pptx loaded');

      if (request.kind === 'finishPresentationPreflight') {
        if (!preflightBuilder) throw new Error('PPTX presentation preflight is not active');
        const preflight = preflightBuilder.finish();
        const partFailures = preflightBuilder.partFailures;
        preflightBuilder = null;
        post({
          kind: 'presentationPreflightReady',
          id,
          preflight,
          ...(partFailures.length === 0 ? {} : { partFailures }),
        });
        return;
      }

      if (request.kind === 'extractMedia') {
        const media = source
          ? source.extractMedia(request.path)
          : host.run(() => host.archive!.extract_media(request.path));
        const bytes = source
          ? source.copyBytes(media)
          : media.buffer as ArrayBuffer;
        post({ kind: 'mediaExtracted', id, bytes }, [bytes]);
        return;
      }

      if (request.kind === 'extractImage') {
        const bytes = source
          ? source.extractImage(request.path)
          : host.run(() => archive.extract_image(request.path).buffer as ArrayBuffer);
        post({ kind: 'imageExtracted', id, bytes }, [bytes]);
        return;
      }

      if (request.kind === 'extractFont') {
        const font = source
          ? source.extractFont(request.path)
          : host.run(() => host.archive!.extract_font(request.path));
        const bytes = source
          ? source.copyBytes(font)
          : font.buffer as ArrayBuffer;
        post({ kind: 'fontExtracted', id, bytes }, [bytes]);
        return;
      }

      if (request.kind === 'resourceUsage') {
        const bytes = source ? source.resourceUsage() : host.run(() => host.archive!.resource_usage());
        const usage = bytes === undefined ? undefined : decodeOoxmlResourceUsage(bytes);
        post({ kind: 'resourceUsage', id, usage });
        return;
      }

      if (request.kind === 'toMarkdown') {
        post({ kind: 'markdownRendered', id, markdown: source ? source.toMarkdown() : host.run(() => host.archive!.to_markdown()) });
      }
    });
  } catch (error) {
    if (ownsParseReservation) {
      presentationState = 'failed';
      try { source?.closeModelSource(); } catch {}
    }
    if (request.kind === 'openSlideSession') slidePull.abandonOpen(request.sessionId);
    try {
      post({ kind: 'error', id, ...serializeWorkerError(error) });
    } catch {
      // Ownership cleanup already converged; the response channel is gone.
    }
  }
};

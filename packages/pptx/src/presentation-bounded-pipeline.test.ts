import { describe, expect, it } from 'vitest';
import type { WorkerBridgeTransport } from '@silurus/ooxml-core';
import {
  PULL_SESSION_PROTOCOL,
  type PullSessionCommand,
  type PullSessionResponse,
} from '@silurus/ooxml-core/worker';
import { PptxPresentation } from './presentation.js';
import type { PptxWorkerRequest } from './worker-protocol.js';
import type { PptxSlideRepository } from './slide-repository.js';
import type { Slide } from './types.js';

type PullResponse = PullSessionResponse<ArrayBuffer, number>;

function slide(index: number): Slide {
  return {
    index,
    slideNumber: index + 1,
    partName: `ppt/slides/slide${index + 1}.xml`,
    background: null,
    elements: [],
  };
}

function pullResponse(
  command: PullSessionCommand<number>,
  value: Record<string, unknown>,
): PullResponse {
  return {
    protocol: PULL_SESSION_PROTOCOL,
    sessionId: command.sessionId,
    operationId: command.operationId,
    generation: command.generation,
    requestId: command.requestId,
    ...value,
  } as PullResponse;
}

describe('PptxPresentation bounded main-mode pipeline', () => {
  it('preflights sequentially, retains compact synchronous facts, and reloads a slide on demand', async () => {
    const pullCommands: PullSessionCommand<number>[] = [];
    let pullRequestId = 1;
    const transport: WorkerBridgeTransport<PullResponse> = {
      request: async (build) => {
        const command = build(pullRequestId++) as PullSessionCommand<number>;
        pullCommands.push(command);
        if (command.kind === 'pull') {
          const payload = new TextEncoder().encode(
            JSON.stringify(slide(command.sessionId <= 2 ? command.sessionId - 1 : 0)),
          ).buffer;
          return pullResponse(command, {
            kind: 'chunk',
            sequence: command.sequence,
            byteLength: payload.byteLength,
            done: true,
            payload,
          });
        }
        return pullResponse(command, { kind: 'accepted', command: command.kind });
      },
      forgetOrphaned: () => undefined,
      terminate: () => undefined,
    };
    const ordinary: PptxWorkerRequest[] = [];
    let ordinaryId = 100;
    const bootstrap = {
      slideCount: 2,
      slideWidth: 9144000,
      slideHeight: 6858000,
      defaultTextColor: '111111',
      majorFont: 'Aptos Display',
      minorFont: 'Aptos',
      hlinkColor: '0563C1',
      folHlinkColor: null,
      slides: [
        { index: 0, partName: 'ppt/slides/slide1.xml' },
        { index: 1, partName: 'ppt/slides/slide2.xml' },
      ],
    } as const;
    const compact = {
      ...bootstrap,
      slides: [
        {
          index: 0,
          partName: bootstrap.slides[0].partName,
          notes: 'speaker note',
          hidden: false,
          mediaElements: [],
          comments: [{ id: '{ROOT}', modernAuthorId: '{ADA}', author: 'Ada', x: 1, y: 2, status: 'active', text: 'Review' }],
        },
        { index: 1, partName: bootstrap.slides[1].partName, notes: null, hidden: true, mediaElements: [] },
      ],
      fontPreloadNames: ['Aptos Display', 'Aptos'],
    } as const;
    const bridge = {
      request: async (build: (id: number) => PptxWorkerRequest) => {
        const request = build(ordinaryId++);
        ordinary.push(request);
        if (request.kind === 'parse') {
          return { kind: 'presentationOpened', id: request.id, bootstrap };
        }
        if (request.kind === 'openSlideSession') {
          return {
            kind: 'slideSessionOpened',
            id: request.id,
            sessionId: request.sessionId,
            operationId: request.operationId,
            generation: request.generation,
          };
        }
        if (request.kind === 'finishPresentationPreflight') {
          return { kind: 'presentationPreflightReady', id: request.id, preflight: compact };
        }
        throw new Error(`unexpected request ${request.kind}`);
      },
      transport: () => transport,
    };
    const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
    instance._mode = 'main';
    instance._bridge = bridge;
    const presentation = instance as unknown as PptxPresentation;
    type Policy = {
      readonly maxArchiveEntryBytes: null;
      readonly maxTotalInflatedBytes: null;
    };
    const policy: Policy = { maxArchiveEntryBytes: null, maxTotalInflatedBytes: null };

    await (presentation as unknown as {
      _parse(buffer: ArrayBuffer, policy: Policy): Promise<void>;
    })._parse(new ArrayBuffer(4), policy);

    expect(presentation.slideCount).toBe(2);
    expect(presentation.getNotes(0)).toBe('speaker note');
    expect(presentation.getComments(0)).toEqual(compact.slides[0].comments);
    expect(presentation.getComments(1)).toEqual([]);
    expect(presentation.isHidden(1)).toBe(true);
    expect(presentation.getSlideIndexByPartName('ppt/slides/slide2.xml')).toBe(1);
    expect(ordinary.filter((request) => request.kind === 'openSlideSession')).toHaveLength(2);
    expect(pullCommands.filter((command) => command.kind === 'ack')).toHaveLength(2);

    const repository = instance._slides as PptxSlideRepository;
    await expect(repository.withSlide(0, (value) => value)).resolves.toEqual(slide(0));
    expect(ordinary.filter((request) => request.kind === 'openSlideSession')).toHaveLength(3);
    expect(pullCommands.filter((command) => command.kind === 'ack')).toHaveLength(3);
    expect(instance._presentation).toBeUndefined();
  });
});

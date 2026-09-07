import { beforeEach, describe, expect, it, vi } from 'vitest';
import { buildCfbFixture } from '@silurus/ooxml-core/testing';

const mocks = vi.hoisted(() => ({
  acquire: vi.fn(),
  closeArchive: vi.fn(),
  closeSession: vi.fn(),
  extractImage: vi.fn(() => new Uint8Array([1, 2, 3])),
  openLegacy: vi.fn(),
}));

vi.mock('@silurus/ooxml-pptx/internal/session', async importOriginal => ({
  ...await importOriginal<typeof import('@silurus/ooxml-pptx/internal/session')>(),
  acquirePptxNodeSession: mocks.acquire,
}));
vi.mock('@silurus/ooxml-legacy-converter/internal/direct-ppt-engine', () => ({
  openLegacyPptSource: mocks.openLegacy,
}));

import { openPptxPresentation } from './pptx.ts';

describe('Node native PPT session resource usage', () => {
  beforeEach(() => vi.clearAllMocks());

  it.each(['source', 'lifecycle'] as const)(
    'keeps usage absent and binds %s abort through extraction and cleanup',
    async (abortKind) => {
    const archive = {
      extract_image: mocks.extractImage,
      extract_media: vi.fn(),
      pull_slide: vi.fn(),
      slide_cursor_resource_usage: vi.fn(() => new Uint8Array()),
      acknowledge_slide: vi.fn(),
      cancel_slide: vi.fn(),
      close_presentation_session: mocks.closeSession,
      assert_healthy: vi.fn(),
      free: vi.fn(),
      presentation_bootstrap: vi.fn(() => new TextEncoder().encode(JSON.stringify({
        slideCount: 0,
        slideWidth: 12_192_000,
        slideHeight: 6_858_000,
        defaultTextColor: null,
        majorFont: null,
        minorFont: null,
        hlinkColor: null,
        folHlinkColor: null,
        embeddedFonts: [],
        slides: [],
      }))),
      // Native sources intentionally have no resource_usage method.
    };
    mocks.openLegacy.mockResolvedValueOnce({
      archive,
      sourceByteLength: 1,
      closeArchive: mocks.closeArchive,
    });
    const bytes = buildCfbFixture(['Root Entry', 'PowerPoint Document']);
    const source = {
      protocol: 'ooxml-legacy-ppt-source/v1' as const,
      builtin: 'ppt' as const,
      wasmUrl: 'https://example.test/ppt.wasm',
    };
    const sourceController = new AbortController();
    const lifecycleController = new AbortController();
    const session = await openPptxPresentation(bytes, {
      signal: lifecycleController.signal,
      legacyConversion: { ppt: { source, signal: sourceController.signal } },
    });

    expect(mocks.openLegacy).toHaveBeenCalledOnce();
    expect(mocks.acquire).not.toHaveBeenCalled();

    expect(session.resourceUsage).toBeUndefined();
    await expect(session.getImage('legacy-ppt/image/1', 'image/png'))
      .resolves.toMatchObject({ size: 3, type: 'image/png' });
    expect(session.resourceUsage).toBeUndefined();
    (abortKind === 'source' ? sourceController : lifecycleController).abort();
    await expect(session.getImage('legacy-ppt/image/1', 'image/png'))
      .rejects.toMatchObject({ name: 'AbortError' });

    await session.close();
    expect(mocks.closeSession).toHaveBeenCalledOnce();
    expect(mocks.closeArchive).toHaveBeenCalledOnce();
    },
  );

  it('keeps the selected source signal bound through acquisition and session lifetime', async () => {
    const controller = new AbortController();
    const lifecycle = new AbortController();
    const archive = {
      presentation_bootstrap: vi.fn(() => {
        controller.abort();
        return new TextEncoder().encode(JSON.stringify({
          slideCount: 0,
          slideWidth: 12_192_000,
          slideHeight: 6_858_000,
          defaultTextColor: null,
          majorFont: null,
          minorFont: null,
          hlinkColor: null,
          folHlinkColor: null,
          embeddedFonts: [],
          slides: [],
        }));
      }),
      close_presentation_session: mocks.closeSession,
      cancel_slide: vi.fn(),
      free: vi.fn(),
    };
    mocks.openLegacy.mockResolvedValueOnce({
      archive,
      sourceByteLength: 1,
      closeArchive: mocks.closeArchive,
    });
    const bytes = buildCfbFixture(['Root Entry', 'PowerPoint Document']);
    const source = {
      protocol: 'ooxml-legacy-ppt-source/v1' as const,
      builtin: 'ppt' as const,
      wasmUrl: 'https://example.test/ppt.wasm',
    };
    await expect(openPptxPresentation(bytes, {
      signal: lifecycle.signal,
      legacyConversion: { ppt: { source, signal: controller.signal } },
    })).rejects.toMatchObject({ name: 'AbortError' });
    expect(mocks.closeArchive).toHaveBeenCalledOnce();
  });
});

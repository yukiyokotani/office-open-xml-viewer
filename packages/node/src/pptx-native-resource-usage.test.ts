import { readFile } from 'node:fs/promises';
import { describe, expect, it, vi } from 'vitest';

const mocks = vi.hoisted(() => ({
  acquire: vi.fn(),
  closeArchive: vi.fn(),
  closeSession: vi.fn(),
  extractImage: vi.fn(() => new Uint8Array([1, 2, 3])),
}));

vi.mock('@silurus/ooxml-pptx/internal/session', async importOriginal => ({
  ...await importOriginal<typeof import('@silurus/ooxml-pptx/internal/session')>(),
  acquirePptxNodeSession: mocks.acquire,
}));

import { openPptxPresentation } from './pptx.ts';

describe('Node native PPT session resource usage', () => {
  it('keeps usage absent through image extraction and cleanup when no ZIP metric exists', async () => {
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
      presentation_bootstrap: vi.fn(),
      // Native sources intentionally have no resource_usage method.
    };
    const metrics = {
      observeUsage: vi.fn(),
      fail: vi.fn(),
      checkpoint: vi.fn(),
      succeed: vi.fn(),
    };
    mocks.acquire.mockResolvedValueOnce({
      archive,
      bootstrap: {
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
      },
      metrics,
      closeArchive: mocks.closeArchive,
    });
    const bytes = await readFile(new URL('../../pptx/public/demo/sample-1.pptx', import.meta.url));
    const session = await openPptxPresentation(bytes);

    expect(session.resourceUsage).toBeUndefined();
    await expect(session.getImage('legacy-ppt/image/1', 'image/png'))
      .resolves.toMatchObject({ size: 3, type: 'image/png' });
    expect(session.resourceUsage).toBeUndefined();
    expect(metrics.observeUsage).not.toHaveBeenCalled();

    await session.close();
    expect(mocks.closeSession).toHaveBeenCalledOnce();
    expect(mocks.closeArchive).toHaveBeenCalledOnce();
    expect(metrics.fail).not.toHaveBeenCalled();
  });
});

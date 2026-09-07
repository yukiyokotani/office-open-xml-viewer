import { describe, expect, it, vi } from 'vitest';
import {
  acquirePptxSessionFromArchive,
  type PptxNodeArchive,
} from './node-session-acquisition.js';

const validBootstrap = {
  slideCount: 0,
  slideWidth: 12192000,
  slideHeight: 6858000,
  defaultTextColor: null,
  majorFont: null,
  minorFont: null,
  hlinkColor: null,
  folHlinkColor: null,
  embeddedFonts: [],
  slides: [],
};

function archiveWith(read: () => Uint8Array): PptxNodeArchive {
  return { presentation_bootstrap: read } as unknown as PptxNodeArchive;
}

function encoded(value: unknown): Uint8Array {
  return new TextEncoder().encode(JSON.stringify(value));
}

describe('PPTX owned archive admission', () => {
  it('admits a bootstrap and closes the owned archive exactly once', () => {
    const closeArchive = vi.fn();
    const acquired = acquirePptxSessionFromArchive({
      archive: archiveWith(() => encoded(validBootstrap)),
      sourceByteLength: 123,
      closeArchive,
    });

    expect(acquired.bootstrap.slideCount).toBe(0);
    acquired.closeArchive();
    acquired.closeArchive();
    expect(closeArchive).toHaveBeenCalledTimes(1);
  });

  it('preserves a provider error when cleanup also throws', () => {
    const original = new SyntaxError('bad bootstrap');
    const closeArchive = vi.fn(() => { throw new Error('cleanup failed'); });
    expect(() => acquirePptxSessionFromArchive({
      archive: archiveWith(() => { throw original; }),
      sourceByteLength: 1,
      closeArchive,
    })).toThrow(original);
    expect(closeArchive).toHaveBeenCalledTimes(1);
  });

  it.each([
    new TextEncoder().encode('{'),
    encoded({ ...validBootstrap, slideCount: 1 }),
  ])('closes when returned bootstrap bytes fail decoding or validation', (bytes) => {
    const closeArchive = vi.fn();
    expect(() => acquirePptxSessionFromArchive({
      archive: archiveWith(() => bytes), sourceByteLength: 1, closeArchive,
    })).toThrow();
    expect(closeArchive).toHaveBeenCalledTimes(1);
  });

  it('surfaces successful-session cleanup failure without retrying disposal', () => {
    const failure = new Error('dispose failed');
    const closeArchive = vi.fn(() => { throw failure; });
    const acquired = acquirePptxSessionFromArchive({
      archive: archiveWith(() => encoded(validBootstrap)),
      sourceByteLength: 1,
      closeArchive,
    });
    expect(() => acquired.closeArchive()).toThrow(failure);
    expect(() => acquired.closeArchive()).not.toThrow();
    expect(closeArchive).toHaveBeenCalledTimes(1);
  });

  it('closes without reading when already aborted', () => {
    const controller = new AbortController();
    controller.abort();
    const read = vi.fn(() => encoded(validBootstrap));
    const closeArchive = vi.fn();
    expect(() => acquirePptxSessionFromArchive({
      archive: archiveWith(read), sourceByteLength: 1, closeArchive,
    }, { signal: controller.signal })).toThrowError(expect.objectContaining({ name: 'AbortError' }));
    expect(read).not.toHaveBeenCalled();
    expect(closeArchive).toHaveBeenCalledTimes(1);
  });

  it('closes and aborts when the signal changes during bootstrap projection', () => {
    const controller = new AbortController();
    const closeArchive = vi.fn();
    const archive = archiveWith(() => {
      controller.abort();
      return encoded(validBootstrap);
    });
    expect(() => acquirePptxSessionFromArchive({
      archive, sourceByteLength: 1, closeArchive,
    }, { signal: controller.signal })).toThrowError(expect.objectContaining({ name: 'AbortError' }));
    expect(closeArchive).toHaveBeenCalledTimes(1);
  });

  it('closes on invalid admission options and preserves their validation error', () => {
    const closeArchive = vi.fn(() => { throw new Error('cleanup failed'); });
    expect(() => acquirePptxSessionFromArchive({
      archive: archiveWith(() => encoded(validBootstrap)),
      sourceByteLength: 1,
      closeArchive,
    }, { resourceLimits: { maxArchiveEntryBytes: 0 } })).toThrow(/maxArchiveEntryBytes/);
    expect(closeArchive).toHaveBeenCalledTimes(1);
  });

  it('rejects an invalid source byte count before reading and closes once', () => {
    const read = vi.fn(() => encoded(validBootstrap));
    const closeArchive = vi.fn();
    expect(() => acquirePptxSessionFromArchive({
      archive: archiveWith(read), sourceByteLength: -1, closeArchive,
    })).toThrow(/sourceByteLength/);
    expect(read).not.toHaveBeenCalled();
    expect(closeArchive).toHaveBeenCalledTimes(1);
  });
});

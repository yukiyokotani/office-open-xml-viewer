/** Authored single-stream XLS container helpers; no private or Office-generated data. */
import { buildCfbWithStreams } from '@silurus/ooxml-core/testing';

const NOSTREAM = 0xffffffff;

/**
 * A CFB whose sole `Workbook` stream is the root storage's child. The shared
 * core builder leaves the red-black tree links zeroed, which [MS-CFB] 2.6.1
 * reads as sibling references; the direct reader requires the stream to be
 * owned by the root (Root Entry child = 1, no siblings).
 */
export function workbookCfb(stream: Uint8Array): Uint8Array {
  const bytes = new Uint8Array(buildCfbWithStreams([{ name: 'Workbook', data: stream }]));
  const view = new DataView(bytes.buffer);
  const directory = (view.getUint32(48, true) + 1) * 2 ** view.getUint16(30, true);
  for (const entry of [0, 1]) {
    view.setUint32(directory + entry * 128 + 68, NOSTREAM, true);
    view.setUint32(directory + entry * 128 + 72, NOSTREAM, true);
  }
  view.setUint32(directory + 76, 1, true);
  view.setUint32(directory + 128 + 76, NOSTREAM, true);
  return bytes;
}

import { releaseOwnedBitmap } from '../image/bitmap-image-by-path.js';

/** Transfer one worker-owned bitmap as an atomic ownership handoff. A
 * successful post transfers the surface to the receiver; a synchronous post
 * failure leaves ownership local, so release it before propagating the exact
 * failure to the worker request boundary. */
export function postOwnedImageBitmap<TMessage>(
  post: (message: TMessage, transfer?: Transferable[]) => void,
  message: TMessage & { readonly bitmap: ImageBitmap },
): void {
  const { bitmap } = message;
  try {
    post(message, [bitmap]);
  } catch (error) {
    releaseOwnedBitmap(bitmap);
    throw error;
  }
}

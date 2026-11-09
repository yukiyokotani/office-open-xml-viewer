// Dev server for the ODP corpus harness: the PPTX package's VRT server plus
// the source alias the OpenDocument sources need (as in the root vitest
// configuration), so the harness renders this checkout's sources unbuilt.
import { mergeConfig } from 'vite';
import pptxConfig from '../../../pptx/vite.config.ts';

export default mergeConfig(pptxConfig, {
  resolve: {
    alias: {
      '@silurus/ooxml/internal/companion': new URL('../../../../src/internal-companion.ts', import.meta.url).pathname,
    },
  },
});

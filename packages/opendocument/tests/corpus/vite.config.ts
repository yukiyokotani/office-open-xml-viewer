// Dev server for the ODT corpus harness: the DOCX package's VRT server plus
// the source alias the OpenDocument sources need (as in the root vitest
// configuration), so the harness renders this checkout's sources unbuilt.
import { mergeConfig } from 'vite';
import docxConfig from '../../../docx/vite.config.ts';

export default mergeConfig(docxConfig, {
  resolve: {
    alias: {
      '@silurus/ooxml/internal/companion': new URL('../../../../src/internal-companion.ts', import.meta.url).pathname,
    },
  },
});

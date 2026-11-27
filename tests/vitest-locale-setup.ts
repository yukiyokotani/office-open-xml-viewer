// Unit tests pin the host locale to en-US. Number separators and built-in
// Excel date formats follow navigator.language, which Node derives from the
// environment (LANG / LC_ALL), so an fr or de shell would otherwise change
// every formatted expectation. A test that exercises another locale stubs
// `navigator` itself.
if (typeof navigator !== 'undefined') {
  Object.defineProperty(navigator, 'language', { value: 'en-US', configurable: true });
  Object.defineProperty(navigator, 'languages', { value: Object.freeze(['en-US']), configurable: true });
}

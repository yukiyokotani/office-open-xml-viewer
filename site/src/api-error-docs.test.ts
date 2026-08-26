import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const component = readFileSync(new URL('./components/ApiReference.astro', import.meta.url), 'utf8');
const productionPage = readFileSync(new URL('./pages/production.astro', import.meta.url), 'utf8');
const errorPage = readFileSync(new URL('./pages/errors.astro', import.meta.url), 'utf8');
const readme = readFileSync(new URL('../../README.md', import.meta.url), 'utf8');

describe('public error documentation', () => {
  it('links the shared production guidance and API details to the error reference', () => {
    expect(productionPage).toContain('href="/errors"');
    expect(component).toContain('href={o.detailsHref}');
    expect(errorPage).toContain('id="delivery"');
    expect(errorPage).toContain('id="ooxml-error"');
    expect(errorPage).toContain('id="ooxml-resource-limit-error"');
    expect(errorPage).toContain('id="decoded-image-limit-error"');
    expect(errorPage).toContain('id="parser-crashed"');
    expect(errorPage).toContain('never the message text');
  });

  it('guides a first-time integrator from delivery choice to a safe user response', () => {
    expect(errorPage).toContain('Start with Promise rejection');
    expect(errorPage).toContain('the same failure is never delivered twice');
    expect(errorPage).toContain('Keep <code>onError</code> non-throwing.');
    expect(errorPage).toContain('<code>PptxPresentation.presentSlide()</code>');
    expect(errorPage).toContain('<code>PresentSlideOptions.onError</code>');
    expect(errorPage).toContain('export function previewErrorMessage(error: unknown): string');
    expect(errorPage).toContain('If <code>configurable</code> is true');
    expect(errorPage).toContain('If <code>configurable</code> is false');
    expect(errorPage).toContain('The compressed upload size cannot predict the inflated size.');
    expect(errorPage).toContain('Rust panic, allocation failure, stack overflow');
  });

  it('makes the callback-versus-Promise behavior explicit in the README', () => {
    expect(readme).toContain('## Error handling');
    expect(readme).toContain('Viewer APIs report failures from awaitable operations by rejecting');
    expect(readme).toContain('or not the Viewer has an `onError(error)` callback');
    expect(readme).toContain('through both channels.');
    expect(readme).toContain('messages are diagnostic text, not a programmatic API');
  });
});

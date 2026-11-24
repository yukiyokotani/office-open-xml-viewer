import { llmsTxt, markdownResponse } from '../lib/llms';

export const prerender = true;

export function GET(): Response {
  return markdownResponse(llmsTxt(), 'text/plain');
}

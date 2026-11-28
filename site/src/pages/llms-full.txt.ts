import { llmsFullTxt, markdownResponse } from '../lib/llms';

export const prerender = true;

export function GET(): Response {
  return markdownResponse(llmsFullTxt(), 'text/plain');
}

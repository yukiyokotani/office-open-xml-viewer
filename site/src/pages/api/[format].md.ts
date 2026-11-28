import type { GetStaticPaths } from 'astro';
import { apiMarkdown, llmsFormats, markdownResponse, type LlmsFormat } from '../../lib/llms';

export const prerender = true;

export const getStaticPaths = (() => llmsFormats.map((format) => ({ params: { format } }))) satisfies GetStaticPaths;

export function GET({ params }: { params: { format: LlmsFormat } }): Response {
  return markdownResponse(apiMarkdown(params.format));
}

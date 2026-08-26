import { announcements } from '../lib/announcements';
import { canonicalPageUrl } from '../lib/seo';

export const prerender = true;

export const sitemapPaths = [
  '/',
  '/docx/',
  '/xlsx/',
  '/pptx/',
  '/api/docx/',
  '/api/xlsx/',
  '/api/pptx/',
  '/production/',
  '/try/',
  '/frameworks/',
  '/frameworks/react/',
  '/frameworks/vue/',
  '/frameworks/svelte/',
  '/frameworks/solid/',
  '/announcements/',
  ...announcements.map(({ slug }) => `/announcements/${slug}/`),
  '/bundle-size/',
  '/errors/',
  '/selection-context/',
  '/review-ui/',
  '/review-ui/source/',
] as const;

export function GET(): Response {
  const entries = sitemapPaths
    .map((path) => `  <url><loc>${canonicalPageUrl(path)}</loc></url>`)
    .join('\n');
  const body = [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">',
    entries,
    '</urlset>',
    '',
  ].join('\n');

  return new Response(body, {
    headers: { 'Content-Type': 'application/xml; charset=utf-8' },
  });
}

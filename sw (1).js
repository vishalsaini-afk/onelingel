/* Lingel Field Tools — service worker
   Makes the hub installable and lets every tool work fully offline after the
   first online visit. Strategy: stale-while-revalidate for pages/scripts/styles
   (serve from cache instantly, refresh in the background). The Apps Script
   backend is never cached, so Sheet/Drive calls always hit the live server. */

const CACHE = 'lingel-fieldtools-v1';
const CORE  = ['./'];   // the hub; every other file is cached the first time it's opened

self.addEventListener('install', (e) => {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(CORE).catch(() => {})));
});

self.addEventListener('activate', (e) => {
  e.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;                 // POSTs (Apps Script uploads/logs) go straight to the network

  const url = new URL(req.url);
  // never cache the live backend (Sheet logging, Drive uploads, status lookups)
  if (url.hostname.endsWith('script.google.com') ||
      url.hostname.endsWith('googleusercontent.com') ||
      url.hostname.endsWith('script.googleusercontent.com')) return;

  e.respondWith((async () => {
    const cache  = await caches.open(CACHE);
    const cached = await cache.match(req);
    const network = fetch(req).then((res) => {
      // cache same-origin OK responses and cross-origin (opaque) CDN libs like jsPDF/html2canvas
      if (res && (res.ok || res.type === 'opaque')) cache.put(req, res.clone()).catch(() => {});
      return res;
    }).catch(() => null);
    // serve cache instantly if we have it (and refresh in the background); otherwise wait for network
    return cached || (await network) || cached || new Response('', { status: 504, statusText: 'offline' });
  })());
});

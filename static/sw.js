/* Banki import konvertáló - minimal offline-capable service worker.
 * Strategy: NetworkFirst for HTML/navigations; CacheFirst for static assets.
 * No precache of API responses. Versioned cache name lets bumps purge stale data.
 */
const VERSION = "banki-v2";
const STATIC_CACHE = `${VERSION}-static`;

self.addEventListener("install", event => self.skipWaiting());

self.addEventListener("activate", event => event.waitUntil((async () => {
  const keys = await caches.keys();
  await Promise.all(keys.filter(key => key.startsWith("banki-") && !key.startsWith(VERSION)).map(key => caches.delete(key)));
  await self.clients.claim();
})()));

self.addEventListener("fetch", event => {
  const request = event.request;
  if (request.method !== "GET") return;
  const url = new URL(request.url);
  if (url.origin !== location.origin) return;
  if (url.pathname.startsWith("/api/") || url.pathname === "/healthz") return;
  if (url.pathname.startsWith("/static/") || url.pathname === "/sw.js" || url.pathname === "/manifest.webmanifest") {
    event.respondWith(cacheFirst(request));
    return;
  }
  event.respondWith(networkFirst(request));
});

async function cacheFirst(request) {
  const cache = await caches.open(STATIC_CACHE);
  const hit = await cache.match(request);
  if (hit) return hit;
  const response = await fetch(request);
  if (response.ok) cache.put(request, response.clone());
  return response;
}

async function networkFirst(request) {
  const cache = await caches.open(STATIC_CACHE);
  try {
    const response = await fetch(request);
    if (response.ok) cache.put(request, response.clone());
    return response;
  } catch {
    const hit = await cache.match(request);
    if (hit) return hit;
    return new Response("Offline és nincs gyorsítótár.", { status: 503, headers: { "Content-Type": "text/plain; charset=utf-8" } });
  }
}

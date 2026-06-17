/* Banki TXT konverter - minimal offline-capable service worker.
 * Strategy: NetworkFirst for HTML/navigations; CacheFirst for /static/* assets.
 * No precache of API responses. Versioned cache name lets bumps purge stale data.
 */
const VERSION = "banki-v1";
const STATIC_CACHE = `${VERSION}-static`;

self.addEventListener("install", e => self.skipWaiting());

self.addEventListener("activate", e => e.waitUntil((async () => {
  const keys = await caches.keys();
  await Promise.all(keys.filter(k => k.startsWith("banki-") && !k.startsWith(VERSION)).map(k => caches.delete(k)));
  await self.clients.claim();
})()));

self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (url.origin !== location.origin) return;
  // never cache API or auth-sensitive paths
  if (url.pathname.startsWith("/api/") || url.pathname === "/healthz") return;
  if (url.pathname.startsWith("/static/") || url.pathname === "/sw.js" || url.pathname === "/manifest.webmanifest") {
    e.respondWith(cacheFirst(req));
    return;
  }
  // navigations / HTML
  e.respondWith(networkFirst(req));
});

async function cacheFirst(req) {
  const cache = await caches.open(STATIC_CACHE);
  const hit = await cache.match(req);
  if (hit) return hit;
  const res = await fetch(req);
  if (res.ok) cache.put(req, res.clone());
  return res;
}
async function networkFirst(req) {
  const cache = await caches.open(STATIC_CACHE);
  try {
    const res = await fetch(req);
    if (res.ok) cache.put(req, res.clone());
    return res;
  } catch {
    const hit = await cache.match(req);
    if (hit) return hit;
    return new Response("Offline és nincs gyorsítótár.", { status: 503, headers: { "Content-Type": "text/plain; charset=utf-8" } });
  }
}

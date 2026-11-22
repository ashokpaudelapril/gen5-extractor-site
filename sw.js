// Service worker: makes Gen5 Extractor work offline after the first visit.
//  - this site's own files: network first (so updates show up at once), cache as the fallback
//  - the Python runtime (Pyodide, served from cdn.jsdelivr.net): cache first (it never changes for a given version)
// Your .xpt files are never stored or sent anywhere. Only the app's own code and libraries are cached.
const CACHE = "gen5-extractor-v1";

self.addEventListener("install", () => self.skipWaiting());

self.addEventListener("activate", (event) => {
  event.waitUntil(
    (async () => {
      for (const key of await caches.keys()) if (key !== CACHE) await caches.delete(key);
      await self.clients.claim();
    })()
  );
});

const keep = (cache, req, res) => { cache.put(req, res.clone()).catch(() => {}); };   // some responses (206) cannot be cached

async function networkFirst(req) {
  const cache = await caches.open(CACHE);
  try {
    const res = await fetch(req);
    if (res.ok) keep(cache, req, res);
    return res;
  } catch (err) {
    const hit = (await cache.match(req)) || (await cache.match(req, { ignoreSearch: true }));
    if (hit) return hit;
    if (req.mode === "navigate") {
      const page = (await cache.match("index.html", { ignoreSearch: true })) || (await cache.match("./"));
      if (page) return page;
    }
    throw err;
  }
}

async function cacheFirst(req) {
  const cache = await caches.open(CACHE);
  const hit = await cache.match(req);
  if (hit) return hit;
  const res = await fetch(req);
  if (res.ok || res.type === "opaque") keep(cache, req, res);
  return res;
}

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (url.origin === self.location.origin) {
    if (url.pathname.endsWith("/sw.js")) return;          // never cache the service worker itself
    event.respondWith(networkFirst(req));
  } else if (url.hostname === "cdn.jsdelivr.net") {
    event.respondWith(cacheFirst(req));
  }
});

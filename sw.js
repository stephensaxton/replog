// Rep Log service worker — cache-as-you-go, works with any URL structure
const CACHE = "replog-v2";

// Install immediately — no pre-caching that could fail and kill registration
self.addEventListener("install", () => self.skipWaiting());

// Clean up old cache versions on activate
self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// Cache-first for same-origin GET requests, populate cache on the way through
self.addEventListener("fetch", (e) => {
  if (e.request.method !== "GET" || !e.request.url.startsWith(self.location.origin)) return;
  e.respondWith(
    caches.open(CACHE).then((cache) =>
      cache.match(e.request).then((cached) => {
        const network = fetch(e.request).then((res) => {
          if (res.ok) cache.put(e.request, res.clone());
          return res;
        });
        // Return cached version immediately if available, refresh in background
        return cached || network;
      })
    )
  );
});

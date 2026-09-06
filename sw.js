// Caches the static app shell so it loads with no connection. Deliberately
// leaves the Apps Script API (cross-origin) and third-party CDNs alone —
// app.js's own online/offline handling (see WRITE_QUEUE / BOOTSTRAP_CACHE
// in app.js) depends on those fetches failing honestly when offline.
//
// Network-first, not cache-first: a cache-first shell means a code update
// pushed to GitHub Pages never reaches a device that already has this app
// installed, since every reload keeps re-serving whatever was cached on
// the very first visit, forever. Network-first tries the real network on
// every load (so a fix ships immediately) and only drops back to the
// cached copy when that fetch actually fails (i.e. truly offline).
const CACHE_NAME = "ela-pastry-kitchen-v2";
const SHELL_ASSETS = ["./", "./index.html", "./app.js", "./styles.css", "./config.js"];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);
  if (url.origin !== self.location.origin) return;
  if (event.request.method !== "GET") return;

  event.respondWith(
    fetch(event.request)
      .then((res) => {
        if (res.ok) {
          const copy = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
        }
        return res;
      })
      .catch(() => caches.match(event.request))
  );
});

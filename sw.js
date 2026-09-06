/* Marquart Lake Creel — offline service worker */
const CACHE = "ffsbc-creel-__CACHE__";
const ASSETS = [
  "./",
  "./index.html",
  "./manifest.webmanifest",
  "./icon-192.png",
  "./icon-512.png"
];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

/* Cache-first for app shell so the form opens with no signal at the lake. */
self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET") return;                 // let POST submissions hit the network
  e.respondWith(
    caches.match(req).then(hit => hit || fetch(req).then(res => {
      // runtime-cache same-origin GETs
      const copy = res.clone();
      if (res.ok && new URL(req.url).origin === location.origin) {
        caches.open(CACHE).then(c => c.put(req, copy));
      }
      return res;
    }).catch(() => caches.match("./index.html")))
  );
});

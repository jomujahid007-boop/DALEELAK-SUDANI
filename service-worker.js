/* دليلك — lightweight PWA shell cache. Live listings and Apps Script remain network-only. */
const CACHE_NAME = "daleelak-shell-v10";
const APP_ROOT = new URL("./", self.registration.scope);
const APP_SHELL = [
  APP_ROOT.href,
  new URL("index.html", APP_ROOT).href,
  new URL("admin.html", APP_ROOT).href,
  new URL("manifest.webmanifest", APP_ROOT).href,
  new URL("icon.svg", APP_ROOT).href,
  new URL("daleelak-banner-q84.jpg", APP_ROOT).href
];
self.addEventListener("install", event => {
  event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(APP_SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", event => {
  event.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(key => key.startsWith("daleelak-shell-") && key !== CACHE_NAME).map(key => caches.delete(key)))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", event => {
  const request = event.request;
  if (request.method !== "GET") return;
  const url = new URL(request.url);
  if (url.origin !== self.location.origin || !url.href.startsWith(APP_ROOT.href)) return;
  if (request.mode === "navigate") {
    event.respondWith((async () => {
      try {
        const response = await fetch(request);
        if (response && response.ok) { const cache = await caches.open(CACHE_NAME); await cache.put(request, response.clone()); }
        return response;
      } catch (_) {
        const cached = await caches.match(request);
        if (cached) return cached;
        if (url.pathname === new URL(APP_ROOT).pathname || url.pathname.endsWith("/index.html")) return await caches.match(new URL("index.html", APP_ROOT).href) || Response.error();
        return Response.error();
      }
    })());
    return;
  }
  if (!APP_SHELL.includes(url.href)) return;
  event.respondWith((async () => {
    const cached = await caches.match(request);
    const network = fetch(request).then(async response => { if (response && response.ok) { const cache = await caches.open(CACHE_NAME); await cache.put(request, response.clone()); } return response; }).catch(() => null);
    return cached || await network || Response.error();
  })());
});
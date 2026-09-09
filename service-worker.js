const CACHE_NAME = "ipw-analysis-v22";
const APP_SHELL = [
  "./index.html",
  "./iphone.html",
  "./iphone.css?v=16",
  "./iphone-app.js?v=49",
  "./lot-classification.js?v=1",
  "./foldable-ui.js?v=3",
  "./local-preferences.js?v=1",
  "./analysis.js?v=14",
  "./v90-analysis.js?v=5",
  "./clean-data.js?v=1",
  "./data-management.js?v=4",
  "./vendor/xlsx.full.min.js",
  "./manifest.webmanifest",
  "./icons/ipw-180.png",
  "./icons/ipw-192.png",
  "./icons/ipw-512.png"
];

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL)));
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((key) => key.startsWith("ipw-analysis-") && key !== CACHE_NAME).map((key) => caches.delete(key))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("message", (event) => {
  if (event.data?.type === "SKIP_WAITING") self.skipWaiting();
});

self.addEventListener("fetch", (event) => {
  const request = event.request;
  if (request.method !== "GET") return;
  const url = new URL(request.url);
  if (url.origin !== self.location.origin || isWorkbookRequest(url)) return;

  if (request.mode === "navigate") {
    event.respondWith(fetch(request).catch(async () => (await caches.open(CACHE_NAME)).match("./iphone.html")));
    return;
  }

  event.respondWith(
    caches.open(CACHE_NAME).then((cache) => cache.match(request)).then((cached) => cached || fetch(request).then((response) => {
      if (response.ok) {
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(request, copy));
      }
      return response;
    }))
  );
});

function isWorkbookRequest(url) {
  return /\/data\//i.test(url.pathname) || /\.(?:xlsx|xlsm|xlsb|xls)$/i.test(url.pathname);
}

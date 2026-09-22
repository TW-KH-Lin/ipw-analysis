const CACHE_NAME = "ipw-complaint-workspace-v17-grouped-labels-ui-library-release1-range-stripes";
const APP_SHELL = [
  "../workspace2/label-table-view.js?v=1",
  "../vendor/exceljs.min.js", "./complaint-export.js?v=1",
  '../workspace2/complaint-import-core.js?v=8',
  '../workspace2/complaint-region.js?v=4',
  '../workspace2/complaint-roll-plan.js?v=1',
  '../workspace2/roll-plan-reference.js?v=1',
  './complaint-label-import.js?v=13',
  './report-import/report-parser.js?v=3',
  './report-import/vendor/msgreader.esm.js',
  './report-import/vendor/pdf.min.js',
  './report-import/vendor/pdf.worker.min.js',

  "./roll-plan-reference.js?v=1",
  "./complaint-roll-plan.js?v=1",
  "./complaint-region.js?v=4",
  "./complaint-comparison.js?v=2",
  "./complaint-context.js?v=6", "./complaint-contract.js",
  "../images/correlation-overall-guide.png",
  "../images/correlation-zone-guide.png",
  "./index.html",
  "../workspace/workspace.css?v=2",
  "../workspace/workspace-ui.js?v=4", "./workspace2.css?v=15", "./workspace2-ui.js?v=1",
  "./manifest.webmanifest",
  "../iphone.css?v=27",
  "../iphone-app.js?v=107",
  "../structured-correlation.js?v=10",
  "../lot-classification.js?v=2",
  "../foldable-ui.js?v=6",
  "../local-preferences.js?v=1",
  "../analysis.js?v=17",
  "../v90-analysis.js?v=6",
  "../clean-data.js?v=2",
  "../data-management.js?v=5",
  "../vendor/xlsx.full.min.js",
  "../icons/ipw-180.png",
  "../icons/ipw-192.png",
  "../icons/ipw-512.png"
];

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll([...new Set(APP_SHELL.map(path => new URL(path, self.location.href).href))])));
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((key) => key.startsWith("ipw-complaint-workspace-") && key !== CACHE_NAME).map((key) => caches.delete(key))))
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
    event.respondWith(fetch(request).catch(async () => (await caches.open(CACHE_NAME)).match("./index.html")));
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

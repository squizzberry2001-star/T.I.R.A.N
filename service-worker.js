const APP_VERSION = 'revamp66-home-nav-loading-secret-spark-pdf-spacing';
const CACHE_NAME = `bestie-visit-${APP_VERSION}`;
const LOCAL_ASSETS = [
  './src/theme.css',
  './src/tailwind.generated.css',
  './src/app.js',
  './src/pdf-generator.js',
  './src/pdf-template-assets.js',
  './data.js',
  './store-master-data.js',
  './ca-assignment-export.js',
  './jszip.min.js',
  './netlify-config.js',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png'
];
const CORE_ASSETS = [
  './',
  './index.html',
  './version.json',
  ...LOCAL_ASSETS.map((asset) => `${asset}?v=${APP_VERSION}`)
];

self.addEventListener('install', (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(CORE_ASSETS)).catch(() => undefined));
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((key) => key !== CACHE_NAME && key.startsWith('bestie-visit-')).map((key) => caches.delete(key))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') self.skipWaiting();
});

function isNavigationRequest(request) {
  return request.mode === 'navigate' || (request.headers.get('accept') || '').includes('text/html');
}

function shouldNetworkFirst(request) {
  if (isNavigationRequest(request)) return true;
  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return false;
  if (url.pathname.startsWith('/.netlify/functions/')) return true;
  if (url.pathname.endsWith('/service-worker.js')) return true;
  if (url.pathname.endsWith('/version.json')) return true;
  return ['script', 'style', 'manifest', 'document'].includes(request.destination);
}

function networkFirst(request) {
  return fetch(request, { cache: 'no-store' })
    .then((response) => {
      if (response && response.ok) {
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(request, copy)).catch(() => undefined);
      }
      return response;
    })
    .catch(() => caches.match(request).then((cached) => cached || caches.match('./index.html')));
}

function cacheFirst(request) {
  return caches.match(request).then((cached) => cached || fetch(request).then((response) => {
    const copy = response.clone();
    caches.open(CACHE_NAME).then((cache) => cache.put(request, copy)).catch(() => undefined);
    return response;
  }).catch(() => cached));
}

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;
  event.respondWith(shouldNetworkFirst(event.request) ? networkFirst(event.request) : cacheFirst(event.request));
});

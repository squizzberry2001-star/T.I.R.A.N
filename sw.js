const CACHE_NAME = "zip-preview-v2";
let latestSession = null;

function previewPath(session, path) {
  const clean = path.replace(/^\/+/, "");
  return `/__zip_preview__/${session}/${clean}`;
}

self.addEventListener("install", event => {
  event.waitUntil(self.skipWaiting());
});

self.addEventListener("activate", event => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener("message", event => {
  const data = event.data || {};
  if (data.type === "ZIP_PREVIEW_SESSION") {
    latestSession = data.session || latestSession;
  }
  if (data.type === "ZIP_PREVIEW_CLEAR") {
    event.waitUntil(caches.delete(CACHE_NAME));
  }
});

self.addEventListener("fetch", event => {
  const url = new URL(event.request.url);

  if (url.origin !== self.location.origin) return;

  if (url.pathname.startsWith("/__zip_preview__/")) {
    event.respondWith(serveFromCache(event.request));
    return;
  }

  // Fallback for root-absolute assets inside previewed apps, e.g. /assets/app.js.
  // This keeps Vite/React static builds working when index.html references root assets.
  if (latestSession && event.request.destination !== "document") {
    const mapped = new Request(new URL(previewPath(latestSession, url.pathname), self.location.origin), event.request);
    event.respondWith(serveFromCache(mapped, false));
  }
});

async function serveFromCache(request, includeNotFound = true) {
  const cache = await caches.open(CACHE_NAME);
  const hit = await cache.match(request, { ignoreSearch: true });
  if (hit) return hit;

  if (!includeNotFound) return fetch(request);

  return new Response(
    `Preview file not found: ${new URL(request.url).pathname}`,
    {
      status: 404,
      headers: { "Content-Type": "text/plain; charset=utf-8" }
    }
  );
}
